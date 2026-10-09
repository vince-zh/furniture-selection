const importFields = [
  { key: "room", label: "空间", aliases: ["空间", "房间", "区域", "room"] },
  { key: "position", label: "位置/类型", aliases: ["位置", "类型", "家具类型", "类别", "家具类别", "品类", "position"] },
  { key: "name", label: "名称（可选）", aliases: ["名称", "家具名称", "商品名称", "产品名称", "品名", "name"] },
  { key: "price", label: "总价（可选）", aliases: ["总价", "家具总价", "总金额", "合计", "金额", "价格", "商品价格", "报价", "price", "total"] },
  { key: "qty", label: "数量", aliases: ["数量", "件数", "qty", "quantity"] },
  { key: "url", label: "商品链接", aliases: ["链接", "淘宝链接", "商品链接", "购买链接", "url"] },
  { key: "spec", label: "颜色/材质", aliases: ["规格", "颜色/材质", "颜色", "材质", "spec"] },
  { key: "size", label: "尺寸", aliases: ["尺寸", "size"] },
  { key: "shipping", label: "运费合计", aliases: ["运费", "运费合计", "shipping"] },
  { key: "installation", label: "安装费合计", aliases: ["安装费", "安装费合计", "installation"] },
  { key: "notes", label: "备注", aliases: ["备注", "说明", "notes"] }
];
let importWorkbook = null;
let importRows = [];
let importHeader = 0;
let importMapping = {};
let pendingImport = [];

function importNumber(value, fallback) {
  if (value === "" || value == null) {
    if (fallback !== undefined) return fallback;
    throw new Error("缺少总价");
  }
  const cleaned = String(value).trim().replace(/^[¥￥]\s*/, "").replace(/[,，]/g, "");
  if (!/^(?:\d+(?:\.\d*)?|\.\d+)$/.test(cleaned)) throw new Error("费用或数量不是有效数字");
  const number = Number(cleaned);
  if (!Number.isFinite(number) || number < 0) throw new Error("费用或数量无效");
  return number;
}

function parseImportRows(rows, mapping, header, fallbackSection) {
  const items = [];
  const errors = [];
  if (!Number.isInteger(header) || header < 0 || header >= rows.length) return { items, errors: ["表头行无效"] };
  const columns = Object.values(mapping).filter(value => value != null);
  if (new Set(columns).size !== columns.length) return { items, errors: ["同一列不能对应多个字段"] };
  rows.slice(header + 1).forEach((row, index) => {
    if (!row.some(value => value != null && String(value).trim() !== "")) return;
    const value = key => mapping[key] == null ? "" : row[mapping[key]] ?? "";
    const text = key => String(value(key)).trim();
    try {
      const room = text("room");
      const position = text("position");
      const sectionName = [room, position].filter(Boolean).join(" · ") || fallbackSection;
      if (!text("price")) {
        if (!room && !position) throw new Error("未确定商品的行需填写空间或类型");
        items.push({ sectionName, room, position, pending: true });
        return;
      }
      const price = importNumber(value("price"));
      const qty = importNumber(value("qty"), 1);
      const shipping = importNumber(value("shipping"), 0);
      const installation = importNumber(value("installation"), 0);
      if (qty < 1 || !Number.isInteger(qty)) throw new Error("数量必须为正整数");
      const url = text("url");
      if (url && !/^https?:\/\//i.test(url)) throw new Error("商品链接需以 http:// 或 https:// 开头");
      items.push({ sectionName, room, position, name: text("name") || `方案 ${index + 1}`, price, priceBasis: "total", qty, shipping, installation, url,
        spec: text("spec"), size: text("size"), notes: text("notes"), image: "", scene: "", selected: false });
    } catch (error) { errors.push(`第 ${header + index + 2} 行：${error.message}`); }
  });
  return { items, errors };
}

function setupImportSheet() {
  const sheet = importWorkbook.Sheets[el("excelSheet").value];
  const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
  if (!range || range.e.r > 10000 || range.e.c > 100) throw new Error("工作表为空或超出支持范围（10000 行、100 列）。");
  importRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: true, raw: true, range: 0 });
  // Preserve vertically merged room/category labels without guessing blank cells.
  for (const merge of sheet["!merges"] || []) {
    if (merge.s.c !== merge.e.c) continue;
    const value = importRows[merge.s.r]?.[merge.s.c];
    for (let row = merge.s.r + 1; row <= merge.e.r; row++) if (importRows[row]) importRows[row][merge.s.c] = value;
  }
  importHeader = Number(el("excelHeader").value) - 1;
  const headers = importRows[importHeader] || [];
  importMapping = {};
  const options = headers.map((header, index) => `<option value="${index}">${escapeHtml(XLSX.utils.encode_col(index) + " · " + (header || "未命名列"))}</option>`).join("");
  el("excelMapping").innerHTML = importFields.map(field => {
    const index = headers.findIndex(header => field.aliases.includes(String(header).trim().toLowerCase()));
    importMapping[field.key] = index < 0 ? null : index;
    return `<label>${field.label}<select data-import-field="${field.key}"><option value="">不导入</option>${options}</select></label>`;
  }).join("");
  document.querySelectorAll("[data-import-field]").forEach(select => {
    select.value = importMapping[select.dataset.importField] ?? "";
    select.addEventListener("change", () => {
      importMapping[select.dataset.importField] = select.value === "" ? null : Number(select.value);
      previewImport();
    });
  });
  previewImport();
}

function previewImport() {
  const sheet = importWorkbook.Sheets[el("excelSheet").value];
  const rows = importRows.map(row => [...row]);
  if (importMapping.url != null) rows.forEach((row, index) => {
    const cell = sheet[XLSX.utils.encode_cell({ r: index, c: importMapping.url })];
    if (cell?.l?.Target) row[importMapping.url] = cell.l.Target;
  });
  const { items, errors } = parseImportRows(rows, importMapping, importHeader, activeSection().name);
  pendingImport = items;
  const plannedPositions = new Set(items.filter(item => item.pending).map(item => item.sectionName)).size;
  el("excelSummary").textContent = `${items.filter(item => !item.pending).length} 项候选家具 · ${plannedPositions} 个待选位置 · ${errors.length} 处待修正`;
  el("excelErrors").textContent = errors.slice(0, 20).join("\n") + (errors.length > 20 ? "\n其余问题请修正后重新导入。" : "");
  el("excelApply").disabled = !!errors.length || !items.length;
  el("excelPreview").innerHTML = `<table><thead><tr><th>位置</th><th>方案</th><th>家具总价</th><th>数量</th><th>含费用小计</th></tr></thead><tbody>${items.slice(0, 30).map(item => item.pending
    ? `<tr><td>${escapeHtml(item.sectionName)}</td><td>商品待确定</td><td>—</td><td>—</td><td>—</td></tr>`
    : `<tr><td>${escapeHtml(item.sectionName)}</td><td>${escapeHtml(item.name)}</td><td>${money(item.price)}</td><td>${item.qty}</td><td>${money(itemTotal(item))}</td></tr>`).join("")}</tbody></table>${items.length > 30 ? "<p>预览前 30 项</p>" : ""}`;
}

el("excelBtn").addEventListener("click", () => {
  if (typeof XLSX === "undefined") { showStatus("Excel 读取组件未加载，请确认网页旁的 xlsx.full.min.js 文件完整。"); return; }
  el("excelFile").click();
});
el("excelFile").addEventListener("change", async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    if (file.size > 20 * 1024 * 1024) throw new Error("请选择不超过 20 MB 的表格。");
    importWorkbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    el("excelSheet").innerHTML = importWorkbook.SheetNames.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
    el("excelSheet").value = importWorkbook.SheetNames[0];
    el("excelHeader").value = 1;
    setupImportSheet();
    el("excelDialog").showModal();
  } catch (error) { showStatus(`无法导入：${error.message}`); }
  event.target.value = "";
});
for (const id of ["excelSheet", "excelHeader"]) el(id).addEventListener("change", () => {
  try { setupImportSheet(); }
  catch (error) { pendingImport = []; el("excelErrors").textContent = error.message; el("excelApply").disabled = true; }
});
el("excelCancel").addEventListener("click", () => el("excelDialog").close());
el("excelApply").addEventListener("click", () => {
  if (!pendingImport.length || el("excelApply").disabled) return;
  for (const candidate of pendingImport) {
    let section = state.sections.find(section => section.name === candidate.sectionName);
    if (!section) {
      const location = candidate.room || candidate.position ? { room: candidate.room, position: candidate.position } : sectionLocation(activeSection());
      section = { id: uid("section"), name: candidate.sectionName, ...location, notes: "", image: "", items: [] };
      state.sections.push(section);
    }
    if (candidate.room || candidate.position) {
      section.room = candidate.room;
      section.position = candidate.position;
    }
    if (candidate.pending) continue;
    const { sectionName, room, position, ...item } = candidate;
    section.items.push({ ...item, id: uid("item") });
  }
  pendingImport = [];
  clearForm();
  render();
  el("excelDialog").close();
});
