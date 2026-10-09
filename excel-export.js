const ExcelProject = (() => {
  const backupSheet = "项目备份";
  const backupFormat = "furniture-project-backup-v1";
  const moneyColumns = [5, 6, 7, 8];

  function build(project) {
    const candidates = [["空间", "类型/位置", "家具名称", "确认状态", "数量", "家具总价", "运费合计", "安装费合计", "总价合计", "颜色/材质", "尺寸", "商品链接", "商品图片数量", "备注"]];
    const quotes = [candidates[0].slice()];
    const summaries = [["空间", "位置数量", "已确认位置", "候选数量", "已确认总价"]];
    const rooms = new Map();
    for (const section of project.sections) {
      const location = sectionLocation(section);
      const room = location.room || "未分类";
      if (!rooms.has(room)) rooms.set(room, [room, 0, 0, 0, 0]);
      const summary = rooms.get(room);
      summary[1]++;
      summary[3] += section.items.length;
      const selected = section.items.find(item => item.selected);
      if (selected) { summary[2]++; summary[4] += Math.round(itemTotal(selected) * 100); }
      if (!section.items.length) candidates.push([room, location.position, "", "待选商品", "", "", "", "", "", "", "", "", 0, section.notes || ""]);
      for (const item of section.items) {
        const row = [room, location.position, item.name, item.selected ? "已确认" : "未确认", Number(item.qty || 1), goodsTotal(item), Number(item.shipping || 0), Number(item.installation || 0), itemTotal(item), item.spec || "", item.size || "", item.url || "", itemImages(item, "image").length, item.notes || ""];
        candidates.push(row);
        if (item === selected) quotes.push(row.slice());
      }
    }
    for (const summary of rooms.values()) summaries.push([...summary.slice(0, 4), summary[4] / 100]);
    const totalCents = [...rooms.values()].reduce((sum, row) => sum + row[4], 0);
    quotes.push(["总计", "", "", "", "", "", "", "", totalCents / 100]);
    summaries.push(["总计", project.sections.length, [...rooms.values()].reduce((sum, row) => sum + row[2], 0), project.sections.reduce((sum, section) => sum + section.items.length, 0), totalCents / 100]);
    const book = XLSX.utils.book_new();
    for (const [name, rows] of [["候选清单", candidates], ["总报价", quotes], ["空间汇总", summaries]]) {
      const sheet = XLSX.utils.aoa_to_sheet(rows);
      sheet["!cols"] = rows[0].map((_, index) => ({ wch: index === 11 ? 42 : index === 13 ? 40 : 18 }));
      sheet["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, rows.length - (name === "候选清单" ? 1 : 2)), c: rows[0].length - 1 } }) };
      for (let row = 1; row < rows.length; row++) {
        for (const column of name === "空间汇总" ? [4] : moneyColumns) {
          const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
          if (cell?.t === "n") cell.z = '"¥"#,##0.00';
        }
        if (name !== "空间汇总" && /^https?:\/\//i.test(rows[row][11] || "")) sheet[XLSX.utils.encode_cell({ r: row, c: 11 })].l = { Target: rows[row][11] };
      }
      XLSX.utils.book_append_sheet(book, sheet, name);
    }
    // Excel cells have a length limit. Split the complete backup so images round-trip.
    const json = JSON.stringify(project);
    const chunks = [[backupFormat], ["序号", "内容"]];
    for (let index = 0; index < json.length;) {
      let end = Math.min(index + 30000, json.length);
      if (end < json.length && /[\uD800-\uDBFF]/.test(json[end - 1])) end--;
      chunks.push([chunks.length - 1, json.slice(index, end)]);
      index = end;
    }
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(chunks), backupSheet);
    book.Workbook = { Sheets: book.SheetNames.map(name => ({ name, Hidden: name === backupSheet ? 1 : 0 })) };
    return book;
  }

  function download(project) {
    const book = build(project);
    const date = new Intl.DateTimeFormat("sv-SE", { timeZone: "Asia/Singapore" }).format(new Date());
    XLSX.writeFile(book, `家具选择与报价-${date}.xlsx`, { compression: true });
  }

  function restore(buffer) {
    const book = XLSX.read(buffer, { type: "array" });
    const sheet = book.Sheets[backupSheet];
    if (!sheet) throw Error("此 Excel 不含完整项目备份，请上传本网页导出的 Excel 或 JSON 备份。");
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
    if (rows[0]?.[0] !== backupFormat) throw Error("Excel 项目备份格式不正确。");
    const json = rows.slice(2).map((row, index) => {
      if (row[0] !== index + 1 || typeof row[1] !== "string") throw Error("Excel 备份不完整。");
      return row[1];
    }).join("");
    return JSON.parse(json);
  }

  return { build, download, restore };
})();
