const LocalImages = (() => {
  let directory;
  let selectedName = "";
  async function restore() {
    if (typeof FurnitureStorage === "undefined") return;
    try {
      const handle = await FurnitureStorage.loadImageDirectory();
      if (handle) { directory = handle; selectedName = handle.name; }
    } catch {}
  }
  async function configure() {
    if (typeof window.showDirectoryPicker !== "function") throw new Error("当前浏览器不支持写入本地文件夹，请使用最新版 Chrome 或 Edge 打开网页。");
    const project = await window.showDirectoryPicker({ id: "furniture-project", mode: "readwrite" });
    directory = await project.getDirectoryHandle("images", { create: true });
    selectedName = `${project.name}/images`;
    if (typeof FurnitureStorage !== "undefined") await FurnitureStorage.saveImageDirectory(directory);
    return selectedName;
  }
  async function authorize() {
    if (!directory) await configure();
    if (typeof directory.queryPermission !== "function") return;
    if (await directory.queryPermission({ mode: "readwrite" }) !== "granted"
      && await directory.requestPermission({ mode: "readwrite" }) !== "granted") throw new Error("需要允许写入图片目录才能保存图片。");
  }
  async function save(image) {
    if (!directory) throw new Error("请先设置图片目录，再粘贴图片。");
    const bytes = await (await fetch(image)).arrayBuffer();
    const digest = await crypto.subtle.digest("SHA-256", bytes);
    const hash = Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, "0")).join("");
    const mime = image.match(/^data:image\/(jpeg|png|webp|gif);/);
    if (!mime) throw new Error("图片格式不支持。");
    const extension = mime[1] === "jpeg" ? "jpg" : mime[1];
    const filename = `furniture-${hash}.${extension}`;
    const handle = await directory.getFileHandle(filename, { create: true });
    const writable = await handle.createWritable();
    try { await writable.write(bytes); await writable.close(); }
    catch (error) { try { await writable.abort(); } catch {} throw error; }
    return filename;
  }
  return { restore, configure, authorize, save, get name() { return selectedName; } };
})();
