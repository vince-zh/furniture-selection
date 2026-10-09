// IndexedDB avoids the small string-storage limit for image-heavy projects.
const FurnitureStorage = (() => {
  let connection;
  function open() {
    if (!connection) connection = new Promise((resolve, reject) => {
      const request = indexedDB.open("furniture-selector", 1);
      request.onupgradeneeded = () => request.result.createObjectStore("projects");
      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(new Error("数据库被其他页面占用，请关闭其他家具网页后重试。"));
      request.onsuccess = () => {
        const database = request.result;
        database.onversionchange = () => { database.close(); connection = null; };
        resolve(database);
      };
    });
    return connection;
  }
  async function transaction(mode, operation) {
    const database = await open();
    return new Promise((resolve, reject) => {
      const tx = database.transaction("projects", mode);
      const request = operation(tx.objectStore("projects"));
      tx.oncomplete = () => resolve(request.result);
      tx.onabort = () => reject(tx.error || request.error || new Error("保存被中断。"));
      tx.onerror = () => reject(tx.error || request.error);
    });
  }
  return {
    load: () => transaction("readonly", store => store.get("current")),
    save: project => transaction("readwrite", store => store.put(project, "current")),
    loadImageDirectory: () => transaction("readonly", store => store.get("image-directory")),
    saveImageDirectory: handle => transaction("readwrite", store => store.put(handle, "image-directory"))
  };
})();
