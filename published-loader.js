const PublishedLoader = (() => {
  async function fetchProject(progress = () => {}) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 90000);
    try {
      progress(null, "正在连接最新发布项目…");
      const response = await fetch(GitHubPublish.site + "published-project.js?latest=" + Date.now(), {cache: "no-store", signal: controller.signal});
      if (!response.ok) throw Error("无法取得最新项目，请稍后重试。");
      const total = Number(response.headers.get("content-length"));
      let source = "";
      if (response.body?.getReader) {
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let received = 0;
        while (true) {
          const {value, done} = await reader.read();
          if (done) break;
          received += value.byteLength;
          if (received > 60 * 1024 * 1024) { await reader.cancel(); throw Error("项目文件过大，请减少图片后重新发布。"); }
          source += decoder.decode(value, {stream: true});
          progress(total > 0 ? Math.min(90, received / total * 90) : null, `正在下载项目 · ${(received / 1048576).toFixed(1)} MB`);
        }
        source += decoder.decode();
      } else source = await response.text();
      progress(92, "正在校验项目…");
      return GitHubPublish.parseSource(source);
    } catch (error) {
      if (error.name === "AbortError") throw Error("加载超时，请检查网络后重试。");
      throw error;
    } finally { clearTimeout(timeout); }
  }
  async function fingerprint(project) {
    const clean = GitHubPublish.cleanProject(project);
    const sections = clean.sections.map(({id, ...section}) => ({...section, items: section.items.map(({id, ...item}) => item)}));
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(JSON.stringify(sections)));
    return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, "0")).join("");
  }
  return {fetchProject, fingerprint};
})();
