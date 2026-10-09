const GitHubPublish = (() => {
  const repository = "vince-zh/furniture-selection";
  const branch = "main";
  const file = "published-project.js";
  const api = `https://api.github.com/repos/${repository}`;
  const site = `https://vince-zh.github.io/furniture-selection/`;
  const imagePattern = /^data:image\/(jpeg|png|webp|gif);base64,/;

  function cleanProject(project, revision = project.publishedRevision || "initial") {
    if (!Array.isArray(project.sections) || !project.sections.length) throw Error("项目中没有家具位置。");
    const string = value => String(value || "");
    const pictures = (item, list, single) => (Array.isArray(item[list]) ? item[list] : [item[single]])
      .filter(image => typeof image === "string" && imagePattern.test(image));
    const numeric = (value, fallback = 0) => {
      const result = Number(value ?? fallback);
      if (!Number.isFinite(result) || result < 0) throw Error("项目中的费用或数量无效。");
      return result;
    };
    const sections = project.sections.map(section => ({
      id: string(section.id), name: string(section.name), room: string(section.room),
      position: string(section.position), notes: string(section.notes),
      image: imagePattern.test(section.image || "") ? section.image : "",
      items: section.items.map(item => {
        let url = "";
        if (item.url) {
          const original = new URL(item.url);
          if (!["http:", "https:"].includes(original.protocol)) throw Error("商品链接格式不正确。");
          if (/(^|\.)(taobao|tmall)\.com$/i.test(original.hostname)) {
            const clean = new URL(original.pathname, original.origin);
            for (const key of ["id", "skuId"]) if (original.searchParams.has(key)) clean.searchParams.set(key, original.searchParams.get(key));
            url = clean.href;
          } else url = original.href;
        }
        const qty = numeric(item.qty, 1);
        if (qty < 1 || !Number.isInteger(qty)) throw Error("家具数量必须是正整数。");
        const price = numeric(item.price);
        return {
          id: string(item.id), name: string(item.name), url,
          spec: string(item.spec), size: string(item.size), notes: string(item.notes),
          price: item.priceBasis === "total" ? price : Math.round(price * 100) * qty / 100,
          priceBasis: "total", qty, shipping: numeric(item.shipping), installation: numeric(item.installation),
          selected: Boolean(item.selected), images: pictures(item, "images", "image"),
          sceneImages: pictures(item, "sceneImages", "scene")
        };
      })
    }));
    return { publishedRevision: revision, sections, activeSectionId: string(project.activeSectionId) };
  }

  function encode(text) {
    const bytes = new TextEncoder().encode(text);
    if (bytes.length > 40 * 1024 * 1024) throw Error("当前项目超过 40 MB，请减少图片后发布。");
    let binary = "";
    for (let index = 0; index < bytes.length; index += 32768) binary += String.fromCharCode(...bytes.subarray(index, index + 32768));
    return btoa(binary);
  }

  function decode(text) {
    const binary = atob(text.replace(/\s/g, ""));
    return new TextDecoder().decode(Uint8Array.from(binary, character => character.charCodeAt(0)));
  }

  function parseSource(source) {
    const match = source.match(/^\s*window\.FurnitureInitialProject\s*=\s*([\s\S]*);\s*$/);
    if (!match) throw Error("仓库中的项目文件格式不正确，未执行覆盖。");
    const project = JSON.parse(match[1]);
    if (!Array.isArray(project.sections)) throw Error("仓库中的项目文件格式不正确。");
    return project;
  }

  async function request(path, token, options = {}) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 90000);
    let response;
    try {
      response = await fetch(api + path, {
        ...options, cache: "no-store", redirect: "error", signal: controller.signal,
        headers: {
          Accept: "application/vnd.github.object+json", "X-GitHub-Api-Version": "2026-03-10",
          Authorization: `Bearer ${token}`, ...(options.body ? { "Content-Type": "application/json" } : {})
        }
      });
      if (!response.ok) {
        if (response.status === 401) throw Error("发布凭证无效或已过期。");
        if (response.status === 403 || response.status === 404) throw Error("无法访问仓库，请检查账号权限、凭证授权或 GitHub 请求限额。");
        if (response.status === 409) throw Error("有人刚刚发布了新版，本次未覆盖。请备份修改并载入最新发布项目后再发布。");
        if (response.status === 422) throw Error("GitHub 拒绝此次提交，请检查仓库分支限制或项目大小。");
        throw Error(`GitHub 请求失败（${response.status}）。`);
      }
      return await response.json();
    } catch (error) {
      if (error.name === "AbortError" || error instanceof TypeError) {
        throw Error(options.method === "PUT"
          ? "提交结果尚未确认，请先在 GitHub 仓库查看是否已提交，再决定是否重试。"
          : "无法连接 GitHub，请检查网络后重试。");
      }
      throw error;
    } finally { clearTimeout(timeout); }
  }

  async function publish(project, token, onProgress = () => {}) {
    if (!/^(github_pat_|ghp_)[A-Za-z0-9_]+$/.test(token)) throw Error("请填写有效的 GitHub 发布凭证。");
    onProgress("正在检查仓库版本…");
    const metadata = await request(`/contents/${file}?ref=${branch}`, token);
    if (!/^[a-f0-9]{40}$/.test(metadata.sha)) throw Error("无法确认仓库文件版本，本次未发布。");
    // Read the exact blob checked above, not a possibly cached branch download.
    const blob = await request(`/git/blobs/${metadata.sha}`, token);
    if (blob.encoding !== "base64") throw Error("无法读取仓库项目，本次未发布。");
    const remote = parseSource(decode(blob.content));
    if ((remote.publishedRevision || "initial") !== (project.publishedRevision || "initial")) {
      throw Error("仓库已有其他人发布的新版本。本次未覆盖，请先备份修改，再载入最新发布项目并合并修改。");
    }
    const revision = crypto.randomUUID();
    const snapshot = cleanProject(project, revision);
    const content = encode("window.FurnitureInitialProject = " + JSON.stringify(snapshot) + ";\n");
    onProgress("正在上传家具清单和图片…");
    const result = await request(`/contents/${file}`, token, {
      method: "PUT",
      body: JSON.stringify({ message: "Publish furniture project", branch, sha: metadata.sha, content })
    });
    if (!result.commit?.sha) throw Error("无法确认提交结果，请到 GitHub 仓库核对。");
    return { revision, commit: result.commit.sha, url: `https://github.com/${repository}/commit/${result.commit.sha}` };
  }

  async function waitForDeployment(revision, onProgress = () => {}, { attempts = 32, interval = 15000 } = {}) {
    for (let attempt = 0; attempt < attempts; attempt++) {
      if (attempt) await new Promise(resolve => setTimeout(resolve, interval));
      try {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 10000);
        let deployedRevision;
        try {
          const response = await fetch(`${site}${file}?revision=${encodeURIComponent(revision)}&check=${attempt}`, { cache: "no-store", signal: controller.signal });
          if (response.ok) {
            if (response.body?.getReader) {
              const reader = response.body.getReader();
              let prefix = "";
              const decoder = new TextDecoder();
              while (prefix.length < 200) {
                const chunk = await reader.read();
                if (chunk.done) break;
                prefix += decoder.decode(chunk.value, { stream: true });
              }
              await reader.cancel();
              deployedRevision = prefix.match(/^window\.FurnitureInitialProject = \{"publishedRevision":"([^"]+)"/)?.[1];
            } else deployedRevision = parseSource(await response.text()).publishedRevision;
          }
        }
        finally { clearTimeout(timeout); }
        if (deployedRevision === revision) return true;
      } catch {}
      onProgress("已提交 GitHub，正在等待网站更新…");
    }
    return false;
  }

  function summary(project) {
    let candidates = 0;
    const images = new Set();
    const cleaned = cleanProject(project);
    for (const section of cleaned.sections) {
      if (section.image) images.add(section.image);
      candidates += section.items.length;
      for (const item of section.items) for (const image of [...item.images, ...item.sceneImages]) images.add(image);
    }
    return `${cleaned.sections.length} 个位置 · ${candidates} 个候选 · ${images.size} 张图片`;
  }

  return { publish, waitForDeployment, summary, cleanProject, parseSource, repository, site };
})();
