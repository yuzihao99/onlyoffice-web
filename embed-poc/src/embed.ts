/**
 * **接入页**——住在 iframe 里的那一个。这个 PoC 真正的产出物。
 *
 * 它做的事很少，而且**刻意不知道后端是谁**：
 *
 *   1. 跟宿主说一声「我好了」
 *   2. 收下一条 open：取件地址、存件地址、这是第几版、插件配置
 *   3. 每次要发请求之前，向宿主要一份头
 *   4. 取件 → 交给组件打开 → 人编辑 → 导出 → 存回去
 *
 * 全部约定在 `protocol.ts`，那份文件里没有任何一处提到某个具体的后端。
 * 将来换后端，改的是宿主，不是这一页。
 *
 * ── 为什么只有 hostOrigin 走查询串 ────────────────────────────────────────
 *
 * 因为它是**在能相信任何消息之前**就得知道的东西：不知道该信哪个源，
 * 就没法判断第一条消息是不是伪造的。其余全部经 postMessage
 * ——一个通道比两个少一半出错的地方，而且凭据本来就不能进查询串。
 */

import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  editorManagerFactory,
  registerLegalNotice,
} from "../../src";
import { 按本页位置登记静态资源, 本页挂载前缀 } from "./mount-prefix";
import {
  PROTOCOL,
  isProtocolMessage,
  type EmbedMessage,
  type HeaderPurpose,
  type OpenCommand,
} from "./protocol";

const 状态框 = document.getElementById("状态")!;
function 说(s: string) {
  状态框.textContent = s;
  console.log("[接入页]", s);
}

/**
 * 用 iframe 视口像素强制整条高度链，不依赖 flex/% 解析。
 * DocsAPI 常在挂载后写死偏矮 height；必须用 setProperty(..., important) 压住。
 * 绝不能在改 style 时用 attributes MutationObserver（会死循环卡死）。
 */
let 铺满中 = false;
let 铺满定时 = 0;

function 写尺寸(el: HTMLElement, w: number, h: number) {
  const pxW = w + "px";
  const pxH = h + "px";
  el.style.setProperty("position", "absolute", "important");
  el.style.setProperty("left", "0", "important");
  el.style.setProperty("top", "0", "important");
  el.style.setProperty("right", "0", "important");
  el.style.setProperty("bottom", "0", "important");
  el.style.setProperty("width", pxW, "important");
  el.style.setProperty("height", pxH, "important");
  el.style.setProperty("max-height", "none", "important");
  el.style.setProperty("min-height", "0", "important");
  el.style.setProperty("border", "0", "important");
  el.style.setProperty("box-sizing", "border-box", "important");
  if (el.tagName === "IFRAME") {
    el.setAttribute("width", String(w));
    el.setAttribute("height", String(h));
  }
}

function 铺满编辑器() {
  if (铺满中) return;
  const w = Math.max(1, window.innerWidth || document.documentElement.clientWidth || 0);
  const h = Math.max(1, window.innerHeight || document.documentElement.clientHeight || 0);
  if (w < 2 || h < 2) return;

  铺满中 = true;
  try {
    const pxW = w + "px";
    const pxH = h + "px";
    document.documentElement.style.setProperty("height", pxH, "important");
    document.documentElement.style.setProperty("width", pxW, "important");
    document.body.style.setProperty("height", pxH, "important");
    document.body.style.setProperty("width", pxW, "important");
    document.body.style.setProperty("margin", "0", "important");
    document.body.style.setProperty("overflow", "hidden", "important");

    const stage = document.querySelector(".stage") as HTMLElement | null;
    const box = document.querySelector(".onlyoffice-container") as HTMLElement | null;
    if (stage) 写尺寸(stage, w, h);
    if (box) 写尺寸(box, w, h);

    const root = box || document.body;
    const nodes = root.querySelectorAll<HTMLElement>(
      "iframe, #iframe-office-id, [id^='iframe']",
    );
    nodes.forEach((el) => {
      // 已铺满则跳过，减少无谓写 style
      if (
        el.style.getPropertyValue("width") === pxW &&
        el.style.getPropertyValue("height") === pxH
      ) {
        return;
      }
      写尺寸(el, w, h);
      if (el.tagName === "IFRAME") {
        try {
          (el as HTMLIFrameElement).contentWindow?.dispatchEvent(new Event("resize"));
        } catch {
          /* 跨域编辑器页 */
        }
      }
    });
  } finally {
    铺满中 = false;
  }
}

function 安排铺满() {
  if (铺满定时) {
    window.clearTimeout(铺满定时);
  }
  铺满定时 = window.setTimeout(() => {
    铺满定时 = 0;
    铺满编辑器();
  }, 50);
}

function 监视编辑器尺寸() {
  铺满编辑器();
  if (typeof ResizeObserver !== "undefined") {
    const ro = new ResizeObserver(() => 安排铺满());
    ro.observe(document.documentElement);
    const box = document.querySelector(".onlyoffice-container");
    if (box) ro.observe(box);
  }
  const box = document.querySelector(".onlyoffice-container");
  if (box) {
    // 只盯子树增删，不盯 attributes，避免改 style 自激死循环
    const mo = new MutationObserver(() => 安排铺满());
    mo.observe(box, { childList: true, subtree: true });
  }
  window.addEventListener("resize", 安排铺满);
  // DocsAPI 会在打开后反复改高度：前 30 秒密刷，之后每 2 秒兜底
  let n = 0;
  const fast = window.setInterval(() => {
    铺满编辑器();
    n += 1;
    if (n >= 60) window.clearInterval(fast);
  }, 500);
  window.setInterval(() => 铺满编辑器(), 2000);
}

监视编辑器尺寸();

/**
 * 该信哪个源。
 *
 * ⚠ **取不到就停在这里，不要回落成 `"*"`**。回落之后这一页照样能跑，
 * 只是任何页面都能塞消息进来、也能收走我们发出去的东西——
 * 一个不报错的敞口比一个打不开的页面糟得多。
 */
const 宿主源 = new URL(location.href).searchParams.get("hostOrigin");
if (!宿主源) {
  说("没给 hostOrigin，不知道该信谁 —— 停在这里。");
  throw new Error("missing hostOrigin");
}

/**
 * 告诉法律声明入口「获取源代码」该指哪儿。**在建编辑器之前注册。**
 *
 * ⚠ **这一步在纯静态部署下是必须的，不是可选的。**
 * 不给的话那条链接指的是 `<legalRoot>/source`，而纯静态部署里没有进程去现打源码包
 * ——点开就是 404。许可证第 13 条要的是「凡通过网络与本程序交互的用户都能免费取得
 * **本版本**的完整对应源码」，**入口点开之后 404 与没有入口是一回事**。
 *
 * 而这个 PoC 演的正是纯静态那一档（笨服务器只发文件），所以它必须在这儿给。
 *
 * ⚠ **指的是标签，不是仓库当前状态**（2026-09-05 改）。
 * 许可证要的是「**本版本**」的对应源码，而仓库随时是个更新的版本
 * ——指默认分支的话，用户拿到的是**另一份**代码，而那条链接照样打得开。
 *
 * 标签名由 `package.json` 的 `version` 推出来（构建期经 vite 的 `define` 注入），
 * **不在这里再抄一个版本号**：两处各写一个必然漂，而漂了之后
 * 那条链接指向另一个版本的源码，且照样回 200，没有任何东西会说不对。
 *
 * ⚠ **发布一版的规矩因此是两步，缺一不可**：
 *   ① 改 `package.json` 的 `version`；
 *   ② 提交之后，**在那个提交上打同名标签** `v<version>` 并推上去。
 * 少了 ②，这条链接指向一个不存在的标签 —— GitHub 回 404，
 * 而那正是「入口点开之后 404 与没有入口是一回事」那一条。
 */
registerLegalNotice({
  // ⚠ **这三份原件的地址也要带上挂载前缀。** 默认是站点根下的 `/legal`，
  // 挂在 `/oow` 底下时那三条链接会去站点根要，**点开就是 404**——
  // 而附加条款第三条要的是「用户拿得到许可信息」，**入口点开之后 404 与没有入口是一回事**。
  // 页面本身照常打得开，所以这一条不真去点一下是看不见的。
  legalRoot: 本页挂载前缀() + "/legal",
  sourceUrl:
    "https://github.com/yuhaibohotmail/onlyoffice-web/tree/v" + __OOW_VERSION__,
});

// 按本页所在位置告诉组件去哪儿取静态资源。**必须在开编辑器之前**。
// 为什么是推出来的、为什么末尾要拼 /packages，见 mount-prefix.ts。
按本页位置登记静态资源();

function 发给宿主(m: EmbedMessage) {
  // ⚠ targetOrigin 写死，**永远不要写 "*"**。
  window.parent.postMessage(m, 宿主源);
}

// ── 要一份头：一问一答，每次请求前都问 ────────────────────────────────────

let 请求号 = 0;
const 等着的 = new Map<number, (h: Record<string, string>) => void>();

/**
 * 向宿主要一份请求头。
 *
 * ⚠ **每次发请求之前都要调一次，不要把结果存起来复用。** 存起来就等于回到了
 * 「开场拿一张用到底」，而那张会过期——症状是存件那一下被拒，
 * 而那时用户的成果只在这个页面的内存里。
 */
function 要一份头(purpose: HeaderPurpose): Promise<Record<string, string>> {
  const id = ++请求号;
  return new Promise((resolve, reject) => {
    const 超时 = window.setTimeout(() => {
      等着的.delete(id);
      reject(new Error("宿主没有答复第 " + id + " 次要头（10 秒）"));
    }, 10000);
    等着的.set(id, (h) => {
      window.clearTimeout(超时);
      等着的.delete(id);
      resolve(h);
    });
    发给宿主({ protocol: PROTOCOL, type: "need-headers", requestId: id, purpose });
  });
}

// ── 收消息 ──────────────────────────────────────────────────────────────

let 打开过 = false;
let 当前: OpenCommand | null = null;
let manager: OnlyOfficeManager | null = null;

window.addEventListener("message", (e) => {
  // ⚠ 这一条不能省：不校验来源的话，任何开着这一页的页面都能塞一份自己的头进来。
  if (e.origin !== 宿主源) return;
  if (!isProtocolMessage(e.data)) return;

  const m = e.data as { type: string } & Record<string, unknown>;
  if (m.type === "headers") {
    const cb = 等着的.get(m.requestId as number);
    if (cb) cb((m.headers as Record<string, string>) || {});
    return;
  }
  if (m.type === "open") {
    if (打开过) {
      说("已经打开过一份了，忽略重复的 open。");
      return;
    }
    打开过 = true;
    void 打开(m as unknown as OpenCommand);
    return;
  }
  if (m.type === "save") {
    // 失败会经 failed 消息回给宿主，这里不重复报，也不要让它变成未处理的 rejection。
    void 存件().catch(() => {});
  }
});

// ── 打开 ────────────────────────────────────────────────────────────────

async function 打开(cmd: OpenCommand) {
  当前 = cmd;
  try {
    说("正在取件…");
    const 头 = await 要一份头("download");
    const r = await fetch(cmd.downloadUrl, { headers: 头 });
    if (!r.ok) {
      const body = await r.text().catch(() => "");
      发给宿主({
        protocol: PROTOCOL, type: "failed", stage: "download",
        message: "取件失败 HTTP " + r.status, status: r.status, body: body.slice(0, 500),
      });
      说("取件失败：HTTP " + r.status);
      return;
    }
    const blob = await r.blob();
    // **原始文件必须下到浏览器**——解析它的转换引擎就跑在这里。
    // 这是纯浏览器方案与「后端跑一个文档服务器」最大的差别。
    const file = new File([blob], cmd.fileName, { type: blob.type });

    说("正在打开…（第一次要下不少东西）");
    editorManagerFactory.destroy(ONLYOFFICE_ID);
    const loadSession = editorManagerFactory.beginLoadSession(ONLYOFFICE_ID);
    manager = await OnlyOfficeManager.createWithFile(
      {
        containerId: ONLYOFFICE_ID,
        fileType: cmd.fileType,
        defaultFileName: cmd.fileName,
        lang: (cmd.lang as "zh") || "zh",
        mode: cmd.mode,
        readOnly: cmd.readOnly,
        // 【2026-09-17 新增】不给就是允许，见 protocol.ts。
        allowDownload: cmd.allowDownload,
        zoom: cmd.zoom,
        migrateFitPageZoom: cmd.migrateFitPageZoom,
        suppressLeavePrompt: cmd.suppressLeavePrompt,
        keepViewControls: cmd.keepViewControls,
        loadSession,
        // 原样递过去，一个字段都不看——见 protocol.ts 里那段说明。
        plugins: cmd.plugins as never,
        ...(typeof cmd.officeXmlLimitBytes === "number"
          ? { officeXmlEvent: { isEnable: true, limitBytes: cmd.officeXmlLimitBytes } }
          : {}),
      },
      file,
    );
    说("打开了：" + cmd.fileName + "（基准版 " + String(cmd.baseVersion) + "）");
  } catch (err) {
    发给宿主({
      protocol: PROTOCOL, type: "failed", stage: "open", message: String(err),
    });
    说("打开失败：" + String(err));
  }
}

// ── 存件 ────────────────────────────────────────────────────────────────

async function 存件() {
  if (!manager || !当前) throw new Error("还没打开文档");
  说("正在导出…");
  /**
   * 【2026-09-17 修改】导出抛错要告诉宿主。
   *
   * 原来这一步抛出去之后，被下面收消息那里的 `.catch(() => {})` 吞掉，宿主一个字都收不到，
   * 只能一直停在「正在保存」。`allowDownload: false` 时导出多了一种失败（够不到编辑器、
   * 放不开下载权限），所以这里补上。
   */
  let out: Awaited<ReturnType<OnlyOfficeManager["exportAsBlob"]>>;
  try {
    out = await manager.exportAsBlob();
  } catch (err) {
    发给宿主({ protocol: PROTOCOL, type: "failed", stage: "export", message: "导出失败：" + String(err) });
    说("导出失败：" + String(err));
    throw err;
  }

  /**
   * ⚠ **这一格必须查。**
   *
   * 组件在超过 office-xml 体积上限时，回的**不是这次编辑的结果，
   * 而是打开时那份原文件**，只在返回值里多一格标记。不查就传上去的话，
   * 服务端会显示「保存成功」，而用户这次的改动被悄悄换回了原样
   * ——一个不报错的数据丢失。
   */
  if (out.isOriginalFileFallback) {
    const 说明 =
      "导出回落成了打开时那份原文件（" + (out.fallbackReason || "未知原因") +
      "）。**拒绝上传**：传上去等于把这次的改动丢掉，而界面上会显示保存成功。";
    发给宿主({ protocol: PROTOCOL, type: "failed", stage: "export", message: 说明 });
    说(说明);
    throw new Error(说明);
  }

  说("正在上传 " + out.blob.size + " 字节…");
  const 头 = await 要一份头("save");
  const 地址 = new URL(当前.saveUrl, location.href);
  if (当前.baseVersion !== null && 当前.baseVersion !== undefined) {
    地址.searchParams.set("baseVersion", String(当前.baseVersion));
  }
  const r = await fetch(地址.toString(), {
    method: "POST",
    headers: { ...头, "content-type": "application/octet-stream" },
    body: out.blob,
  });
  const 文本 = await r.text().catch(() => "");
  if (!r.ok) {
    发给宿主({
      protocol: PROTOCOL, type: "failed", stage: "save",
      message: "存件失败 HTTP " + r.status, status: r.status, body: 文本.slice(0, 800),
    });
    说("存件失败：HTTP " + r.status + " " + 文本.slice(0, 200));
    throw new Error("save failed " + r.status);
  }

  let 版本: number | null = null;
  try {
    const j = JSON.parse(文本);
    if (typeof j?.version === "number") 版本 = j.version;
  } catch {
    // 后端不回 JSON 也不算错——版本号只是让宿主能更新基准版，取不到就让它自己去查。
  }
  // 存成功之后基准版要往前走，否则同一次会话里的第二次保存必然撞冲突。
  if (版本 !== null) 当前.baseVersion = 版本;
  发给宿主({ protocol: PROTOCOL, type: "saved", version: 版本, size: out.blob.size });
  说("存回去了：第 " + String(版本) + " 版，" + out.blob.size + " 字节");
  return { version: 版本, size: out.blob.size };
}

/** 页面上那个保存按钮与自动实测走**同一条路**——否则「手点着好使」和「脚本跑绿」会变成两件事。 */
(window as unknown as { __embed: unknown }).__embed = {
  save: () => 存件(),
  state: () => ({ 打开过, 基准版: 当前?.baseVersion ?? null, 有编辑器: !!manager }),
};

发给宿主({ protocol: PROTOCOL, type: "ready" });
说("已就绪，等 open。");
