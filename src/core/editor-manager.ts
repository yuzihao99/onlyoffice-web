import {
  installOnlyOfficeProxies,
  installReporterWindowHook,
  registerScopedIo,
  unregisterScopedIo,
  type OnlyOfficeProxyWindow,
  type ReporterHookWindow,
  type ScopedIoFactory,
} from "../internal/editor/runtime-bridge";
import {
  callCrossOriginEditor,
  canAccessIframeWindow,
  setCrossOriginReadOnly,
  subscribeCrossOriginEditorEvent,
  unregisterCrossOriginBridge,
  watchCrossOriginIframe,
} from "../internal/editor/runtime-bridge";
import {
  CROSS_ORIGIN_EDITOR_COMMAND,
  CROSS_ORIGIN_EDITOR_EVENT,
} from "../internal/editor/runtime-bridge";
import { EditorServer } from "../internal/editor/server";
import io, { type MockSocketOptions } from "../internal/editor/runtime-bridge";
import { EditorLogger } from "../internal/editor/logger";
import {
  type DocEditor,
  DocumentType,
  isOfficeXmlSizeLimitExceededError,
  type OfficeXmlSizeLimitExceededPayload,
  type OfficeTheme,
  type OnlyOfficeConnector,
  type OnlyOfficeConnectorOptions,
  type User,
} from "../internal/editor/types";
import { getDocumentType, isOnlyOfficeCdnMode } from "../const";
import type { AscWordApiCallback, AscWordApiMethod } from "../type/word-api";
import { type OnlyOfficeIframeWindow } from "../type/sdk-internal";
import {
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  ONLYOFFICE_ID,
  ASC_RESTRICTION_NONE,
  ASC_RESTRICTION_VIEW,
  type OfficeXmlEventConfig,
} from "../const";
import {
  type CommentChangeHandlers,
  type CommentData,
  type CommentInput,
  type CommentItem,
  isResolvedComment,
  normalizeCommentInput,
  toPluginCommentPayload,
} from "../feature/comments";
import { mountLegalNotice } from "../legal/notice";
import { onlyofficeEventbus } from "./eventbus";
import {
  type RevisionChangeHandlers,
  type RevisionItem,
  type RevisionsEditorApi,
  collectRevisionItems,
  resolveRevisionShowChanges,
  goToRevision as goToRevisionInSdk,
  applyRevisionChange,
  prepareRevisionReviewDisplay,
} from "../feature/revisions";
import { getOnlyOfficeLang, type OnlyOfficeLang } from "../store/lang";
import { getDocumentObj, setDocumentObj } from "../store/document";
import { initializeOnlyOffice } from "../util/initialize";
import {
  removeOfficeXmlSizeLimitOverlay,
  showOfficeXmlSizeLimitOverlay,
} from "../internal/ui/office-xml-size-limit-overlay";

/** 文档交互模式：view=只读预览，edit=可编辑。 */
export type OnlyOfficeDocumentMode = "view" | "edit";

/**
 * 从 mode / readOnly 解析是否为只读预览。
 * 同时出现时以 mode 为准；都不传则视为可编辑。
 */
export function resolveIsViewMode(options: {
  mode?: OnlyOfficeDocumentMode;
  readOnly?: boolean;
}): boolean {
  if (options.mode === "view") {
    return true;
  }
  if (options.mode === "edit") {
    return false;
  }
  return !!options.readOnly;
}

export type CreateEditorViewOptions = {
  isNew: boolean;
  fileName: string;
  file?: File;
  url?: string;
  loader?: (url: string) => Promise<ArrayBuffer>;
  fileType?: string;
  /**
   * 首选开关：view=只读预览，edit=可编辑。
   * 与 readOnly 同时出现时以本字段为准。
   */
  mode?: OnlyOfficeDocumentMode;
  /** 兼容旧字段；等价于 mode:"view"。与 mode 同时出现时以 mode 为准。 */
  readOnly?: boolean;
  user?: User;
  lang?: string;
  containerId?: string;
  editorManager?: EditorManager;
  editing?: boolean;
  theme?: OfficeTheme;
  /** 由 EditorManagerFactory.beginLoadSession 生成，用于丢弃过期的异步初始化 */
  loadSession?: number;
  /** 修订审阅页：开启 markup 显示与页边修订气泡 */
  revisionReview?: boolean;
  officeXmlEvent?: OfficeXmlEventConfig;
  /** 【本 PoC 新增】插件配置，见 OnlyOfficePluginsConfig */
  plugins?: OnlyOfficePluginsConfig;
  /** 【本项目新增 2026-08-30】编辑器还是查看器，见 OnlyOfficeEditorVariant */
  variant?: OnlyOfficeEditorVariant;
  /** 【本项目新增 2026-09-17】界面上给不给下载 / 另存，默认给。见 grantDownloadForExport */
  allowDownload?: boolean;
  /**
   * OnlyOffice customization.zoom。
   * -2=适应宽度，-1=适合页面，或 50–500 百分比。默认 -2。
   */
  zoom?: number;
  /**
   * localStorage 里若是「适合页面」(-1)，是否迁成当前 zoom。
   * 仅当 zoom===-2 时默认 true；其它 zoom 默认 false。
   */
  migrateFitPageZoom?: boolean;
  /**
   * view 模式下是否抑制 iframe「离开此网站？」提示。默认 true。
   */
  suppressLeavePrompt?: boolean;
  /**
   * view 模式下是否保留缩放/翻页等视图控件（不盖 toolbar-mask）。
   * 默认 true，对 Word / Excel / PPT / PDF 一律生效。
   * 编辑拦截仍靠 asc_setRestriction(VIEW)。
   */
  keepViewControls?: boolean;
};

/**
 * 【本 PoC 新增】传给编辑器的插件配置。
 *
 * 上游这个组件**从来不传这一格**——它建 editorConfig 的时候只放了
 * lang / coEditing / user / customization。于是插件只剩「登记表」一条路
 * （静态根目录下那份 plugins.json），而那条路给不了 `options`。
 *
 * 而 `options` 是**编辑器给插件下发配置的唯一通道**：真实部署里后端
 * 往这里放的是插件访问后端要用的凭证。所以要回答「接真实后端时
 * 插件能不能拿到凭证」，就必须把这一格补上——这也正是本轮的发现之一：
 * **这套纯前端组件要支持插件，得改它自己的源码**（它是 AGPL-3.0）。
 */
export type OnlyOfficePluginsConfig = {
  /** 每一项是一个 config.json 的地址；编辑器直接 fetch 它，并从地址推出插件根目录 */
  pluginsData?: string[];
  /** 打开文档就自动展开的插件 guid */
  autostart?: string[];
  /** 按 guid 分段下发给插件的配置，插件那边读 Asc.plugin.info.options */
  options?: Record<string, unknown>;
};

/**
 * 【本项目新增 2026-08-30】用哪一档界面打开文档。
 *
 * 静态资源里每个编辑器应用都不止一个入口，`api.js` 挑哪个入口只看一格 `config.type`：
 *
 *   const path_type = corrected_type === "mobile" ? "mobile" :
 *                     corrected_type === "embedded" ? (fillForms && isForm ? "forms" : "embed") : "main";
 *
 * 上游这个组件把那一格**写死成 `"desktop"`**（两处），于是永远只加载 `main`——
 * 一个完整编辑器。而资源里一直躺着一个 `embed`，那是真正的查看器应用，
 * 盘上 0.5 MB 对 94.6 MB。这一格就是把那条路开放出来。
 *
 * ⚠ **它与 `readOnly` 是两档，不是一件事**，别当成 `readOnly` 的替代：
 *
 * |                    | 加载哪个应用 | 能编辑 | 有插件面板 | 有工具栏 |
 * |--------------------|------------|-------|-----------|---------|
 * | `readOnly: true`   | main       | 否     | **是**    | 是       |
 * | `variant: "viewer"` | embed      | 否     | **否**    | 几乎没有 |
 *
 * 也就是说：只读是**编辑器关掉编辑**，查看器是**换了另一个应用**。
 * 要插件面板、要工具栏、要将来还能改回可编辑的，用 `readOnly`；
 * 只是把文档给人看的，用 `variant: "viewer"`。两格正交、可以同时设，
 * 但对查看器来说 `readOnly` 不再有额外作用——那个应用本来就没有编辑入口。
 */
export type OnlyOfficeEditorVariant = "editor" | "viewer";

/**
 * `variant` 映射到给编辑器那份配置里的 `type` 一格。
 *
 * ⚠ **两处调用点都要用它**：挂载时一处、`refreshFile` 时一处。
 * 只改挂载那一处的话，一切换只读就会 `refreshFile` 回 `desktop`、跳回编辑器，
 * **而那一跳不报错**。
 */
function editorConfigType(variant: OnlyOfficeEditorVariant | undefined) {
  return variant === "viewer" ? "embedded" : "desktop";
}

export function getFileType(fileName: string, fileType?: string) {
  return fileType || fileName.split(".").pop()?.toLowerCase() || "docx";
}

type OnlyOfficeSdkApi = {
  i1f?: (priority?: number) => void;
  asyncFontsDocumentEndLoaded?: (priority?: number) => void;
  ra?: { Ghj?: () => void };
  asc_registerCallback?: (type: string, fn: AscWordApiCallback) => void;
  asc_unregisterCallback?: (type: string, fn: AscWordApiCallback) => void;
  asc_addComment?: (data: CommentData) => string | undefined;
  asc_changeComment?: (id: string, data: CommentData) => void;
  asc_removeComment?: (id: string) => void;
  sync_ChangeCommentData?: (
    id: string,
    data: unknown,
    ...args: unknown[]
  ) => unknown;
  __ONLYOFFICE_RESOLVE_PATCHED__?: boolean;
  asc_selectComment?: (id: string) => void;
  asc_showComment?: (id: string) => void;
  asc_showComments?: () => void;
  asc_hideComments?: () => void;
  asc_SetGlobalTrackRevisions?: (enabled: boolean) => void;
  asc_GetGlobalTrackRevisions?: () => boolean;
  asc_GetRevisionsChangesStack?: () => unknown[];
  asc_HaveRevisionsChanges?: (all?: boolean) => boolean;
  asc_GetTrackRevisionsReportByAuthors?: () => Record<string, unknown[]>;
  asc_BeginViewModeInReview?: (finalMode?: boolean) => void;
  asc_EndViewModeInReview?: () => void;
  asc_SetLocalTrackRevisions?: (enabled: boolean) => void;
  asc_GetNextRevisionsChange?: () => unknown;
  asc_GetPrevRevisionsChange?: () => unknown;
  asc_FollowRevisionMove?: (change: unknown) => void;
  pluginMethod_MoveToNextReviewChange?: (next: boolean) => void;
  pluginMethod_SetDisplayModeInReview?: (mode: string) => void;
  te?: () => unknown;
  asc_AcceptChanges?: (change?: unknown) => void;
  asc_RejectChanges?: (change?: unknown) => void;
  asc_AcceptChangesBySelection?: (all?: boolean) => void;
  asc_RejectChangesBySelection?: (all?: boolean) => void;
  pluginMethod_GetAllComments?: () => Array<{ Id: string; Data: CommentData }>;
  pluginMethod_AddComment?: (data: CommentData) => string | null;
  pluginMethod_ChangeComment?: (id: string, data: CommentData) => void;
  pluginMethod_InputText?: (text: string) => void;
  pluginMethod_PasteText?: (text: string) => void;
  asc_AddText?: (text: string) => void;
  /** OnlyOffice 内部 WOPI 重命名通道；通过 socket rpc 请求宿主重命名。 */
  asc_wopi_renameFile?: (fileName: string) => void;
};

/** iframe 内运行时；混淆字段见 type/sdk-internal.ts，asc_* 为公开 API */
type OnlyOfficeWindow = OnlyOfficeIframeWindow & {
  Asc?: Omit<NonNullable<OnlyOfficeIframeWindow["Asc"]>, "editor"> & {
    editor?: OnlyOfficeSdkApi & {
      asc_Recalculate?: () => void;
    };
  };
};

type ShellMainController = {
  mode?: { isEdit?: boolean; canEdit?: boolean };
  appOptions?: { isEdit?: boolean; canEdit?: boolean };
  permissions?: { edit?: boolean };
  editorConfig?: { mode?: string };
  api?: {
    asc_setViewMode?: (view: boolean) => void;
    asc_enableKeyEvents?: (enabled: boolean) => void;
  };
  disableEditing?: (disable: boolean) => void;
};

type WordHeaderView = {
  btnDocMode?: { setVisible?: (visible: boolean) => void };
  btnPDFMode?: { setVisible?: (visible: boolean) => void };
};

type ZoomControlLike = {
  setDisabled?: (disabled: boolean) => void;
  isDisabled?: () => boolean;
  keepState?: unknown[];
  disabled?: boolean;
};

type ViewTabControllerLike = {
  SetDisabled?: (state: boolean) => void;
  getView?: () => ViewTabViewLike | undefined;
  view?: ViewTabViewLike;
};

type ViewTabViewLike = {
  lockedControls?: ZoomControlLike[];
  btnsFitToPage?: ZoomControlLike[] | { setDisabled?: (v: boolean) => void };
  btnsFitToWidth?: ZoomControlLike[] | { setDisabled?: (v: boolean) => void };
  btnZoom100?: ZoomControlLike;
  btnMultiplePages?: ZoomControlLike;
  cmbsZoom?: ZoomControlLike[];
};

export class EditorManager {
  private editor: DocEditor | null = null;
  /**
   * Connector 和 logger 一样属于一个 EditorManager（也就是一个编辑器 iframe）。
   * 9.4 的 DocsAPI 每创建一个 Connector 都会注册一组 postMessage 监听器；
   * 因此不能由业务调用方重复创建。
   */
  private connector: OnlyOfficeConnector | null = null;
  private server: EditorServer;
  private dirty = false;
  private readOnly = false;
  private editorLang: OnlyOfficeLang = getOnlyOfficeLang();
  private uiTheme: OfficeTheme = "theme-white";
  /** 【本 PoC 新增】见 OnlyOfficePluginsConfig */
  private pluginsConfig?: OnlyOfficePluginsConfig;
  /** 【本项目新增 2026-08-30】见 OnlyOfficeEditorVariant */
  private variant: OnlyOfficeEditorVariant = "editor";
  /** 【本项目新增 2026-09-17】见 grantDownloadForExport */
  private allowDownload = true;
  /** customization.zoom；默认适应宽度 */
  private zoom = -2;
  private migrateFitPageZoom = true;
  private suppressLeavePrompt = true;
  private keepViewControls = true;
  private viewControlsObserver: MutationObserver | null = null;
  private viewControlsTimer: number | null = null;
  private viewControlsWatchdogStopTimer: number | null = null;
  /** 防止 MutationObserver ↔ DOM 解锁互相触发把 iframe 主线程打满 */
  private viewControlsEnsuring = false;
  private viewControlsUnlocked = false;
  private viewControlsObsTimer: number | null = null;
  /** 与容器一一对应，供事件与 Connector 使用同一稳定路由键。 */
  private instanceId: string;
  private containerId: string;
  private logger: EditorLogger;
  private fileName = "New Document.docx";
  private fileType = "docx";
  private media: Record<string, Uint8Array> = {};
  private comments = new Map<string, CommentData>();
  private crossOriginCommentRefreshPromise: Promise<CommentItem[]> | null =
    null;
  private revisions: RevisionItem[] = [];
  private refreshingRevisions = false;
  private crossOriginRevisionRefreshPromise: Promise<RevisionItem[]> | null =
    null;
  private trackRevisions = false;
  private revisionReviewMode = false;
  private wordContentSyncPromise: Promise<void> | null = null;
  private wordContentSyncTeardown: (() => void) | null = null;
  private crossOriginBridgeTeardown: (() => void) | null = null;
  /** 【本项目新增】法律声明入口的摘除函数。 */
  private legalNoticeTeardown: (() => void) | null = null;

  private officeXmlSizeLimitOverlayTeardown: (() => void) | null = null;
  private officeXmlSizeLimitPayload: OfficeXmlSizeLimitExceededPayload | null =
    null;
  private pendingRename:
    | {
        resolve: (fileName: string) => void;
        reject: (error: Error) => void;
        timer: number;
      }
    | null = null;

  constructor(containerId = ONLYOFFICE_ID) {
    this.containerId = containerId;
    this.instanceId = containerId;
    this.logger = new EditorLogger(containerId);
    this.server = new EditorServer({
      getState: () => ({ readOnly: this.readOnly }),
      logger: this.logger,
      onUserSave: (snapshot) => {
        this.dirty = false;
        this.notifyUserSave(snapshot);
      },
      onLoadError: (error) => {
        this.handleServerLoadError(error);
      },
      onDocumentRename: (fileName) => {
        this.syncRenamedDocument(fileName);
      },
    });
  }

  private getContainerElement() {
    return document.getElementById(this.containerId);
  }

  private getOfficeOverlayHostElement() {
    const container = this.getContainerElement();
    return (
      container?.closest<HTMLElement>(
        ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR,
      ) ?? container
    );
  }

  private clearOfficeXmlSizeLimitOverlay() {
    this.officeXmlSizeLimitOverlayTeardown?.();
    this.officeXmlSizeLimitOverlayTeardown = null;
    this.officeXmlSizeLimitPayload = null;
    const container = this.getOfficeOverlayHostElement();
    if (container) {
      removeOfficeXmlSizeLimitOverlay(container);
    }
  }

  private renderOfficeXmlSizeLimitOverlay() {
    if (!this.officeXmlSizeLimitPayload) {
      return;
    }

    const container = this.getOfficeOverlayHostElement();
    if (container) {
      this.officeXmlSizeLimitOverlayTeardown = showOfficeXmlSizeLimitOverlay(
        container,
        this.officeXmlSizeLimitPayload,
      );
    }
  }

  private handleServerLoadError(error: Error) {
    if (!isOfficeXmlSizeLimitExceededError(error)) {
      return;
    }

    this.officeXmlSizeLimitPayload = error.payload;
    this.renderOfficeXmlSizeLimitOverlay();

    const data = {
      ...error.payload,
      instanceId: this.instanceId,
      containerId: this.containerId,
    };

    try {
      onlyofficeEventbus.emit(
        ONLYOFFICE_EVENT_KEYS.OFFICE_XML_SIZE_LIMIT_EXCEEDED,
        data,
      );
    } catch (eventError) {
      console.error(
        "[EditorManager] officeXmlSizeLimitExceeded handler failed",
        eventError,
      );
    }
  }

  private createScopedIo() {
    return (url?: string, options: MockSocketOptions = {}) => {
      const socket = io(url, { ...options, logger: this.logger });

      socket.on("connect", () => {
        this.server.handleConnect({ socket });
      });
      socket.on("disconnect", () => {
        this.server.handleDisconnect({ socket });
      });

      return socket;
    };
  }

  private createCrossOriginServerIo(): ScopedIoFactory {
    return () => {
      const socket = io(undefined, {
        deferConnect: true,
        logger: this.logger,
      });
      socket.connected = true;
      socket.disconnected = false;
      return socket;
    };
  }

  private syncEditorBridge() {
    this.crossOriginBridgeTeardown?.();
    this.crossOriginBridgeTeardown = null;
    unregisterScopedIo(this.containerId);

    registerScopedIo(this.containerId, this.createScopedIo());
    if (!isOnlyOfficeCdnMode()) {
      return;
    }

    this.crossOriginBridgeTeardown = watchCrossOriginIframe(
      this.containerId,
      () => this.getEditorFrameElement(),
      this.server,
      this.createCrossOriginServerIo(),
    );
  }

  private syncCrossOriginReadOnly(readOnly: boolean, retries = 10) {
    if (setCrossOriginReadOnly(this.containerId, readOnly)) {
      return true;
    }

    if (retries > 0) {
      window.setTimeout(() => {
        this.syncCrossOriginReadOnly(readOnly, retries - 1);
      }, 50);
    }

    return false;
  }

  private getEditorFrameElement() {
    const containerFrame = document
      .getElementById(this.containerId)
      ?.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');

    if (containerFrame) {
      return containerFrame;
    }

    const frames = Array.from(
      document.querySelectorAll<HTMLIFrameElement>(
        'iframe[name="frameEditor"]',
      ),
    );
    const matchedFrame = frames.find((frame) => {
      try {
        const url = new URL(frame.src, window.location.origin);
        return url.searchParams.get("frameEditorId") === this.containerId;
      } catch {
        return false;
      }
    });

    if (matchedFrame) {
      return matchedFrame;
    }

    if (this.containerId === ONLYOFFICE_ID) {
      return frames[0];
    }

    return document
      .querySelector<HTMLElement>(
        `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"]`,
      )
      ?.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
  }

  private installProxiesOnWindow(win: OnlyOfficeProxyWindow) {
    // 9.4 的协作客户端会把 socket.io 请求相对解析到静态 SDK 根目录。
    // 必须替换 iframe 内的 io，才能将 /doc/... 连接交给内存 EditorServer。
    installOnlyOfficeProxies(win, this.server, this.createScopedIo());
  }

  /**
   * 劫持 iframe 内 XHR/fetch/io，将协作与 downloadAs 请求路由到 mock EditorServer。
   * 必须在 downloadAs 前安装，否则 export 无法收到 /downloadas/ 分片。
   */
  private installIframeProxies() {
    const iframe = this.getEditorFrameElement();
    if (!iframe) {
      throw new Error("Iframe not loaded");
    }

    if (!canAccessIframeWindow(iframe)) {
      return;
    }

    const win = iframe.contentWindow as
      (OnlyOfficeWindow & ReporterHookWindow) | undefined;
    const iframeDoc = iframe.contentDocument;

    if (!iframeDoc || !win) {
      throw new Error("Iframe not loaded");
    }

    if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
      return;
    }

    this.installProxiesOnWindow(win);
    installReporterWindowHook(win, (target) => {
      this.installProxiesOnWindow(target as OnlyOfficeProxyWindow);
    });
    this.installSaveShortcutBlocker();
  }

  private getEditorFrameWindow() {
    const iframe = this.getEditorFrameElement();

    if (!iframe || !canAccessIframeWindow(iframe)) {
      return undefined;
    }

    return iframe.contentWindow as OnlyOfficeWindow | undefined;
  }

  private getSdkApi() {
    return this.getEditorFrameWindow()?.Asc?.editor;
  }

  private requireSdkApi() {
    const api = this.getSdkApi();

    if (!api) {
      throw new Error("OnlyOffice SDK API is not ready");
    }

    return api;
  }

  private installCommentResolveCleanup() {
    const api = this.getSdkApi();

    if (!api || api.__ONLYOFFICE_RESOLVE_PATCHED__) {
      return;
    }

    api.__ONLYOFFICE_RESOLVE_PATCHED__ = true;

    const removeResolvedComment = (id: unknown) => {
      // OnlyOffice resolves comments through an internal change event first.
      // Removing synchronously during that event can race its own render pass,
      // so schedule the delete for the next tick after the resolved state lands.
      window.setTimeout(() => {
        api.asc_removeComment?.(String(id));
        this.comments.delete(String(id));
      }, 0);
    };

    const originalSyncChange = api.sync_ChangeCommentData?.bind(api);
    if (originalSyncChange) {
      api.sync_ChangeCommentData = (id, data, ...args) => {
        const result = originalSyncChange(id, data, ...args);

        if (isResolvedComment(data)) {
          removeResolvedComment(id);
        }

        return result;
      };
    }

    api.asc_registerCallback?.("asc_onChangeCommentData", (id, data) => {
      if (!isResolvedComment(data)) {
        return;
      }

      removeResolvedComment(id);
    });
  }

  private getDocumentPermissions(editing: boolean) {
    const doc = this.server.getDocument();
    return {
      edit: editing && doc.fileType !== "pdf",
      chat: false,
      rename: editing,
      protect: editing,
      // 允许接受/拒绝文档内已有修订；不自动进入「修订」录制模式
      review: true,
      print: false,
      // 【本项目新增 2026-09-17】为假时编辑器按「不许下载」加载：文件菜单里没有「下载为」
      // 和另存面板。组件自己导出时要临时放开，见 grantDownloadForExport。
      download: this.allowDownload,
    };
  }

  /**
   * 预览默认「适应宽度」。OnlyOffice 会把用户选过的缩放写进 localStorage，
   * 并覆盖 customization.zoom；若以前卡在「适合页面」(-1)，百分比下拉几乎用不了。
   * 开启 migrateFitPageZoom 时，把 -1 迁成当前 zoom（通常 -2）。
   * 编辑器在 iframe 内读自己的 localStorage，父页与 iframe 都迁一次。
   */
  private preferFitWidthZoomPreference(storage?: Storage | null) {
    if (!this.migrateFitPageZoom) {
      return;
    }
    const target = String(this.zoom);
    const keys = [
      "de-settings-zoom",
      "pdfe-settings-zoom",
      "pe-settings-zoom",
      "sse-settings-zoom",
    ];
    const targets = [storage, window.localStorage].filter(
      (item): item is Storage => !!item,
    );
    for (const store of targets) {
      try {
        for (const key of keys) {
          if (store.getItem(key) === "-1") {
            store.setItem(key, target);
          }
        }
      } catch {
        // ignore quota / private mode
      }
    }
  }

  /** 关闭 autosave 与保存按钮；保存快捷键由 installSaveShortcutBlocker 拦截。 */
  private buildEditorCustomization() {
    return {
      uiTheme: this.uiTheme,
      autosave: false,
      // -2 = 适应宽度；-1 = 适合页面（预览框偏矮时会被缩得很小，缩放下拉也不好用）
      zoom: this.zoom,
      layout: {
        header: {
          save: false,
          // Word 头部「编辑 / 审阅 / 查看」切换（PPT/Excel 无此入口）
          editMode: false,
        },
        toolbar: {
          file: {
            save: false,
          },
          save: false,
        },
      },
      review: {
        trackChanges: this.revisionReviewMode,
        // showReviewChanges:true 会在加载时弹出 asc-window review-changes modal-dlg
        showReviewChanges: false,
        ...(this.revisionReviewMode
          ? { reviewDisplay: "markup" as const }
          : {}),
      },
      features: {
        spellcheck: {
          change: false,
        },
      },
      // 【本项目修改 2026-08-30】这里原来有一个 logo 块，把编辑器左上角换成了
      // 微软 Office 的图标（从 jsDelivr 上取 simple-icons 的 microsoftoffice.svg）。已删掉，
      // 现在显示的是编辑器自带的 ONLYOFFICE 标识。三个理由：
      //   一、那是别家的商标，我们没有任何权利用它；
      //   二、换掉 ONLYOFFICE 自己的标识就是白标，而白标是商业版才有的功能，
      //       AGPL 这条路上不允许；
      //   三、许可要求「能认出 ONLYOFFICE 是原始开发者」，那个标识正是最直接的一处。
      // 顺带去掉了一个外部地址依赖——断网的环境里原来那个图标是取不到的。
    };
  }

  /** 禁用 Ctrl/Cmd+S 与工具栏保存，避免与 export/downloadAs 共用管道冲突。 */
  private installSaveShortcutBlocker() {
    const win = this.getEditorFrameWindow();
    const doc = win?.document;

    if (!doc || win?.__ONLYOFFICE_SAVE_BLOCKED__) {
      return;
    }

    const blockSaveShortcut = (event: KeyboardEvent) => {
      const key = event.key.toLowerCase();
      if ((event.ctrlKey || event.metaKey) && key === "s") {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
      }
    };

    doc.addEventListener("keydown", blockSaveShortcut, true);
    win.__ONLYOFFICE_SAVE_BLOCKED__ = true;
  }

  /** 文档若自带 w:trackRevisions，OnlyOffice 默认会跟进入修订模式；接入层强制关闭录制。 */
  private applyDefaultReviewSettings() {
    this.trackRevisions = false;
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SET_TRACK,
        {
          enabled: false,
        },
      ).catch(() => {});
      return;
    }

    const api = this.getSdkApi();
    api?.asc_SetGlobalTrackRevisions?.(false);
  }

  private mergeCommentItems(items: CommentItem[]) {
    for (const item of items) {
      if (isResolvedComment(item.Data)) {
        this.comments.delete(item.Id);
        continue;
      }

      this.comments.set(item.Id, item.Data);
    }
  }

  private fetchCommentsFromSdk(): CommentItem[] {
    const raw = this.getSdkApi()?.pluginMethod_GetAllComments?.();
    if (!Array.isArray(raw)) {
      return [];
    }

    return raw
      .map((item, index) => {
        const source =
          item && typeof item === "object"
            ? (item as Record<string, unknown>)
            : {};
        const id = String(source.Id ?? source.id ?? `comment-${index}`);
        const data = (source.Data ?? source.data ?? source) as CommentData;

        return { Id: id, Data: data };
      })
      .filter((item) => !isResolvedComment(item.Data));
  }

  private refreshCommentsFromSdk() {
    if (isOnlyOfficeCdnMode()) {
      void this.refreshCrossOriginComments().catch(() => {});
      return;
    }

    this.mergeCommentItems(this.fetchCommentsFromSdk());
  }

  private normalizeCrossOriginComments(value: unknown): CommentItem[] {
    if (!Array.isArray(value)) {
      return [];
    }

    return value.map((item, index) => {
      const source =
        item && typeof item === "object"
          ? (item as Record<string, unknown>)
          : {};
      const data =
        source.Data && typeof source.Data === "object"
          ? (source.Data as CommentData)
          : {};

      return {
        Id: String(source.Id ?? `comment-${index}`),
        Data: data,
      };
    });
  }

  private refreshCrossOriginComments() {
    if (this.crossOriginCommentRefreshPromise) {
      return this.crossOriginCommentRefreshPromise;
    }

    this.crossOriginCommentRefreshPromise = this.callCrossOriginComment(
      CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_LIST,
      {},
    )
      .then((items) => {
        const normalized = this.normalizeCrossOriginComments(items);
        this.comments.clear();
        this.mergeCommentItems(normalized);
        return Array.from(this.comments.entries()).map(([Id, Data]) => ({
          Id,
          Data,
        }));
      })
      .finally(() => {
        this.crossOriginCommentRefreshPromise = null;
      });

    return this.crossOriginCommentRefreshPromise;
  }

  private refreshRevisionsFromSdk(options?: { forceRefreshStack?: boolean }) {
    if (isOnlyOfficeCdnMode()) {
      void this.refreshCrossOriginRevisions(options).catch(() => {});
      return;
    }

    if (this.refreshingRevisions) return;

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      this.revisions = [];
      return;
    }

    this.refreshingRevisions = true;
    try {
      this.revisions = collectRevisionItems(
        api as RevisionsEditorApi,
        frameWin,
        options,
      );
    } finally {
      this.refreshingRevisions = false;
    }
  }

  private applyRevisionsFromSdkStack(stack: unknown) {
    if (isOnlyOfficeCdnMode()) {
      this.revisions = this.normalizeCrossOriginRevisions(stack);
      return;
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) return;

    this.revisions = resolveRevisionShowChanges(
      stack,
      api as RevisionsEditorApi,
      frameWin,
    );
  }

  private syncRevisionsAfterMutation() {
    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
  }

  private applyAllRevisionChanges(mode: "accept" | "reject") {
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        mode === "accept"
          ? CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT_ALL
          : CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT_ALL,
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
      return;
    }

    const api = this.requireSdkApi() as RevisionsEditorApi;
    const frameWin = this.getEditorFrameWindow();
    if (!frameWin) {
      return;
    }

    const applyBulk =
      mode === "accept"
        ? () => api.asc_AcceptChanges?.()
        : () => api.asc_RejectChanges?.();

    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
    let guard = 0;
    let stagnant = 0;
    let lastCount = this.revisions.length;

    while (this.haveRevisionsChanges() && guard++ < 20) {
      this.refreshRevisionsFromSdk({ forceRefreshStack: true });
      const [first] = this.revisions;

      if (!first) {
        applyBulk();
        this.syncRevisionsAfterMutation();
        continue;
      }

      applyRevisionChange(mode, first, api, frameWin, this.revisions);
      this.syncRevisionsAfterMutation();

      const nextCount = this.revisions.length;
      if (nextCount >= lastCount && this.haveRevisionsChanges()) {
        stagnant += 1;
        if (stagnant >= 3) {
          applyBulk();
          this.syncRevisionsAfterMutation();
          stagnant = 0;
        }
      } else {
        stagnant = 0;
      }

      lastCount = nextCount;
    }

    this.syncRevisionsAfterMutation();
  }

  private normalizeCrossOriginRevisions(value: unknown): RevisionItem[] {
    if (!Array.isArray(value)) {
      return [];
    }

    return value.map((item, index) => {
      const source =
        item && typeof item === "object"
          ? (item as Record<string, unknown>)
          : {};
      const data =
        source.Data && typeof source.Data === "object"
          ? (source.Data as RevisionItem["Data"])
          : {};

      return {
        Id: String(source.Id ?? `rev-stack-${index}`),
        Index:
          typeof source.Index === "number" && Number.isFinite(source.Index)
            ? source.Index
            : index,
        Data: data,
        Raw: (source.Raw ?? {}) as RevisionItem["Raw"],
      };
    });
  }

  private refreshCrossOriginRevisions(options?: {
    forceRefreshStack?: boolean;
  }) {
    if (this.crossOriginRevisionRefreshPromise) {
      return this.crossOriginRevisionRefreshPromise;
    }

    this.crossOriginRevisionRefreshPromise = this.callCrossOriginRevision(
      CROSS_ORIGIN_EDITOR_COMMAND.REVISION_LIST,
      {
        forceRefreshStack: !!options?.forceRefreshStack,
      },
    )
      .then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
        return this.revisions;
      })
      .finally(() => {
        this.crossOriginRevisionRefreshPromise = null;
      });

    return this.crossOriginRevisionRefreshPromise;
  }

  private callCrossOriginRevision(
    command: string,
    payload: Record<string, unknown> = {},
  ) {
    return callCrossOriginEditor(this.containerId, command, payload);
  }

  private revisionTargetId(revision: RevisionItem | string) {
    return typeof revision === "string" ? revision : revision.Id;
  }

  addDemoRevision(text = `审批修订 ${new Date().toLocaleTimeString()}`) {
    this.trackRevisions = true;
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ADD_DEMO,
        { text },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
        return this.revisions;
      });
    }

    const api = this.requireSdkApi() as OnlyOfficeSdkApi;
    api.asc_SetGlobalTrackRevisions?.(true);
    api.asc_SetLocalTrackRevisions?.(true);
    if (api.pluginMethod_InputText) {
      api.pluginMethod_InputText(text);
    } else if (api.pluginMethod_PasteText) {
      api.pluginMethod_PasteText(text);
    } else if (api.asc_AddText) {
      api.asc_AddText(text);
    } else {
      throw new Error("OnlyOffice text insertion API is not available");
    }
    this.syncRevisionsAfterMutation();
    return this.revisions;
  }

  private teardownWordContentSync() {
    this.wordContentSyncTeardown?.();
    this.wordContentSyncTeardown = null;
    this.wordContentSyncPromise = null;
  }

  private scheduleWordContentSync() {
    window.setTimeout(() => {
      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();
    }, 0);
  }

  private ensureWordContentSync() {
    if (this.fileType !== "docx" && getDocumentType(this.fileType) !== "word") {
      return Promise.resolve();
    }

    if (this.wordContentSyncPromise) {
      return this.wordContentSyncPromise;
    }

    this.wordContentSyncPromise = (async () => {
      if (isOnlyOfficeCdnMode()) {
        await Promise.all([
          this.refreshCrossOriginComments(),
          this.refreshCrossOriginRevisions(),
          this.callCrossOriginComment(
            CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
            {},
          ),
          this.callCrossOriginRevision(
            CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE,
          ),
        ]);

        const unsubscribers = [
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                return;
              }

              this.comments.set(commentId, commentData);
            },
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                return;
              }

              this.comments.set(commentId, commentData);
            },
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT,
            ([id]) => {
              this.comments.delete(String(id));
            },
          ),
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            ([items]) => {
              this.revisions = this.normalizeCrossOriginRevisions(items);
            },
          ),
        ];

        this.wordContentSyncTeardown = () => {
          unsubscribers.forEach((unsubscribe) => unsubscribe());
        };
        return;
      }

      const api = this.requireSdkApi();

      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();

      const unsubscribers = await Promise.all([
        this.subscribe({
          type: "asc_onAddComment",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data as CommentData;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }

            this.comments.set(commentId, commentData);
          },
        }),
        this.subscribe({
          type: "asc_onChangeCommentData",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data as CommentData;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }

            this.comments.set(commentId, commentData);
          },
        }),
        this.subscribe({
          type: "asc_onRemoveComment",
          fn: (id) => {
            this.comments.delete(String(id));
          },
        }),
        this.subscribe({
          type: CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
          fn: (stack) => {
            this.applyRevisionsFromSdkStack(stack);
          },
        }),
      ]);

      this.wordContentSyncTeardown = () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe?.());
      };

      this.scheduleWordContentSync();
      api.asc_Recalculate?.();
    })().catch(() => {
      this.teardownWordContentSync();
    });

    return this.wordContentSyncPromise;
  }

  private getRestrictionSdkApi() {
    return this.getSdkApi() as
      | {
          asc_setRestriction?: (type: number) => void;
          asc_removeRestriction?: (type: number) => void;
          asc_setCanSendChanges?: (enabled: boolean) => void;
        }
      | undefined;
  }

  private getShellMainController() {
    const win = this.getEditorFrameWindow() as OnlyOfficeIframeWindow & {
      SSE?: { getController?: (name: string) => ShellMainController };
      PE?: { getController?: (name: string) => ShellMainController };
      DE?: { getController?: (name: string) => ShellMainController };
      getApplication?: () => {
        getController?: (name: string) => ShellMainController;
      };
    };

    return (
      win?.SSE?.getController?.("Main") ??
      win?.PE?.getController?.("Main") ??
      win?.getApplication?.()?.getController?.("Main") ??
      win?.DE?.getController?.("Main")
    );
  }

  private getWordHeaderView() {
    const win = this.getEditorFrameWindow() as OnlyOfficeIframeWindow & {
      DE?: {
        getController?: (name: string) => {
          getView?: (name: string) => WordHeaderView;
        };
      };
    };

    return win?.DE?.getController?.("Viewport")?.getView?.(
      "Common.Views.Header",
    );
  }

  /** 隐藏 Word 头部「编辑 / 审阅 / 查看」切换（customization + 运行时兜底）。 */
  private hideWordDocModeSwitcher() {
    if (getDocumentType(this.fileType) !== DocumentType.Word) {
      return;
    }

    const header = this.getWordHeaderView();
    header?.btnDocMode?.setVisible?.(false);
    header?.btnPDFMode?.setVisible?.(false);

    const hedset = this.getEditorFrameWindow()?.document?.querySelector(
      '[data-layout-name="header-editMode"]',
    );
    if (hedset instanceof HTMLElement) {
      hedset.style.display = "none";
    }
  }

  private scheduleWordDocModeHide() {
    this.hideWordDocModeSwitcher();
    window.setTimeout(() => this.hideWordDocModeSwitcher(), 0);
  }

  /** 同步 web-apps 工具栏/侧栏的 editing:disable（viewMode 与只读一致）。 */
  private syncShellEditingDisable(
    disabled: boolean,
    documentType = getDocumentType(this.fileType),
  ) {
    const nc = this.getEditorFrameWindow()?.Common?.NotificationCenter;
    if (!nc?.trigger) {
      return;
    }

    // DisableToolbar(true) 会盖 .toolbar-mask，缩放/翻页一起置灰。
    // keepViewControls 时不盖 mask，编辑拦截靠 asc_setRestriction(VIEW)。
    const maskToolbar = disabled ? !this.keepViewControls : true;
    const disableStatusBar = disabled && this.keepViewControls ? false : true;

    if (documentType === DocumentType.Slide) {
      nc.trigger("editing:disable", disabled, {
        viewMode: disabled,
        allowSignature: !disabled,
        rightMenu: { clear: false, disable: true },
        statusBar: disableStatusBar,
        leftMenu: { disable: disabled, previewMode: disabled },
        fileMenu: false,
        comments: { disable: false, previewMode: disabled },
        chat: false,
        review: true,
        viewport: disabled && !this.keepViewControls,
        documentHolder: { clear: disabled, disable: true },
        toolbar: maskToolbar,
        header: { search: false },
        shortcuts: disabled ? false : undefined,
      });
      if (disabled && this.keepViewControls) {
        this.ensureViewControlsUsable();
      }
      return;
    }

    if (documentType === DocumentType.Word) {
      if (disabled) {
        nc.trigger("editing:disable", true, {
          viewMode: true,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: false,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: true, disable: true },
          statusBar: disableStatusBar,
          leftMenu: { disable: true, previewMode: true },
          fileMenu: { protect: true, history: false },
          navigation: {
            disable: false,
            previewMode: !this.keepViewControls,
          },
          comments: { disable: false, previewMode: true },
          chat: false,
          review: true,
          // viewport:true 会连带影响视图区；保留缩放时不要锁
          viewport: !this.keepViewControls,
          documentHolder: { clear: true, disable: true },
          toolbar: maskToolbar,
          protect: true,
          header: { search: false, startfill: false },
          shortcuts: false,
        });
        if (this.keepViewControls) {
          this.ensureViewControlsUsable();
        }
      } else {
        nc.trigger("editing:disable", false, {
          viewMode: false,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: true,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: false, disable: true },
          statusBar: true,
          leftMenu: { disable: false, previewMode: false },
          fileMenu: false,
          navigation: { disable: false, previewMode: false },
          comments: { disable: false, previewMode: false },
          chat: false,
          review: true,
          viewport: false,
          documentHolder: { clear: false, disable: true },
          toolbar: true,
          protect: true,
        });
      }

      this.scheduleWordDocModeHide();
      return;
    }

    // PDF / Excel 等：只读且 keepViewControls 时不盖 mask、不锁状态栏缩放。
    nc.trigger("editing:disable", disabled, {
      viewMode: disabled,
      reviewMode: false,
      fillFormMode: false,
      viewDocMode: false,
      allowMerge: true,
      allowSignature: false,
      allowProtect: false,
      rightMenu: { clear: false, disable: true },
      statusBar: disableStatusBar,
      leftMenu: { disable: false, previewMode: disabled },
      fileMenu: false,
      navigation: {
        disable: false,
        previewMode: disabled && !this.keepViewControls,
      },
      comments: { disable: false, previewMode: disabled },
      chat: false,
      review: true,
      viewport: false,
      documentHolder: { clear: false, disable: true },
      toolbar: maskToolbar,
    });
    if (disabled && this.keepViewControls) {
      this.ensureViewControlsUsable();
    }
  }

  private stopViewControlsWatchdog() {
    if (this.viewControlsTimer !== null) {
      window.clearInterval(this.viewControlsTimer);
      this.viewControlsTimer = null;
    }
    if (this.viewControlsWatchdogStopTimer !== null) {
      window.clearTimeout(this.viewControlsWatchdogStopTimer);
      this.viewControlsWatchdogStopTimer = null;
    }
    if (this.viewControlsObsTimer !== null) {
      window.clearTimeout(this.viewControlsObsTimer);
      this.viewControlsObsTimer = null;
    }
    if (this.viewControlsObserver) {
      this.viewControlsObserver.disconnect();
      this.viewControlsObserver = null;
    }
    this.viewControlsEnsuring = false;
    this.viewControlsUnlocked = false;
  }

  /**
   * OnlyOffice 会在 restriction / 工具栏延迟渲染后再次给缩放钮加 disabled。
   * 前 15s 较快轮询 + 防抖 MutationObserver；稳定解锁后降频，避免 DOM 反馈环卡死 iframe。
   */
  private startViewControlsWatchdog() {
    this.stopViewControlsWatchdog();
    if (!this.keepViewControls || !this.readOnly) {
      return;
    }

    const tick = () => {
      if (!this.readOnly || !this.keepViewControls) {
        this.stopViewControlsWatchdog();
        return;
      }
      this.ensureViewControlsUsable();
    };

    tick();
    this.viewControlsTimer = window.setInterval(tick, 800);
    this.viewControlsWatchdogStopTimer = window.setTimeout(() => {
      if (this.viewControlsTimer !== null) {
        window.clearInterval(this.viewControlsTimer);
        this.viewControlsTimer = null;
      }
      // 稳定后慢轮询，防 OO 后期又加锁
      this.viewControlsTimer = window.setInterval(tick, 3000);
    }, 15000);

    try {
      const win = this.getEditorFrameWindow();
      const toolbar = win?.document?.querySelector(".toolbar");
      if (toolbar && typeof MutationObserver !== "undefined") {
        this.viewControlsObserver = new MutationObserver(() => {
          if (this.viewControlsEnsuring) {
            return;
          }
          if (this.viewControlsObsTimer !== null) {
            window.clearTimeout(this.viewControlsObsTimer);
          }
          this.viewControlsObsTimer = window.setTimeout(() => {
            this.viewControlsObsTimer = null;
            tick();
          }, 250);
        });
        this.viewControlsObserver.observe(toolbar, {
          attributes: true,
          childList: true,
          subtree: true,
          attributeFilter: ["class", "disabled", "aria-disabled"],
        });
      }
    } catch {
      // ignore
    }

    this.installEditingDisableGuard();
  }

  /**
   * 拦截 iframe 内 editing:disable：keepViewControls 时强制不盖 toolbar-mask、不锁 statusBar。
   * Main.disableEditing / onDocModeApply 仍会自行 trigger，只靠事后清不够。
   */
  private installEditingDisableGuard() {
    if (!this.keepViewControls || !this.readOnly) {
      return;
    }
    const win = this.getEditorFrameWindow() as
      | (OnlyOfficeIframeWindow & {
          Common?: {
            NotificationCenter?: {
              trigger?: (...args: unknown[]) => unknown;
              __oowKeepViewPatched?: boolean;
            };
          };
        })
      | undefined;
    const nc = win?.Common?.NotificationCenter;
    if (!nc?.trigger || nc.__oowKeepViewPatched) {
      return;
    }
    const original = nc.trigger.bind(nc);
    nc.trigger = (...args: unknown[]) => {
      if (args[0] === "editing:disable" && this.keepViewControls && this.readOnly) {
        const disable = args[1];
        const options = (args[2] || {}) as Record<string, unknown>;
        if (disable) {
          args[2] = {
            ...options,
            toolbar: false,
            statusBar: false,
            viewport: false,
          };
        }
      }
      const ret = original(...args);
      if (args[0] === "editing:disable" && this.keepViewControls && this.readOnly) {
        window.setTimeout(() => this.ensureViewControlsUsable(), 0);
      }
      return ret;
    };
    nc.__oowKeepViewPatched = true;
  }

  /**
   * 去掉 toolbar-mask，并用 Common.Utils.lockControls 清 View 缩放锁。
   * injectButtons 初始 disabled:true + keepState；直接 setDisabled 会被 keepState 立刻锁回。
   */
  private ensureViewControlsUsable() {
    if (!this.keepViewControls || this.viewControlsEnsuring) {
      return;
    }

    const win = this.getEditorFrameWindow() as
      | (OnlyOfficeIframeWindow & {
          Common?: {
            Utils?: {
              lockControls?: (
                causes: unknown,
                lock: boolean,
                opts?: { array?: unknown[]; clear?: unknown[] },
              ) => void;
            };
            enumLock?: Record<string, unknown>;
          };
          DE?: { getController?: (name: string) => ViewTabControllerLike };
          PDFE?: { getController?: (name: string) => ViewTabControllerLike };
          PE?: { getController?: (name: string) => ViewTabControllerLike };
          SSE?: { getController?: (name: string) => ViewTabControllerLike };
          getApplication?: () => {
            getController?: (name: string) => ViewTabControllerLike;
          };
        })
      | undefined;
    if (!win) {
      return;
    }

    // 已解锁且无 mask 时跳过 DOM 写入，打断 MutationObserver 反馈环
    try {
      const root = win.document;
      const fw =
        root?.querySelector(".slot-btn-ftw .btn") ||
        root?.querySelector('[id*="tlbtn-btn-ftw"]');
      const hasMask = !!root?.querySelector(
        ".toolbar-mask, .toolbar-group-mask",
      );
      const fwOk =
        !!fw &&
        !fw.classList.contains("disabled") &&
        !fw.hasAttribute("disabled");
      if (
        this.viewControlsUnlocked &&
        fwOk &&
        !hasMask &&
        !!root?.getElementById("oow-keep-view-controls-css")
      ) {
        return;
      }
    } catch {
      // ignore and continue full unlock
    }

    this.viewControlsEnsuring = true;
    // 暂停观察，避免我们自己的 class/disabled 改动再次入队
    try {
      this.viewControlsObserver?.disconnect();
    } catch {
      // ignore
    }

    try {
      try {
        win.document
          ?.querySelectorAll(".toolbar-mask, .toolbar-group-mask")
          .forEach((el) => el.remove());
        win.document
          ?.querySelectorAll(".toolbar.masked")
          .forEach((el) => el.classList.remove("masked"));
      } catch {
        // ignore
      }

      const enableBtn = (btn: ZoomControlLike | null | undefined) => {
        if (!btn || typeof btn.setDisabled !== "function") {
          return;
        }
        try {
          if (Array.isArray(btn.keepState)) {
            btn.keepState.length = 0;
          }
          // 绕过 setDisabled 内部 this.disabled != disabled 短路
          if (btn.isDisabled?.()) {
            btn.setDisabled(false);
          } else {
            (btn as { disabled?: boolean }).disabled = true;
            btn.setDisabled(false);
          }
        } catch {
          // ignore
        }
      };

      const enableMany = (items: unknown) => {
        if (!items) {
          return;
        }
        if (Array.isArray(items)) {
          items.forEach((item) => enableBtn(item as ZoomControlLike));
          return;
        }
        const set = items as {
          setDisabled?: (v: boolean) => void;
          each?: (fn: (b: ZoomControlLike) => void) => void;
          forEach?: (fn: (b: ZoomControlLike) => void) => void;
        };
        if (typeof set.each === "function") {
          set.each((b) => enableBtn(b));
        } else if (typeof set.forEach === "function") {
          set.forEach((b) => enableBtn(b));
        } else if (typeof set.setDisabled === "function") {
          enableBtn(set as ZoomControlLike);
        }
      };

      try {
        const viewTab =
          win.DE?.getController?.("ViewTab") ??
          win.PDFE?.getController?.("ViewTab") ??
          win.PE?.getController?.("ViewTab") ??
          win.SSE?.getController?.("ViewTab") ??
          win.getApplication?.()?.getController?.("ViewTab");

        viewTab?.SetDisabled?.(false);
        const view = viewTab?.view ?? viewTab?.getView?.();
        const locked = view?.lockedControls;
        const enumLock = win.Common?.enumLock;
        const lockControls = win.Common?.Utils?.lockControls;
        if (lockControls && enumLock && Array.isArray(locked) && locked.length) {
          const causes = [
            enumLock.disableOnStart,
            enumLock.lostConnect,
            enumLock.viewMode,
            enumLock.previewReviewMode,
            enumLock.viewFormMode,
            enumLock.docLockView,
          ].filter((c) => c !== undefined);
          if (causes.length) {
            lockControls(causes, false, { array: locked });
          }
        }
        enableMany(view?.btnsFitToPage);
        enableMany(view?.btnsFitToWidth);
        enableBtn(view?.btnZoom100);
        enableBtn(view?.btnMultiplePages);
        enableMany(view?.cmbsZoom);
        enableMany(locked);
      } catch {
        // ignore
      }

      // DOM 兜底：去掉 disabled 样式，避免 keepState 残留导致仍灰
      const root = win.document;
      if (root) {
        if (!root.getElementById("oow-keep-view-controls-css")) {
          const style = root.createElement("style");
          style.id = "oow-keep-view-controls-css";
          style.textContent = `
.slot-btn-ftp .btn, .slot-btn-ftw .btn, #slot-btn-zoom-100 .btn, #slot-btn-multiple-pages .btn,
.slot-field-zoom .form-control, .slot-field-zoom input, .slot-field-zoom .btn,
[id*="tlbtn-btn-ftp"], [id*="tlbtn-btn-ftw"],
.slot-btn-ftp .btn.disabled, .slot-btn-ftw .btn.disabled, #slot-btn-zoom-100 .btn.disabled,
#slot-btn-multiple-pages .btn.disabled,
.slot-field-zoom .form-control.disabled, .slot-field-zoom input:disabled {
  opacity: 1 !important;
  filter: none !important;
  pointer-events: auto !important;
  cursor: pointer !important;
  color: inherit !important;
  background-color: inherit !important;
}
.slot-btn-ftp .btn[disabled], .slot-btn-ftw .btn[disabled], #slot-btn-zoom-100 .btn[disabled],
#slot-btn-multiple-pages .btn[disabled],
[id*="tlbtn-btn-ftp"][disabled], [id*="tlbtn-btn-ftw"][disabled] {
  pointer-events: auto !important;
}
`;
          root.head?.appendChild(style);
        }
        const slots = root.querySelectorAll(
          ".slot-btn-ftp button, .slot-btn-ftw button, #slot-btn-zoom-100 button, #slot-btn-multiple-pages button, [id*='tlbtn-btn-ftp'] button, [id*='tlbtn-btn-ftw'] button, [id*='tlbtn-btn-ftp'], [id*='tlbtn-btn-ftw'], #slot-btn-zoom-100 .btn, #slot-btn-multiple-pages .btn, .slot-field-zoom input, .slot-field-zoom .form-control, .slot-field-zoom button",
        );
        slots.forEach((el) => {
          el.classList.remove("disabled");
          el.removeAttribute("disabled");
          if (el instanceof HTMLElement) {
            el.style.pointerEvents = "auto";
            el.style.opacity = "1";
          }
        });

        const fw =
          root.querySelector(".slot-btn-ftw .btn") ||
          root.querySelector('[id*="tlbtn-btn-ftw"]');
        const unlocked =
          !!fw &&
          !fw.classList.contains("disabled") &&
          !fw.hasAttribute("disabled") &&
          !root.querySelector(".toolbar-mask, .toolbar-group-mask");
        this.viewControlsUnlocked = unlocked;

        // 探针 / 排障：只在状态变化时打一次
        try {
          const ftp =
            root.querySelector(".slot-btn-ftp .btn") ||
            root.querySelector('[id*="tlbtn-btn-ftp"]');
          const zoom = root.querySelector(
            ".slot-field-zoom input, .slot-field-zoom .form-control",
          );
          const report = {
            hasMask: !!root.querySelector(".toolbar-mask, .toolbar-group-mask"),
            hasCss: !!root.getElementById("oow-keep-view-controls-css"),
            fitWidthDisabled: fw
              ? fw.classList.contains("disabled") || fw.hasAttribute("disabled")
              : null,
            fitPageDisabled: ftp
              ? ftp.classList.contains("disabled") || ftp.hasAttribute("disabled")
              : null,
            zoomDisabled: zoom
              ? zoom.classList.contains("disabled") ||
                zoom.hasAttribute("disabled")
              : null,
            fitWidthPresent: !!fw,
          };
          const key = JSON.stringify(report);
          const g = win as { __oowViewZoomLast?: string };
          if (g.__oowViewZoomLast !== key) {
            g.__oowViewZoomLast = key;
            console.info("[oow-view-zoom]", key);
          }
        } catch {
          // ignore
        }
      }
    } finally {
      this.viewControlsEnsuring = false;
      // 重新挂上观察（若 watchdog 仍在）
      try {
        if (
          this.viewControlsObserver &&
          this.keepViewControls &&
          this.readOnly
        ) {
          const toolbar = win.document?.querySelector(".toolbar");
          if (toolbar) {
            this.viewControlsObserver.observe(toolbar, {
              attributes: true,
              childList: true,
              subtree: true,
              attributeFilter: ["class", "disabled", "aria-disabled"],
            });
          }
        }
      } catch {
        // ignore
      }
    }
  }

  /** 按配置强制套用缩放；localStorage 常在 DocEditor 创建后才可读，必须在 documentReady 再调一次。 */
  private applyConfiguredZoom() {
    const api = this.getSdkApi() as
      | {
          zoomFitToWidth?: () => void;
          zoomFitToPage?: () => void;
          zoom?: (percent: number) => void;
        }
      | undefined;
    if (!api) {
      return;
    }

    try {
      if (this.zoom === -2 && api.zoomFitToWidth) {
        api.zoomFitToWidth();
      } else if (this.zoom === -1 && api.zoomFitToPage) {
        api.zoomFitToPage();
      } else if (this.zoom > 0 && api.zoom) {
        api.zoom(this.zoom);
      }
    } catch {
      // ignore
    }
  }

  /**
   * 预览只读时去掉 iframe 的「离开此网站？」提示。
   * PDF 以 edit 权限挂载后再锁 VIEW 时，web-apps 仍挂 onBeforeUnload，
   * 缩放/限制会把文档标成已修改，切资源就会弹对话框。
   */
  private suppressLeavePromptForPreview() {
    if (!this.readOnly || !this.suppressLeavePrompt) {
      return;
    }

    const win = this.getEditorFrameWindow() as
      | (OnlyOfficeIframeWindow & {
          Common?: { localStorage?: { save?: () => void } };
        })
      | undefined;
    if (!win) {
      return;
    }

    win.onbeforeunload = () => {
      try {
        win.Common?.localStorage?.save?.();
      } catch {
        // ignore
      }
    };

    try {
      const api = this.getSdkApi() as
        | { SetDocumentModified?: (v: boolean) => void }
        | undefined;
      api?.SetDocumentModified?.(false);
    } catch {
      // ignore
    }
  }

  /** PPT 在通用只读逻辑之上追加：锁定 Main 控制器 + 禁用幻灯片侧栏。 */
  private syncSlideReadOnlyExtras(locked: boolean) {
    if (getDocumentType(this.fileType) !== DocumentType.Slide || !locked) {
      return;
    }
    this.lockShellEditMode();
  }

  /**
   * processRightsChange(true) 在 OnlyOffice 内无效果；false 会 asc_coAuthoringDisconnect 且 mode.isEdit=false。
   * 本地/mock 场景用 asc_setRestriction + 外壳 UI 同步，避免切回编辑仍停留在只读。
   */
  private restoreShellEditMode() {
    const main = this.getShellMainController();

    if (main?.mode) {
      main.mode.isEdit = true;
      main.mode.canEdit = true;
    }
    if (main?.appOptions) {
      main.appOptions.isEdit = true;
      main.appOptions.canEdit = true;
    }
    if (main?.permissions) {
      main.permissions.edit = true;
    }
    if (main?.editorConfig) {
      main.editorConfig.mode = "edit";
    }
    try {
      main?.disableEditing?.(false);
    } catch {
      // 解锁失败时仍继续清工具栏/单元格锁
    }
    main?.api?.asc_setViewMode?.(false);
    main?.api?.asc_enableKeyEvents?.(true);

    const win = this.getEditorFrameWindow() as OnlyOfficeIframeWindow & {
      SSE?: { getController?: (name: string) => {
        setMode?: (mode: unknown) => void;
        setPreviewMode?: (mode: boolean) => void;
        SetDisabled?: (disabled: boolean) => void;
        DisableToolbar?: (disable: boolean, viewMode: boolean) => void;
      } };
      DE?: { getController?: (name: string) => {
        setMode?: (mode: unknown) => void;
        setPreviewMode?: (mode: boolean) => void;
        SetDisabled?: (disabled: boolean) => void;
        DisableToolbar?: (disable: boolean, viewMode: boolean) => void;
      } };
      PE?: { getController?: (name: string) => {
        setMode?: (mode: unknown) => void;
        setPreviewMode?: (mode: boolean) => void;
        SetDisabled?: (disabled: boolean) => void;
        DisableToolbar?: (disable: boolean, viewMode: boolean) => void;
      } };
    };
    const app = win?.SSE ?? win?.DE ?? win?.PE;
    try {
      const cellEditor = app?.getController?.("CellEditor");
      cellEditor?.setMode?.(main?.appOptions);
      cellEditor?.setPreviewMode?.(false);
    } catch {
      // ignore
    }
    try {
      app?.getController?.("DocumentHolder")?.SetDisabled?.(false);
    } catch {
      // ignore
    }
    try {
      app?.getController?.("Toolbar")?.DisableToolbar?.(false, false);
    } catch {
      // ignore
    }

    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(true);
  }

  private lockShellEditMode() {
    const main = this.getShellMainController();

    if (main?.mode) {
      main.mode.isEdit = false;
      main.mode.canEdit = false;
    }

    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(false);
  }

  /** 兜底：拦截 SDK 层新增/复制幻灯片（toolbar 锁定之外）。 */
  private installSlideStructureEditBlocker() {
    if (getDocumentType(this.fileType) !== DocumentType.Slide) {
      return;
    }

    const patchApi = (
      api:
        | {
            AddSlide?: (...args: unknown[]) => unknown;
            DublicateSlide?: (...args: unknown[]) => unknown;
            __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
          }
        | undefined,
    ) => {
      if (!api || api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__) {
        return;
      }

      const guard = <T extends (...args: unknown[]) => unknown>(fn?: T) => {
        if (!fn) {
          return fn;
        }

        const bound = fn.bind(api);
        return (...args: unknown[]) => {
          if (this.readOnly) {
            return undefined;
          }
          return bound(...args);
        };
      };

      api.AddSlide = guard(api.AddSlide);
      api.DublicateSlide = guard(api.DublicateSlide);
      api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__ = true;
    };

    patchApi(
      this.getSdkApi() as unknown as {
        AddSlide?: (...args: unknown[]) => unknown;
        DublicateSlide?: (...args: unknown[]) => unknown;
        __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
      },
    );
    patchApi(
      (
        this.getShellMainController() as {
          api?: {
            AddSlide?: (...args: unknown[]) => unknown;
            DublicateSlide?: (...args: unknown[]) => unknown;
            __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
          };
        }
      )?.api,
    );
  }

  /** downloadAs → /downloadas/ → 更新 fsMap 中的 Editor.bin。 */
  private async captureDocumentSnapshot() {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }

    // 【本项目修改 2026-09-17】allowDownload 为假时，导出期间临时放开下载权限，见 grantDownloadForExport。
    const restoreDownload = this.grantDownloadForExport();
    try {
      return await this.server.captureCurrentDocument(() => {
        this.installIframeProxies();
        this.editor?.downloadAs("bin");
      });
    } finally {
      restoreDownload();
    }
  }

  /**
   * 【本项目新增 2026-09-17】`allowDownload: false` 时，让组件自己的导出照常工作。
   *
   * 编辑器按 `permissions.download: false` 加载之后，界面上的下载 / 另存入口都不出现，
   * 这正是要的。但组件自己导出（存回服务器之前那一步）走的也是 `downloadAs` 命令，
   * 而文字、表格、演示、PDF 四个编辑器的 `onDownloadAs` 第一句都是「`canDownload` 为假就拒绝」
   * （报 AccessDeny）——不处理的话，关掉下载就等于关掉了保存。
   *
   * 所以导出期间把编辑器 `Main` 控制器上的 `appOptions.canDownload` 临时置真，导出完放回去。
   * 界面上的下载入口是加载时按权限建的，这段时间里不会冒出来。
   *
   * ⚠ **拿不到 `Main` 控制器就直接抛错，不去碰运气**：那样发下去的 downloadAs 必然被拒，
   * 而被拒只是编辑器里一条报错，导出那边等到超时才知道。拿不到通常是编辑器 iframe 跨源
   * （静态资源与页面不同源），或者编辑器还没加载完。
   *
   * ⚠ 用的是 `DE / SSE / PE / PDFE / VE` 这几个应用全局对象。`getShellMainController`
   * 只认 `DE / PE`，表格和普通 PDF 拿不到，所以这里单独写一份，不去改它（它还管着别的事）。
   */
  private grantDownloadForExport(): () => void {
    if (this.allowDownload) {
      return () => {};
    }

    const win = this.getEditorFrameWindow() as
      | (OnlyOfficeIframeWindow &
          Partial<
            Record<
              "DE" | "SSE" | "PE" | "PDFE" | "VE",
              {
                getController?: (name: string) =>
                  | { appOptions?: { canDownload?: boolean } }
                  | undefined;
              }
            >
          >)
      | undefined;
    const appOptions = (["DE", "SSE", "PE", "PDFE", "VE"] as const)
      .map((key) => win?.[key]?.getController?.("Main")?.appOptions)
      .find((options) => options !== undefined);

    if (!appOptions) {
      throw new Error(
        "allowDownload 为 false 时，导出要临时放开编辑器的下载权限，但够不到编辑器里的 Main 控制器" +
          "（编辑器 iframe 跨源，或者编辑器还没加载完）",
      );
    }

    const previous = appOptions.canDownload;
    appOptions.canDownload = true;
    return () => {
      appOptions.canDownload = previous;
    };
  }

  /**
   * 只读模式下 downloadAs 可能被 SDK 拦截；导出前临时恢复编辑权再抓取。
   */
  private async captureDocumentSnapshotAllowingReadOnly() {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }

    const locked = this.readOnly;
    if (locked) {
      this.syncEditingRights(true);
    }

    try {
      return await this.captureDocumentSnapshot();
    } finally {
      if (locked) {
        this.syncEditingRights(false);
        this.syncSlideReadOnlyExtras(true);
      }
    }
  }

  private async captureDocumentIfDirty() {
    if (!this.editor || this.readOnly || !this.dirty) {
      return;
    }

    await this.captureDocumentSnapshot();
    this.dirty = false;
  }

  private destroyDocEditorInstance() {
    this.stopViewControlsWatchdog();
    this.disconnectConnector();
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.comments.clear();
    this.revisions = [];
    this.teardownWordContentSync();
  }

  /**
   * 创建 Developer Edition Connector。
   * Connector 运行在父页面，借助 DocsAPI 与编辑器 iframe 通信，因此可用于 CDN 跨域场景。
   */
  createConnector(options?: OnlyOfficeConnectorOptions): OnlyOfficeConnector {
    if (!this.editor) {
      throw new Error("OnlyOffice editor is not ready");
    }

    // 同一编辑器始终复用同一个 Connector。调用方可以在不再使用时主动
    // disconnect；下一次默认创建请求会重新 connect，而不会注册第二套监听器。
    if (this.connector) {
      if (options?.autoconnect !== false && !this.connector.isConnected) {
        this.connector.connect();
      }
      return this.connector;
    }

    const iframe = this.getEditorFrameElement();
    if (!iframe?.contentWindow) {
      throw new Error("OnlyOffice editor iframe is not ready");
    }

    // 9.4 的 DocsAPI 把 Connector 消息固定发送给 `iframeEditor`，但本组件
    // 以 containerId 生成 frameEditorId。先禁止自动连接，替换发送函数后再 connect。
    const connector = this.editor.createConnector({
      ...options,
      autoconnect: false,
    }) as OnlyOfficeConnector & {
      guid?: string;
      sendMessage?: (data: Record<string, unknown>) => void;
    };
    // DocsAPI 将 guid 仅作为 connector 回调的路由键。一个 EditorManager
    // 只维护一个 connector，因此用 containerId 可稳定地与编辑器一一对应。
    connector.guid = this.containerId;
    const iframeUrl = new URL(iframe.src, window.location.href);
    const frameEditorId =
      iframeUrl.searchParams.get("frameEditorId") ?? "iframeEditor";
    connector.sendMessage = (data) => {
      iframe.contentWindow?.postMessage(
        JSON.stringify({
          frameEditorId,
          type: "onExternalPluginMessage",
          subType: "connector",
          data: { ...data, guid: connector.guid },
        }),
        iframeUrl.origin,
      );
    };
    if (options?.autoconnect !== false) {
      connector.connect();
    }
    this.connector = connector;
    return this.connector;
  }

  private disconnectConnector() {
    if (this.connector?.isConnected) {
      this.connector.disconnect();
    }
    this.connector = null;
  }

  /**
   * 初始只读与运行时切只读走同一套 asc_setRestriction。
   * 挂载阶段若 permissions.edit=false，xlsx 等会在打开时样式/格式异常；
   * 因此挂载时保持完整编辑权限，documentReady 后再施加只读限制。
   * CDN 跨域模式通过 iframe 内 bridge 执行同样的 SDK/UI 同步，避免跨域访问 window.Asc。
   */
  private applyInitialReadOnlyState(documentType: DocumentType) {
    if (isOnlyOfficeCdnMode()) {
      this.syncCrossOriginReadOnly(this.readOnly);
      return;
    }

    this.installSlideStructureEditBlocker();
    this.syncEditingRights(false);
    this.suppressLeavePromptForPreview();
    this.applyConfiguredZoom();
    this.ensureViewControlsUsable();

    if (documentType === DocumentType.Slide) {
      // 工具栏 delayed render 后再锁一次，确保「新增幻灯片」按钮被 DisableToolbar 处理。
      window.setTimeout(() => {
        if (this.readOnly) {
          this.syncShellEditingDisable(true, documentType);
          this.suppressLeavePromptForPreview();
          this.ensureViewControlsUsable();
          this.applyConfiguredZoom();
        }
      }, 0);
    }

    if (documentType === DocumentType.Cell) {
      this.getSdkApi()?.asc_Recalculate?.();
    }

    // 工具栏延迟渲染后可能再次挂上 beforeunload / mask；再清一次并强制适应宽度。
    window.setTimeout(() => {
      this.suppressLeavePromptForPreview();
      this.ensureViewControlsUsable();
      this.applyConfiguredZoom();
    }, 0);
    window.setTimeout(() => {
      this.suppressLeavePromptForPreview();
      this.ensureViewControlsUsable();
      this.applyConfiguredZoom();
    }, 400);
  }

  /** 语言写在 iframe URL 的 lang 参数里，运行时 refreshFile 不会更新界面语言。 */
  /**
   * 把文档交给编辑器。
   *
   * ── 【本项目新增 2026-08-30】关于 document.directUrl ────────────────────
   *
   * 打开 PDF 时，编辑器**先加载一个分派页**（`web-apps/apps/common/index.html`），
   * 那个页面读文件开头 300 字节，判断这是不是一份**带可编辑内容的 PDF**
   * （OnlyOffice 自己导出的那种，头部有 `/ONLYOFFICEFORM` 标记、流里嵌着一个 OOXML 包）。
   * 判断结果决定往哪跳：
   *
   *   是可编辑 PDF → documenteditor（能填能改）
   *   普通 PDF     → pdfeditor（看和批注）
   *
   * 它读文件的办法是：有 `directUrl` 就直接部分下载它，**没有就退回文档服务器那套
   * `downloadfile/<key>` 端点**——而我们这套没有那个端点，请求 404，
   * 于是这个判断**永远回「不是」**。
   *
   * 后果分两种，其中一种不出声：普通 PDF **恰好判对**（答案本来就是「不是」）；
   * 而 **OnlyOffice 自己生成的 PDF 会被当成普通 PDF 打开、丢掉可编辑性，一句错都不报**。
   *
   * ── 为什么给的是 isForm 而不是 directUrl ──────────────────────────────
   *
   * 两条路都能让判断跑起来。选 `isForm` 是因为**同一个答案我们这边也要用**：
   * 可编辑 PDF 必须转成 Editor.bin 才进得了 documenteditor，而普通 PDF 要原样透传。
   * 给 `directUrl` 的话，编辑器判一次、我们判一次，
   * **等于要求两个判断永远一致**——它们迟早不一致，而不一致的样子是
   * 「进对了应用，却打不开文档」。所以我们判一次（见 internal/editor/pdf-form.ts），
   * 结果同时用于转换和这一格；编辑器见到 `isForm` 是 true 或 false 就不再自己判。
   *
   * ⚠ **只给 PDF 加。** 别的类型走不到那个分派页。
   */
  private mountDocEditor() {
    const doc = this.server.getDocument();
    const user = this.server.getUser();
    const documentType = getDocumentType(doc.fileType);

    // 【本项目新增 2026-08-30】挂上法律声明入口。
    // 许可的附加条款第三条要求界面里有一个清晰可达、显著可见的入口，
    // 让用户能认出原始开发者、知道这是修改版、拿到许可信息。
    // **挂在这里而不是挂在页面里**：这个组件是给别人引用的，放在页面里的话
    // 别人引用组件时就丢了，而丢了不会有任何东西报错。
    this.legalNoticeTeardown?.();
    const legalHost = this.getOfficeOverlayHostElement();
    this.legalNoticeTeardown = legalHost ? mountLegalNotice(legalHost) : null;

    this.server.setClient({
      buildVersion: window.DocsAPI!.DocEditor.version(),
    });

    this.preferFitWidthZoomPreference();

    this.editor = new window.DocsAPI!.DocEditor(this.containerId, {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        // 【本项目新增 2026-08-30】只给 PDF 加这一格，见上面那段说明。
        ...(documentType === DocumentType.Pdf ? { isForm: doc.isForm } : {}),
        permissions: this.getDocumentPermissions(true),
      },
      documentType,
      editorConfig: {
        lang: this.editorLang,
        // PDF 预览直接用 view：不挂「文档已修改」离开提示，也不进批注可编辑态。
        // Word/表格仍保持 edit 挂载后再锁，避免打开时样式异常。
        ...(this.readOnly && documentType === DocumentType.Pdf
          ? { mode: "view" as const }
          : {}),
        coEditing: {
          mode: "fast",
          change: false,
        },
        user: {
          ...user,
        },
        customization: this.buildEditorCustomization(),
        // 【本 PoC 新增】上游这里没有这一格，插件因此拿不到 options。见 OnlyOfficePluginsConfig。
        ...(this.pluginsConfig ? { plugins: this.pluginsConfig } : {}),
      },
      events: {
        onAppReady: () => {
          // 尽早安装代理，供 WebSocket auth 与后续 downloadAs 使用。
          this.installIframeProxies();
          this.preferFitWidthZoomPreference(
            this.getEditorFrameWindow()?.localStorage,
          );
        },
        onDocumentReady: () => {
          this.installSaveShortcutBlocker();
          this.installCommentResolveCleanup();
          this.installSlideStructureEditBlocker();
          this.applyDefaultReviewSettings();
          void this.ensureWordContentSync();
          onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, {
            fileName: doc.title,
            fileType: doc.fileType,
            instanceId: this.instanceId,
          });

          if (this.readOnly) {
            this.applyInitialReadOnlyState(documentType);
          } else {
            this.restoreShellEditMode();
            this.syncShellEditingDisable(false, documentType);
            if (documentType === DocumentType.Word) {
              this.scheduleWordDocModeHide();
            }
            this.applyConfiguredZoom();
          }
        },
        onDocumentStateChange: (event: { data: boolean }) => {
          if (event.data) {
            this.dirty = true;
          }
        },
        // DocsAPI 仅在注册此回调时暴露「文件 → 重命名」入口。
        // 回调由 OnlyOffice 的内部 Gateway 发出；标题显示已由 iframe 更新，
        // 使用 iframe 内 asc_wopi_renameFile 发起 RPC，由内存服务按 WOPI 协议回包。
        // onRequestRename: (event: { data?: unknown }) => {
        //   const fileName =
        //     typeof event.data === "string" ? event.data : "";
        //   void this.renameDocument(fileName).catch((error) => {
        //     this.logger.error("operation", "OnlyOffice document rename failed", {
        //       fileName,
        //       instanceId: this.instanceId,
        //       containerId: this.containerId,
        //       error: error instanceof Error ? error.message : String(error),
        //     });
        //   });
        // },
        // 不注册 onSave/onSaveDocument：内部保存已禁用，导出统一走 export() → downloadAs。
        onDownloadAs: () => {
          // Required so DocsAPI.downloadAs can request the current editor binary.
        },
      },
      // 【本项目修改 2026-08-30】原本写死 "desktop"，恒加载完整编辑器 main。
      // 见 OnlyOfficeEditorVariant。**这里与 buildRefreshFileConfig 那处必须一起改。**
      type: editorConfigType(this.variant),
      width: "100%",
      height: "100%",
    });
  }

  private buildRefreshFileConfig(editing: boolean) {
    const doc = this.server.getDocument();
    return {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        permissions: this.getDocumentPermissions(editing),
      },
      documentType: getDocumentType(doc.fileType),
      editorConfig: {
        mode: editing ? "edit" : "view",
        lang: this.editorLang,
      },
      // 【本项目修改 2026-08-30】同 mountDocEditor 那处。
      // 只改那一处的话，一切换只读就会从查看器跳回编辑器，而那一跳不报错。
      type: editorConfigType(this.variant),
      width: "100%",
      height: "100%",
    };
  }

  /** 就地切换编辑权限（主路径 asc_setRestriction；PPT 只读额外 syncSlideReadOnlyExtras）。 */
  private syncEditingRights(editing: boolean) {
    if (!this.editor) {
      return;
    }

    const documentType = getDocumentType(this.fileType);

    if (isOnlyOfficeCdnMode()) {
      this.syncCrossOriginReadOnly(!editing);
      return;
    }

    const sdk = this.getRestrictionSdkApi();

    if (sdk?.asc_setRestriction) {
      if (editing) {
        sdk.asc_removeRestriction?.(ASC_RESTRICTION_VIEW);
        sdk.asc_setRestriction(ASC_RESTRICTION_NONE);
        this.restoreShellEditMode();
        this.stopViewControlsWatchdog();
        this.syncShellEditingDisable(false, documentType);
      } else {
        sdk.asc_setRestriction(ASC_RESTRICTION_VIEW);
        this.syncSlideReadOnlyExtras(true);
        // keepViewControls：不要走 editing:disable。
        // 即便 toolbar:false，viewMode/leftMenu 等仍会让 View 缩放钮反复变灰。
        // 编辑拦截只靠 VIEW restriction；PPT 另有结构拦截。
        if (this.keepViewControls && documentType !== DocumentType.Slide) {
          if (documentType === DocumentType.Word) {
            this.scheduleWordDocModeHide();
          }
          this.ensureViewControlsUsable();
          this.startViewControlsWatchdog();
        } else {
          this.syncShellEditingDisable(true, documentType);
          if (this.keepViewControls) {
            this.ensureViewControlsUsable();
            this.startViewControlsWatchdog();
          }
        }
      }
      return;
    }

    if (editing) {
      this.restoreShellEditMode();
      this.stopViewControlsWatchdog();
      this.syncShellEditingDisable(false, documentType);
      this.editor.refreshFile?.(this.buildRefreshFileConfig(true));
    } else {
      this.syncSlideReadOnlyExtras(true);
      if (documentType !== DocumentType.Slide) {
        this.editor.denyEditingRights?.("");
      }
      if (this.keepViewControls && documentType !== DocumentType.Slide) {
        if (documentType === DocumentType.Word) {
          this.scheduleWordDocModeHide();
        }
        this.ensureViewControlsUsable();
        this.startViewControlsWatchdog();
      } else {
        this.syncShellEditingDisable(true, documentType);
        if (this.keepViewControls) {
          this.ensureViewControlsUsable();
          this.startViewControlsWatchdog();
        }
      }
    }
  }

  private createExportData(
    snapshot: ReturnType<EditorServer["getDocumentSnapshot"]>,
  ) {
    const binData = snapshot.binData;

    if (!binData) {
      throw new Error("No OnlyOffice document data is available to export");
    }

    return {
      fileName: snapshot.fileName || this.fileName,
      fileType: snapshot.fileType || this.fileType,
      binData,
      instanceId: this.instanceId,
      media: {
        ...snapshot.media,
        ...this.media,
      },
      themes: snapshot.themes,
    };
  }

  private userSaveTimer: number | null = null;
  private pendingUserSaveSnapshot: ReturnType<
    EditorServer["getDocumentSnapshot"]
  > | null = null;

  /** 用户保存：更新快照并广播 SAVE_DOCUMENT + ONSAVE（同 tick 内合并重复回调）。 */
  private notifyUserSave(
    snapshot?: ReturnType<EditorServer["getDocumentSnapshot"]>,
  ) {
    if (snapshot) {
      this.pendingUserSaveSnapshot = snapshot;
    }

    if (this.userSaveTimer !== null) {
      window.clearTimeout(this.userSaveTimer);
    }

    this.userSaveTimer = window.setTimeout(() => {
      this.userSaveTimer = null;
      this.dirty = false;

      const snap =
        this.pendingUserSaveSnapshot ?? this.server.getDocumentSnapshot();
      this.pendingUserSaveSnapshot = null;

      const data = this.createExportData(snap);
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.ONSAVE, {
        fileName: data.fileName,
        instanceId: this.instanceId,
      });
    }, 0);
  }

  private isLoadSessionActive(containerId: string, loadSession?: number) {
    return (
      loadSession === undefined ||
      editorManagerFactory.isLoadSessionActive(containerId, loadSession)
    );
  }

  async create(options: CreateEditorViewOptions) {
    const containerId =
      options.containerId || this.containerId || ONLYOFFICE_ID;

    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }

    this.destroy();
    this.readOnly = resolveIsViewMode(options);
    this.revisionReviewMode = !!options.revisionReview;
    this.server.setOfficeXmlEventConfig(options.officeXmlEvent);
    if (options.user) {
      this.server.setUser(options.user);
    }
    this.containerId = containerId;

    const fileType = getFileType(options.fileName, options.fileType);
    this.fileName = options.fileName;
    this.fileType = fileType;
    this.media = {};
    this.comments.clear();
    this.revisions = [];
    this.clearOfficeXmlSizeLimitOverlay();
    this.teardownWordContentSync();

    if (options.isNew) {
      this.server.openNew(fileType, options.fileName);
    } else if (options.file) {
      await this.server.open(options.file, {
        fileName: options.fileName,
        fileType,
      });
    } else if (options.url) {
      await this.server.openUrl(options.url, {
        fileName: options.fileName,
        fileType,
        loader: options.loader,
      });
    } else {
      throw new Error("OnlyOffice requires a file, url, or new document type");
    }

    if (this.officeXmlSizeLimitPayload) {
      this.renderOfficeXmlSizeLimitOverlay();
      return this;
    }

    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }

    await initializeOnlyOffice(getDocumentType(fileType));

    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }

    this.editorLang =
      (options.lang as OnlyOfficeLang | undefined) || getOnlyOfficeLang();
    this.uiTheme = options.theme || "theme-white";
    this.pluginsConfig = options.plugins;   // 【本 PoC 新增】
    this.variant = options.variant ?? "editor";   // 【本项目新增 2026-08-30】
    this.allowDownload = options.allowDownload ?? true;   // 【本项目新增 2026-09-17】
    this.zoom = typeof options.zoom === "number" ? options.zoom : -2;
    this.migrateFitPageZoom =
      options.migrateFitPageZoom ?? this.zoom === -2;
    this.suppressLeavePrompt = options.suppressLeavePrompt ?? true;
    this.keepViewControls = options.keepViewControls ?? true;

    this.syncEditorBridge();
    this.mountDocEditor();
    this.renderOfficeXmlSizeLimitOverlay();

    return this;
  }

  exists() {
    return !!this.editor;
  }

  /** 导出链路：downloadAs("bin") → server.resolvePendingExport → SAVE_DOCUMENT 事件。 */
  async export() {
    let snapshot;
    if (this.editor && (!this.readOnly || this.dirty)) {
      snapshot = await this.captureDocumentSnapshotAllowingReadOnly();
      this.dirty = false;
    } else {
      snapshot = this.server.getDocumentSnapshot();
    }
    const data = this.createExportData(snapshot);

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);

    return data;
  }

  getUser(): User {
    return this.server.getUser();
  }

  setUser(user: User) {
    this.server.setUser(user);
    this.editor?.setUsers?.([{ id: user.id, name: user.name }]);
  }

  /** 就地切换只读；切到只读前先 downloadAs 落盘，避免后续导出仍是打开时的 Editor.bin。 */
  async setReadOnly(readOnly: boolean) {
    return this.setMode(readOnly ? "view" : "edit");
  }

  /** 简单切换：view=只读预览，edit=可编辑。 */
  async setMode(mode: OnlyOfficeDocumentMode) {
    const readOnly = mode === "view";
    if (this.readOnly === readOnly) {
      return;
    }

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true,
    });

    try {
      if (readOnly && this.editor) {
        await this.captureDocumentSnapshot();
        this.dirty = false;
      }

      this.readOnly = readOnly;
      if (isOnlyOfficeCdnMode()) {
        this.syncCrossOriginReadOnly(readOnly);
        return;
      }

      this.installSlideStructureEditBlocker();
      this.syncEditingRights(!readOnly);
      if (readOnly) {
        this.suppressLeavePromptForPreview();
      }
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false,
      });
    }
  }

  getMode(): OnlyOfficeDocumentMode {
    return this.readOnly ? "view" : "edit";
  }

  getReadOnly() {
    return this.readOnly;
  }

  async setLanguage(lang: OnlyOfficeLang) {
    if (this.editorLang === lang) {
      return;
    }

    this.editorLang = lang;

    if (!this.editor) {
      return;
    }

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true,
    });

    try {
      await this.captureDocumentIfDirty();
      this.destroyDocEditorInstance();
      this.mountDocEditor();
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false,
      });
    }
  }

  getTheme(): OfficeTheme {
    return this.uiTheme;
  }

  /** uiTheme 写在 iframe URL 参数里，运行时需 remount 才能生效。 */
  async setTheme(theme: OfficeTheme) {
    if (this.uiTheme === theme) {
      return;
    }

    this.uiTheme = theme;

    if (!this.editor) {
      return;
    }

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true,
    });

    try {
      await this.captureDocumentIfDirty();
      this.destroyDocEditorInstance();
      this.mountDocEditor();
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false,
      });
    }
  }

  getInstanceId() {
    return this.instanceId;
  }

  isOfficeXmlSizeLimitExceeded() {
    return !!this.officeXmlSizeLimitPayload;
  }

  getContainerId() {
    return this.containerId;
  }

  getLogger() {
    return this.logger;
  }

  printLogs() {
    this.logger.print();
  }

  getFileName() {
    return this.fileName;
  }

  /**
   * @description 通过 iframe 内 asc_wopi_renameFile 重命名当前实例。
  * SDK 经 socket rpc 收到内存 WOPI 回包后才更新本地标题，因此返回 Promise。
   */
  renameDocument(fileName: string): Promise<string> {
    const requestedName =
      typeof fileName === "string" ? fileName.trim() : "";
    if (!requestedName) {
      return Promise.reject(
        new Error("OnlyOffice document name cannot be empty"),
      );
    }

    if (this.pendingRename) {
      return Promise.reject(new Error("OnlyOffice document rename is pending"));
    }

    return new Promise<string>((resolve, reject) => {
      const timer = window.setTimeout(() => {
        if (this.pendingRename?.timer !== timer) {
          return;
        }
        this.pendingRename = null;
        reject(new Error("OnlyOffice document rename timed out"));
      }, 5000);

      this.pendingRename = { resolve, reject, timer };

      try {
        if (isOnlyOfficeCdnMode()) {
          void callCrossOriginEditor(
            this.containerId,
            CROSS_ORIGIN_EDITOR_COMMAND.DOCUMENT_RENAME,
            { fileName: requestedName },
          )
            .then(() => {
              // 跨域 bridge 只能确认 iframe 已调用 SDK API；原生「文件 → 重命名」
              // 路径不一定把 WOPI RPC 回包转回父页。由宿主提交标题，确保导出
              // 快照与当前文件名同步；若 RPC 已先到达，pendingRename 已被清空。
              if (this.pendingRename) {
                this.server.rename(requestedName);
              }
            })
            .catch((error) => this.rejectPendingRename(error));
          return;
        }

        const api = this.requireSdkApi();
        if (!api.asc_wopi_renameFile) {
          throw new Error("OnlyOffice WOPI rename API is not available");
        }
        api.asc_wopi_renameFile(requestedName);
      } catch (error) {
        this.rejectPendingRename(error);
      }
    });
  }

  private syncRenamedDocument(renamedFileName: string) {
    this.fileName = renamedFileName;

    const document = getDocumentObj(this.containerId);
    setDocumentObj(
      {
        ...document,
        fileName: renamedFileName,
      },
      this.containerId,
    );

    const pendingRename = this.pendingRename;
    if (pendingRename) {
      window.clearTimeout(pendingRename.timer);
      this.pendingRename = null;
      pendingRename.resolve(renamedFileName);
    }

    this.logger.operation("OnlyOffice document renamed through WOPI RPC", {
      fileName: renamedFileName,
      instanceId: this.instanceId,
      containerId: this.containerId,
    });
  }

  private rejectPendingRename(error: unknown) {
    const pendingRename = this.pendingRename;
    if (!pendingRename) {
      return;
    }

    window.clearTimeout(pendingRename.timer);
    this.pendingRename = null;
    pendingRename.reject(
      error instanceof Error ? error : new Error(String(error)),
    );
  }

  getContainerParentSelector() {
    return `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"], ${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}`;
  }

  getContainerStyle() {
    return ONLYOFFICE_CONTAINER_CONFIG.STYLE;
  }

  updateMedia(key: string, data: Uint8Array) {
    this.media[key] = data;
  }

  getMedia() {
    return { ...this.media };
  }

  isDirty() {
    return this.dirty;
  }

  async subscribe({
    type,
    fn,
  }: {
    type: AscWordApiMethod;
    fn: AscWordApiCallback;
  }) {
    if (isOnlyOfficeCdnMode()) {
      if (
        type === CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT ||
        type === CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT ||
        type === CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT
      ) {
        await this.callCrossOriginComment(
          CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
          {},
        );
        return subscribeCrossOriginEditorEvent(this.containerId, type, (args) =>
          fn(...args),
        );
      }

      if (
        type === CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE ||
        type === CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE
      ) {
        await this.callCrossOriginRevision(
          CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE,
        );
        return subscribeCrossOriginEditorEvent(this.containerId, type, (args) =>
          fn(...args),
        );
      }

      if (type === CROSS_ORIGIN_EDITOR_EVENT.DOCUMENT_MODIFIED_CHANGED) {
        await callCrossOriginEditor(
          this.containerId,
          CROSS_ORIGIN_EDITOR_COMMAND.EDITOR_SUBSCRIBE,
          { event: type },
        );
        return subscribeCrossOriginEditorEvent(this.containerId, type, (args) =>
          fn(...args),
        );
      }

      throw new Error(
        `OnlyOffice cross-origin callback is not supported: ${type}`,
      );
    }

    const api = this.requireSdkApi();

    if (!api.asc_registerCallback || !api.asc_unregisterCallback) {
      throw new Error("OnlyOffice callback subscription is not supported");
    }

    api.asc_registerCallback(type, fn);

    return () => {
      api.asc_unregisterCallback?.(type, fn);
    };
  }

  async getAllComments(): Promise<CommentItem[]> {
    if (isOnlyOfficeCdnMode()) {
      return this.refreshCrossOriginComments();
    }

    this.refreshCommentsFromSdk();
    return Array.from(this.comments.entries()).map(([Id, Data]) => ({
      Id,
      Data,
    }));
  }

  refreshComments() {
    return this.getAllComments();
  }

  private createSdkCommentPayload(data: CommentData): unknown {
    const asc = this.getEditorFrameWindow()?.Asc as
      | (Record<string, unknown> & {
          asc_CCommentDataWord?: new (value: unknown) => {
            asc_putText?: (value: string) => void;
            asc_putUserName?: (value: string) => void;
            asc_putTime?: (value: string) => void;
            asc_putQuoteText?: (value: string) => void;
            asc_putSolved?: (value: boolean) => void;
            asc_putUserData?: (value: string) => void;
          };
        })
      | undefined;
    const CommentDataWord = asc?.asc_CCommentDataWord;

    if (!CommentDataWord) {
      return data;
    }

    const comment = new CommentDataWord(null);
    const payload = toPluginCommentPayload(data);

    if (payload.Text != null) {
      comment.asc_putText?.(String(payload.Text));
    }
    if (payload.UserName != null) {
      comment.asc_putUserName?.(String(payload.UserName));
    }
    if (payload.Time != null) {
      comment.asc_putTime?.(String(payload.Time));
    }
    if (payload.QuoteText != null) {
      comment.asc_putQuoteText?.(String(payload.QuoteText));
    }
    if (typeof payload.Solved === "boolean") {
      comment.asc_putSolved?.(payload.Solved);
    }
    if (payload.UserData != null) {
      comment.asc_putUserData?.(String(payload.UserData));
    }

    return comment;
  }

  async callCrossOriginComment(
    command: string,
    payload: Record<string, unknown>,
  ) {
    return callCrossOriginEditor(this.containerId, command, payload);
  }

  addComment(input: CommentInput) {
    if (isOnlyOfficeCdnMode()) {
      const data = toPluginCommentPayload(normalizeCommentInput(input));
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_ADD,
        { data },
      ).then((id) => {
        if (id) {
          this.comments.set(String(id), data);
        }
        return id ? String(id) : "";
      });
    }

    const api = this.requireSdkApi();
    const data = toPluginCommentPayload(normalizeCommentInput(input));
    const id =
      api.pluginMethod_AddComment?.(data) ??
      api.asc_addComment?.(this.createSdkCommentPayload(data) as CommentData);
    if (id) {
      this.comments.set(String(id), data);
    }
    return id ? String(id) : "";
  }

  updateComment(id: string, data: CommentData) {
    if (isResolvedComment(data)) {
      return this.removeComment(id);
    }

    if (isOnlyOfficeCdnMode()) {
      const payload = toPluginCommentPayload(data);
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_UPDATE,
        {
          id,
          data: payload,
        },
      ).then(() => {
        this.comments.set(id, payload);
      });
    }

    const api = this.requireSdkApi();
    const payload = toPluginCommentPayload(data);

    if (typeof api.pluginMethod_ChangeComment === "function") {
      api.pluginMethod_ChangeComment(id, payload);
    } else {
      api.asc_changeComment?.(
        id,
        this.createSdkCommentPayload(payload) as CommentData,
      );
    }
    this.comments.set(id, payload);
  }

  removeComment(id: string) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_REMOVE,
        { id },
      ).then(() => {
        this.comments.delete(id);
      });
    }

    this.requireSdkApi().asc_removeComment?.(id);
    this.comments.delete(id);
  }

  goToComment(
    id: string,
    { showBalloon = false }: { showBalloon?: boolean } = {},
  ) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_GO_TO,
        { id, showBalloon },
      );
    }

    const api = this.requireSdkApi();
    api.asc_selectComment?.(id);
    if (showBalloon) {
      api.asc_showComment?.(id);
    }
  }

  async registerCommentCallbacks(handlers: CommentChangeHandlers) {
    if (isOnlyOfficeCdnMode()) {
      const unsubscribers: Array<() => void> = [];
      await this.callCrossOriginComment(
        CROSS_ORIGIN_EDITOR_COMMAND.COMMENT_SUBSCRIBE,
        {},
      );

      if (handlers.onAdd) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.ADD_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              this.comments.set(commentId, commentData);
              handlers.onAdd?.(commentId, commentData);
            },
          ),
        );
      }

      if (handlers.onChange) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.CHANGE_COMMENT,
            ([id, data]) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                this.comments.delete(commentId);
                handlers.onRemove?.(commentId);
                return;
              }

              this.comments.set(commentId, commentData);
              handlers.onChange?.(commentId, commentData);
            },
          ),
        );
      }

      if (handlers.onRemove) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.REMOVE_COMMENT,
            ([id]) => {
              const commentId = String(id);
              this.comments.delete(commentId);
              handlers.onRemove?.(commentId);
            },
          ),
        );
      }

      return () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe());
      };
    }

    const unsubscribers = await Promise.all([
      handlers.onAdd
        ? this.subscribe({
            type: "asc_onAddComment",
            fn: (id, data) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              this.comments.set(commentId, commentData);
              handlers.onAdd?.(commentId, commentData);
            },
          })
        : undefined,
      handlers.onChange
        ? this.subscribe({
            type: "asc_onChangeCommentData",
            fn: (id, data) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                window.setTimeout(() => {
                  this.removeComment(commentId);
                  handlers.onRemove?.(commentId);
                }, 0);
                return;
              }

              this.comments.set(commentId, commentData);
              handlers.onChange?.(commentId, commentData);
            },
          })
        : undefined,
      handlers.onRemove
        ? this.subscribe({
            type: "asc_onRemoveComment",
            fn: (id) => {
              const commentId = String(id);
              this.comments.delete(commentId);
              handlers.onRemove?.(commentId);
            },
          })
        : undefined,
    ]);

    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }

  setTrackRevisions(enabled: boolean) {
    this.trackRevisions = enabled;
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SET_TRACK,
        { enabled },
      );
    }

    const api = this.requireSdkApi() as RevisionsEditorApi;
    api.asc_SetGlobalTrackRevisions?.(enabled);
    api.asc_SetLocalTrackRevisions?.(enabled);
  }

  /** 修订审阅页初始化：开启追踪 + markup 显示模式（不会批量接受/拒绝） */
  prepareRevisionReview() {
    this.trackRevisions = true;
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_PREPARE_REVIEW,
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    const api = this.requireSdkApi() as RevisionsEditorApi;
    api.asc_SetGlobalTrackRevisions?.(true);
    prepareRevisionReviewDisplay(api, this.getEditorFrameWindow());
  }

  isTrackRevisions() {
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_IS_TRACK,
      ).then((enabled) => {
        this.trackRevisions = !!enabled;
      });
      return this.trackRevisions;
    }

    return !!this.getSdkApi()?.asc_GetGlobalTrackRevisions?.();
  }

  haveRevisionsChanges() {
    if (isOnlyOfficeCdnMode()) {
      void this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_HAVE_CHANGES,
      ).then((hasChanges) => {
        if (!hasChanges) {
          this.revisions = [];
        } else {
          void this.refreshCrossOriginRevisions();
        }
      });
      return this.revisions.length > 0;
    }

    const api = this.getSdkApi() as RevisionsEditorApi | undefined;
    if (typeof api?.asc_HaveRevisionsChanges === "function") {
      if (api.asc_HaveRevisionsChanges(true)) {
        return true;
      }
      if (api.asc_HaveRevisionsChanges()) {
        return true;
      }
    }

    return this.revisions.length > 0;
  }

  async getAllRevisions(): Promise<RevisionItem[]> {
    if (isOnlyOfficeCdnMode()) {
      return this.refreshCrossOriginRevisions({ forceRefreshStack: true });
    }

    this.refreshRevisionsFromSdk({ forceRefreshStack: true });
    return this.revisions;
  }

  refreshRevisions() {
    return this.getAllRevisions();
  }

  goToNextRevision() {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_NEXT,
      );
    }

    (
      this.getSdkApi() as RevisionsEditorApi | undefined
    )?.asc_GetNextRevisionsChange?.();
  }

  goToPrevRevision() {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_PREV,
      );
    }

    (
      this.getSdkApi() as RevisionsEditorApi | undefined
    )?.asc_GetPrevRevisionsChange?.();
  }

  goToRevision(id: string) {
    if (isOnlyOfficeCdnMode()) {
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_GO_TO,
        {
          id,
          index: cached?.Index,
        },
      );
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) return;

    let cached = this.revisions.find((entry) => entry.Id === id);
    if (!cached) {
      this.refreshRevisionsFromSdk();
      cached = this.revisions.find((entry) => entry.Id === id);
    }

    goToRevisionInSdk(
      cached ?? id,
      api as RevisionsEditorApi,
      frameWin,
      this.revisions,
    );
  }

  acceptRevision(revision: RevisionItem | string) {
    if (isOnlyOfficeCdnMode()) {
      const id = this.revisionTargetId(revision);
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT,
        {
          id,
          index: cached?.Index,
        },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      return;
    }

    if (
      applyRevisionChange(
        "accept",
        revision,
        api as RevisionsEditorApi,
        frameWin,
        this.revisions,
      )
    ) {
      this.syncRevisionsAfterMutation();
    }
  }

  rejectRevision(revision: RevisionItem | string) {
    if (isOnlyOfficeCdnMode()) {
      const id = this.revisionTargetId(revision);
      const cached = this.revisions.find((entry) => entry.Id === id);
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT,
        {
          id,
          index: cached?.Index,
        },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    const api = this.getSdkApi();
    const frameWin = this.getEditorFrameWindow();
    if (!api || !frameWin) {
      return;
    }

    if (
      applyRevisionChange(
        "reject",
        revision,
        api as RevisionsEditorApi,
        frameWin,
        this.revisions,
      )
    ) {
      this.syncRevisionsAfterMutation();
    }
  }

  acceptAllRevisions() {
    this.applyAllRevisionChanges("accept");
  }

  rejectAllRevisions() {
    this.applyAllRevisionChanges("reject");
  }

  acceptRevisionsBySelection(all?: boolean) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_ACCEPT_SELECTION,
        { all },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    this.requireSdkApi().asc_AcceptChangesBySelection?.(all);
  }

  rejectRevisionsBySelection(all?: boolean) {
    if (isOnlyOfficeCdnMode()) {
      return this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_REJECT_SELECTION,
        { all },
      ).then((items) => {
        this.revisions = this.normalizeCrossOriginRevisions(items);
      });
    }

    this.requireSdkApi().asc_RejectChangesBySelection?.(all);
  }

  async registerRevisionCallbacks(handlers: RevisionChangeHandlers) {
    if (isOnlyOfficeCdnMode()) {
      const unsubscribers: Array<() => void> = [];
      await this.callCrossOriginRevision(
        CROSS_ORIGIN_EDITOR_COMMAND.REVISION_SUBSCRIBE,
      );

      if (handlers.onShowChanges) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            ([items]) => {
              this.revisions = this.normalizeCrossOriginRevisions(items);
              handlers.onShowChanges?.(this.revisions);
            },
          ),
        );
      }

      if (handlers.onTrackRevisionsChange) {
        unsubscribers.push(
          subscribeCrossOriginEditorEvent(
            this.containerId,
            CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
            ([enabled]) => {
              this.trackRevisions = !!enabled;
              handlers.onTrackRevisionsChange?.(!!enabled);
            },
          ),
        );
      }

      return () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe());
      };
    }

    const unsubscribers = await Promise.all([
      handlers.onShowChanges
        ? this.subscribe({
            type: CROSS_ORIGIN_EDITOR_EVENT.SHOW_REVISIONS_CHANGE,
            fn: (stack) => {
              const api = this.getSdkApi();
              const frameWin = this.getEditorFrameWindow();
              if (!api || !frameWin) return;

              handlers.onShowChanges?.(
                resolveRevisionShowChanges(
                  stack,
                  api as RevisionsEditorApi,
                  frameWin,
                ),
              );
            },
          })
        : undefined,
      handlers.onTrackRevisionsChange
        ? this.subscribe({
            type: CROSS_ORIGIN_EDITOR_EVENT.TRACK_REVISIONS_CHANGE,
            fn: (enabled) => handlers.onTrackRevisionsChange?.(!!enabled),
          })
        : undefined,
    ]);
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }

  destroy() {
    this.stopViewControlsWatchdog();
    this.rejectPendingRename(new Error("OnlyOffice editor was destroyed"));
    if (this.userSaveTimer !== null) {
      window.clearTimeout(this.userSaveTimer);
      this.userSaveTimer = null;
    }
    this.pendingUserSaveSnapshot = null;
    this.teardownWordContentSync();
    this.crossOriginBridgeTeardown?.();
    this.crossOriginBridgeTeardown = null;
    this.clearOfficeXmlSizeLimitOverlay();
    this.legalNoticeTeardown?.();
    this.legalNoticeTeardown = null;
    unregisterScopedIo(this.containerId);
    unregisterCrossOriginBridge(this.containerId);
    this.disconnectConnector();
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.dirty = false;
    this.comments.clear();
    this.revisions = [];
    this.server.reset();
  }
}

class EditorManagerFactory {
  private defaultManager = new EditorManager();
  private managers = new Map<string, EditorManager>();
  private loadSessions = new Map<string, number>();

  beginLoadSession(containerId: string) {
    const next = (this.loadSessions.get(containerId) ?? 0) + 1;
    this.loadSessions.set(containerId, next);
    return next;
  }

  isLoadSessionActive(containerId: string, loadSession: number) {
    return this.loadSessions.get(containerId) === loadSession;
  }

  getDefault() {
    return this.defaultManager;
  }

  create(containerId: string) {
    const manager =
      this.managers.get(containerId) || new EditorManager(containerId);
    this.managers.set(containerId, manager);
    return manager;
  }

  get(containerId: string) {
    return this.managers.get(containerId) || this.create(containerId);
  }

  getAll() {
    return [this.defaultManager, ...this.managers.values()];
  }

  destroy(containerId: string) {
    const manager = this.managers.get(containerId);
    manager?.destroy();
    this.managers.delete(containerId);
  }

  destroyAll() {
    this.defaultManager.destroy();
    for (const manager of this.managers.values()) {
      manager.destroy();
    }
    this.managers.clear();
  }
}
export const editorManagerFactory = new EditorManagerFactory();
export const editorManager = editorManagerFactory.getDefault();
if (typeof window !== "undefined") {
  (window as any).editorManagerFactory = editorManagerFactory;
}
