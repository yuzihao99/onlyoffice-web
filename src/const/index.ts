import { DocumentType } from "../internal/editor/types";

// ── 编辑器容器 / 事件 ──────────────────────────────────────────

export const ONLYOFFICE_ID = "iframe-office-id";

export const ONLYOFFICE_CONTAINER_CONFIG = {
  PARENT_SELECTOR: ".onlyoffice-container",
  PARENT_CLASS_NAME: "onlyoffice-container",
  STYLE: {
    position: "absolute",
    inset: 0,
  },
} as const;

export const ONLYOFFICE_EVENT_KEYS = {
  SAVE_DOCUMENT: "saveDocument",
  DOCUMENT_READY: "documentReady",
  LOADING_CHANGE: "loadingChange",
  ONSAVE: "onSave",
  OFFICE_XML_SIZE_LIMIT_EXCEEDED: "officeXmlSizeLimitExceeded",
} as const;

export type OnlyOfficeEventKey =
  (typeof ONLYOFFICE_EVENT_KEYS)[keyof typeof ONLYOFFICE_EVENT_KEYS];

export const ONLYOFFICE_LANG_KEY = {
  ZH: "zh",
  EN: "en",
} as const;

/** OnlyOffice 界面主题（对应 editorConfig.customization.uiTheme） */
export const OFFICE_THEME = {
  LIGHT: "theme-light",
  CLASSIC_LIGHT: "theme-classic-light",
  WHITE: "theme-white",
  DARK: "theme-dark",
  NIGHT: "theme-night",
  CONTRAST_DARK: "theme-contrast-dark",
} as const;

export type OfficeThemeId =
  (typeof OFFICE_THEME)[keyof typeof OFFICE_THEME];

export const DEFAULT_OFFICE_THEME: OfficeThemeId = OFFICE_THEME.WHITE;

/** 示例 / UI 切换用的主题列表 */
export const OFFICE_THEME_OPTIONS: ReadonlyArray<{
  id: OfficeThemeId;
  label: string;
}> = [
  { id: OFFICE_THEME.WHITE, label: "浅色" },
  { id: OFFICE_THEME.CLASSIC_LIGHT, label: "经典浅色" },
  { id: OFFICE_THEME.LIGHT, label: "Light" },
  { id: OFFICE_THEME.DARK, label: "深色" },
  { id: OFFICE_THEME.NIGHT, label: "夜间" },
  { id: OFFICE_THEME.CONTRAST_DARK, label: "高对比深色" },
];

/** 只读 ↔ 编辑切换时，loading 最少展示时长（ms） */
export const READONLY_SWITCH_MIN_DELAY_MS = 200;

export type OfficeXmlEventConfig = {
  isEnable?: boolean;
  limitBytes?: number;
};

/** Office ZIP 内 XML 包内容解压后大小限制；超过时会在 x2t 转换前拦截。 */
export const OFFICE_XML_EVENT_CONFIG = {
  default: {
    isEnable: false,
    limitBytes: 1024 * 1024 * 2048,
  },
} as const satisfies { default: Required<OfficeXmlEventConfig> };

/** Asc.c_oAscRestrictionType（sdk-all-min.js: k.Mf / k.Hca） */
export const ASC_RESTRICTION_NONE = 0;
export const ASC_RESTRICTION_VIEW = 128;

// 【本项目修改 2026-08-30】这里原来有一个 OFFICE_EDITOR_LOGO，把编辑器左上角
// 换成微软 Office 的图标。整个删掉了，理由见 core/editor-manager.ts 里那段注释
// （别家的商标 + 白标在 AGPL 下不允许 + 许可要求能认出原始开发者）。
// **别再加回来**：要改标识就是白标，那是商业版才有的功能。

// ── 静态资源（SDK / x2t）────────────────────────────────────────

export type StaticResource = {
  /** 版本目录：升级资源时只改这里 */
  version: {
    onlyofficeSdk: string;
    x2t: string;
  };
  onlyoffice: {
    root: string;
    apiJs: string;
    preloadHtml: string;
    apiUrl: string;
    preloadUrl: string;
  };
  x2t: {
    root: string;
    script: string;
    wasm: string;
    /** PDF 导出字体目录（见 X2T_PDF_FONT_MANIFEST） */
    pdfFonts: {
      root: string;
      default: string;
    };
  };
};

export type OnlyOfficeStaticResourceOptions = {
  /** CDN packages 根地址，例如 https://ca0eac5f.onlyoffice-packages.pages.dev。 */
  cdnOrigin?: string | null;
  /** CDN 上的 OnlyOffice 目录版本；未传时使用当前 SDK 版本。 */
  onlyofficeVersion?: string | null;
};

const X2T_PDF_DEFAULT_FONT_FILE = "Carlito-Regular.ttf";

/**
 * 导出 PDF 时用哪些字体：左边是盘上的文件，右边是 x2t 会按哪些名字来要它。
 * - Carlito 四款 → Calibri（字宽与 Calibri 一致）
 * - Liberation Sans 四款 → Arial（字宽与 Arial 一致；**别映射到 DroidSansFallback**，会乱码）
 * - DroidSansFallback → 中日韩兜底，接住宋体/微软雅黑那几个常用名
 *
 * 【本项目修改 2026-08-30】这份清单改了两处，都不是小事：
 *
 * **一、四个 Arial 换成了 Liberation Sans。** 上游那套 `x2t-fonts/` 里装的是真的
 * Monotype Arial（name 表写着 Version 2.82、「本字体是 Monotype 的财产」），
 * 那不是能随分发包发出去的东西。Liberation Sans 是专门照 Arial 的字宽做的替代品、
 * SIL OFL 许可，而且就在我们自己那个社区版镜像里，换掉之后排版宽度不变。
 *
 * **二、上游那套字体的粗体和斜体是装反的**——名字叫 `*-Bold.ttf` 的文件里装的是斜体，
 * 叫 `*-Italic.ttf` 的装的是粗体，Arial 和 Carlito 两族都反。因为这张表是按文件名映射的，
 * 结果就是**导出的 PDF 里粗体印成斜体、斜体印成粗体，一句报错都没有**。
 * 我们这套字体由 `scripts/build-x2t-fonts.mjs` 配，它会逐个打开字体核对
 * 里面自报的样式与文件名一致，对不上当场停；`scripts/check-x2t-fonts.mjs` 随时能再核一遍。
 */
export const X2T_PDF_FONT_MANIFEST = [
  {
    file: "Carlito-Regular.ttf",
    aliases: ["Carlito.ttf", "Calibri.ttf"],
  },
  {
    file: "Carlito-Bold.ttf",
    aliases: ["Carlito_Bold.ttf", "Calibri_Bold.ttf"],
  },
  {
    file: "Carlito-Italic.ttf",
    aliases: ["Carlito_Italic.ttf", "Calibri_Italic.ttf"],
  },
  {
    file: "Carlito-BoldItalic.ttf",
    aliases: ["Carlito_Bold_Italic.ttf", "Calibri_Bold_Italic.ttf"],
  },
  {
    file: "LiberationSans-Regular.ttf",
    aliases: ["Arial.ttf", "Liberation Sans.ttf"],
  },
  {
    file: "LiberationSans-Bold.ttf",
    aliases: ["Arial_Bold.ttf", "Liberation Sans_Bold.ttf"],
  },
  {
    file: "LiberationSans-Italic.ttf",
    aliases: ["Arial_Italic.ttf", "Liberation Sans_Italic.ttf"],
  },
  {
    file: "LiberationSans-BoldItalic.ttf",
    aliases: ["Arial_Bold_Italic.ttf", "Liberation Sans_Bold_Italic.ttf"],
  },
  {
    file: "DroidSansFallback.ttf",
    aliases: [
      "Droid Sans Fallback.ttf",
      "SimSun.ttf",
      "NSimSun.ttf",
      "宋体.ttf",
      "Microsoft YaHei.ttf",
      "微软雅黑.ttf",
      "PingFang SC.ttf",
    ],
  },
] as const;

/** 与 public/packages/onlyoffice 下的 Document Server 导出目录保持一致。 */
// 【本项目修改 2026-08-30】版本目录名改成社区版容器自报的引擎版本。
// 原来是 "9.4.0-develop"——那是上游从 Developer Edition 镜像抽资源时用的目录名。
// 我们的资源抽自社区版镜像，目录名取容器里 x2t 自报的那个号，好与 SOURCE.json 对上。
const DEFAULT_ONLYOFFICE_VERSION = "9.4.0.129";
const DEFAULT_ONLYOFFICE_ROOT = `/packages/onlyoffice/${DEFAULT_ONLYOFFICE_VERSION}`;

let staticResourceOptions: OnlyOfficeStaticResourceOptions | null = null;

function trimTrailingSlash(value: string) {
  return value.replace(/\/+$/, "");
}

/**
 * 预热哪一个编辑器：文档类型 → 预载页。
 *
 * 那几份页面由 `scripts/build-preload-pages.mjs` 生成（⚠ 文件名的拼法在那边也有一份，
 * 改名要两处一起改；跑一遍 `npm run e2e` 的 B17 会当场红）。
 *
 * ⚠ **上游那份 `preload.html` 刻意不再用**：它把四个编辑器全预热一遍，
 * 与这次要开的是什么文档无关，冷载因此多下约 90 MB（实测 221.7 MB → 141.1 MB）。
 * 留着那两个字段是因为它们是公开 API，别人可能引着。
 */
const PRELOAD_PAGE_BY_DOCUMENT_TYPE: Record<DocumentType, string> = {
  [DocumentType.Word]: "preload-documenteditor.html",
  [DocumentType.Cell]: "preload-spreadsheeteditor.html",
  [DocumentType.Slide]: "preload-presentationeditor.html",
  [DocumentType.Draw]: "preload-visioeditor.html",
  [DocumentType.Pdf]: "preload-pdfeditor.html",
};

/**
 * 这次该预热哪一份，返回站点相对路径。
 *
 * ⚠ 不传文档类型时按 Word 算，与 `getDocumentType()` 认不出扩展名时的默认值一致
 * ——两处要是各选各的默认，认不出扩展名的文档会预热 A 而实际打开 B，
 * **多下一份 SDK，而且没有任何东西报错**。
 */
export function getOnlyOfficePreloadPage(documentType?: DocumentType): string {
  const type = documentType ?? DocumentType.Word;
  const page = PRELOAD_PAGE_BY_DOCUMENT_TYPE[type] ?? PRELOAD_PAGE_BY_DOCUMENT_TYPE[DocumentType.Word];
  return `/web-apps/apps/api/documents/${page}`;
}

/** 同上，拼成能直接放进 iframe 的地址。 */
export function getOnlyOfficePreloadUrl(documentType?: DocumentType): string {
  return getStaticResource().onlyoffice.root + getOnlyOfficePreloadPage(documentType);
}

function buildStaticResource(): StaticResource {
  const apiJs = "/web-apps/apps/api/documents/api.js";
  const preloadHtml = "/web-apps/apps/api/documents/preload.html";
  const cdnOrigin = staticResourceOptions?.cdnOrigin
    ? trimTrailingSlash(staticResourceOptions.cdnOrigin)
    : "";
  const onlyofficeVersion = staticResourceOptions?.onlyofficeVersion || DEFAULT_ONLYOFFICE_VERSION;
  const onlyofficeRoot = cdnOrigin
    ? `${cdnOrigin}/onlyoffice/${onlyofficeVersion}`
    : DEFAULT_ONLYOFFICE_ROOT;
  const x2tRoot = `${onlyofficeRoot}/x2t`;
  const x2tPdfFontsRoot = `${onlyofficeRoot}/x2t-fonts`;

  return {
    version: {
      onlyofficeSdk: onlyofficeRoot,
      x2t: x2tRoot,
    },
    onlyoffice: {
      root: onlyofficeRoot,
      apiJs,
      preloadHtml,
      apiUrl: onlyofficeRoot + apiJs,
      preloadUrl: onlyofficeRoot + preloadHtml,
    },
    x2t: {
      root: x2tRoot,
      // 查询参数用来避开浏览器按启发式缓存住的旧 x2t.js。旧文件里的 heif 桩会 abort，编辑器只显示「未知错误」。
      script: `${x2tRoot}/x2t.js?v=20261023`,
      // 同一查询参数让浏览器重新拉取 wasm。旧 wasm 在部分图片回调为空时会直接中断转换。
      wasm: `${x2tRoot}/x2t.wasm?v=20261023d`,
      pdfFonts: {
        root: x2tPdfFontsRoot,
        default: `${x2tPdfFontsRoot}/${X2T_PDF_DEFAULT_FONT_FILE}`,
      },
    },
  };
}

let staticResourceCache: StaticResource | null = null;

/** 运行时注册 OnlyOffice 静态资源地址；须在首次初始化 DocsAPI / x2t 前调用。 */
export function registerOnlyOfficeStaticResource(
  options: OnlyOfficeStaticResourceOptions,
): StaticResource {
  staticResourceOptions = { ...options };
  staticResourceCache = null;
  return getStaticResource();
}

/** 清空运行时注册地址，恢复默认静态资源地址。 */
export function resetOnlyOfficeStaticResource(): StaticResource {
  staticResourceOptions = null;
  staticResourceCache = null;
  return getStaticResource();
}

/** 延迟构建，支持主实例初始化前运行时注册资源地址。 */
export function getStaticResource(): StaticResource {
  if (!staticResourceCache) {
    staticResourceCache = buildStaticResource();
  }
  return staticResourceCache;
}

/** 静态资源走外部 CDN（iframe 与主站跨域）。 */
export function isOnlyOfficeCdnMode(): boolean {
  const root = getStaticResource().onlyoffice.root;
  if (!/^https?:\/\//i.test(root) || typeof window === "undefined") {
    return false;
  }
  try {
    return new URL(root).origin !== window.location.origin;
  } catch {
    return false;
  }
}

/** 延迟读取，避免 Worker / SSR 在模块加载时固定资源地址。 */
export const STATIC_RESOURCE = {
  get version() {
    return getStaticResource().version;
  },
  get onlyoffice() {
    return getStaticResource().onlyoffice;
  },
  get x2t() {
    return getStaticResource().x2t;
  },
} as StaticResource;

/** 站点相对路径 → 绝对 URL（Worker 内 origin 用 self.location.origin） */
export function resolveSiteUrl(origin: string, path: string): string {
  if (/^https?:\/\//i.test(path)) {
    return path;
  }
  return `${origin}${path.startsWith("/") ? path : `/${path}`}`;
}

export function getX2tBaseUrl(origin: string): string {
  return resolveSiteUrl(origin, `${getStaticResource().x2t.root}/`);
}

// ── 文档类型 ────────────────────────────────────────────────────

/** x2t / 编辑器入参用的三类主格式（大写） */
export const FILE_TYPE = {
  DOCX: "DOCX",
  XLSX: "XLSX",
  PPTX: "PPTX",
} as const;

export type FileType = (typeof FILE_TYPE)[keyof typeof FILE_TYPE];

/** 与 SDK AppType 数值一致，仅内部映射用 */
const AppType = {
  word: 1,
  slide: 3,
  cell: 2,
  draw: 4,
  pdf: 5,
} as const;

const docTypeMap: Record<string, (typeof AppType)[keyof typeof AppType]> = {
  docx: AppType.word,
  doc: AppType.word,
  odt: AppType.word,
  rtf: AppType.word,
  txt: AppType.word,
  html: AppType.word,
  mht: AppType.word,
  epub: AppType.word,
  fb2: AppType.word,
  mobi: AppType.word,
  docm: AppType.word,
  dotx: AppType.word,
  dotm: AppType.word,
  oform: AppType.word,
  docxf: AppType.word,
  pptx: AppType.slide,
  ppt: AppType.slide,
  odp: AppType.slide,
  ppsx: AppType.slide,
  pptm: AppType.slide,
  ppsm: AppType.slide,
  potx: AppType.slide,
  potm: AppType.slide,
  otp: AppType.slide,
  odg: AppType.slide,
  xlsx: AppType.cell,
  xls: AppType.cell,
  ods: AppType.cell,
  csv: AppType.cell,
  xlsm: AppType.cell,
  xltx: AppType.cell,
  xltm: AppType.cell,
  xlsb: AppType.cell,
  ots: AppType.cell,
  vsdx: AppType.draw,
  vssx: AppType.draw,
  vstx: AppType.draw,
  vsdm: AppType.draw,
  vssm: AppType.draw,
  vstm: AppType.draw,
  pdf: AppType.pdf,
};

const appTypeName: Record<number, DocumentType> = {
  [AppType.word]: DocumentType.Word,
  [AppType.slide]: DocumentType.Slide,
  [AppType.cell]: DocumentType.Cell,
  [AppType.draw]: DocumentType.Draw,
  [AppType.pdf]: DocumentType.Pdf,
};

export function getDocumentType(ext: string) {
  const code = docTypeMap[ext.toLowerCase()];
  if (code === undefined) {
    return DocumentType.Word;
  }
  return appTypeName[code] ?? DocumentType.Word;
}

/** 新建文档页 URL */
export function getNewUrl(type: string) {
  return `/editor?new=${type}`;
}
