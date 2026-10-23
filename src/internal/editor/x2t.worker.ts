/// <reference lib="webworker" />

import { AvsFileType, X2tConvertParams, X2tConvertResult } from "./types";
import { loadX2tPdfFonts } from "./x2t-assets";
import {
  fetchMaybeBrotliAsset,
  fetchMaybeBrotliScript,
} from "./x2t-assets";
import { getStaticResource, resolveSiteUrl } from "../../const";

/**
 * @description x2t 转换 worker，在后台线程执行文档转换，避免阻塞主界面。
 */

/* eslint-disable no-restricted-globals */

/**
 * @description Worker 内使用当前 origin 拼接 x2t 静态资源绝对地址。
 */
const X2T_ORIGIN = self.location.origin;

let x2t: any = null;
let initPromise: Promise<void> | null = null;
let initResourceKey = "";
const transparentPng = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  ),
  (value) => value.charCodeAt(0),
);

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(data: Uint8Array) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function readU16(data: Uint8Array, offset: number) {
  return data[offset] | (data[offset + 1] << 8);
}

function readU32(data: Uint8Array, offset: number) {
  return (
    data[offset] |
    (data[offset + 1] << 8) |
    (data[offset + 2] << 16) |
    (data[offset + 3] << 24)
  ) >>> 0;
}

function writeU16(data: Uint8Array, offset: number, value: number) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
}

function writeU32(data: Uint8Array, offset: number, value: number) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
  data[offset + 2] = (value >>> 16) & 0xff;
  data[offset + 3] = (value >>> 24) & 0xff;
}

function concatBytes(parts: Uint8Array[]) {
  const size = parts.reduce((total, part) => total + part.length, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

async function decompressDeflateRaw(data: Uint8Array) {
  if (!("DecompressionStream" in self)) {
    return null;
  }

  const stream = new Blob([data as Uint8Array<ArrayBuffer>])
    .stream()
    .pipeThrough(new (self as any).DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function shouldRewriteZipEntry(name: string) {
  return (
    /\.(xml|rels)$/i.test(name) &&
    (name.startsWith("word/") ||
      name.startsWith("xl/") ||
      name.startsWith("ppt/"))
  );
}

function shouldRewriteDocxBookmarkEntry(name: string) {
  return /^word\/.+\.xml$/i.test(name);
}

function rewriteFontText(text: string, aliases?: Record<string, string>) {
  if (!aliases) return text;

  const entries = Object.entries(aliases)
    .filter(([from, to]) => from && to && from !== to)
    .sort(([a], [b]) => b.length - a.length);

  let output = text;
  for (const [from, to] of entries) {
    const escaped = from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    /**
     * @description 避免「仿宋」替换命中已是「仿宋_GB2312」的前缀，产生 _GB2312_GB2312。
     */
    const suffix =
      to.startsWith(from) && to.length > from.length
        ? to.slice(from.length).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
        : "";
    const pattern = suffix ? `${escaped}(?!${suffix})` : escaped;
    output = output.replace(new RegExp(pattern, "gi"), to);
  }
  return output;
}

function hasOfficeZipExtension(fileName: string) {
  return /\.(docx|xlsx|pptx)$/i.test(fileName);
}

function hasDocxExtension(fileName: string) {
  return /\.docx$/i.test(fileName);
}

function isDecimalBookmarkId(value: string) {
  return /^\d+$/.test(value);
}

function normalizeDocxBookmarkIds(text: string) {
  const bookmarkTagPattern =
    /<w:bookmark(Start|End)\b[^>]*\bw:id="([^"]*)"[^>]*>/g;
  const usedNumericIds = new Set<string>();
  let maxNumericId = -1;
  let hasInvalidId = false;

  text.replace(bookmarkTagPattern, (_tag, _kind, id: string) => {
    if (isDecimalBookmarkId(id)) {
      usedNumericIds.add(id);
      maxNumericId = Math.max(maxNumericId, Number(id));
    } else {
      hasInvalidId = true;
    }
    return _tag;
  });

  if (!hasInvalidId) return text;

  const remappedIds = new Map<string, string>();
  const nextId = () => {
    do {
      maxNumericId += 1;
    } while (usedNumericIds.has(String(maxNumericId)));

    const value = String(maxNumericId);
    usedNumericIds.add(value);
    return value;
  };

  return text.replace(bookmarkTagPattern, (tag, _kind, id: string) => {
    if (isDecimalBookmarkId(id)) return tag;

    let mapped = remappedIds.get(id);
    if (!mapped) {
      mapped = nextId();
      remappedIds.set(id, mapped);
    }
    return tag.replace(/\bw:id="[^"]*"/, `w:id="${mapped}"`);
  });
}

function encodeUtf16Le(value: string) {
  const output = new Uint8Array(value.length * 2);
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    output[i * 2] = code & 0xff;
    output[i * 2 + 1] = code >>> 8;
  }
  return output;
}

function replaceBytesWithPadding(
  data: Uint8Array,
  from: Uint8Array,
  to: Uint8Array,
  skipSuffix?: Uint8Array,
) {
  if (!from.length || to.length > from.length) return data;

  const output = data.slice();
  for (let i = 0; i <= output.length - from.length; i++) {
    let matched = true;
    for (let j = 0; j < from.length; j++) {
      if (output[i + j] !== from[j]) {
        matched = false;
        break;
      }
    }
    if (!matched) continue;

    if (skipSuffix && skipSuffix.length) {
      let suffixMatched = true;
      for (let k = 0; k < skipSuffix.length; k++) {
        if (output[i + from.length + k] !== skipSuffix[k]) {
          suffixMatched = false;
          break;
        }
      }
      if (suffixMatched) continue;
    }

    output.fill(0, i, i + from.length);
    output.set(to, i);
    i += from.length - 1;
  }

  return output;
}

function restoreEditorBinFontNames(
  data: Uint8Array,
  restoreAliases?: Record<string, string>,
) {
  if (!restoreAliases) return data;

  const entries = Object.entries(restoreAliases).sort(
    (a, b) => b[0].length - a[0].length,
  );

  let output = data;
  for (const [from, to] of entries) {
    const skipSuffix =
      to.startsWith(from) && to.length > from.length
        ? encodeUtf16Le(to.slice(from.length))
        : undefined;
    output = replaceBytesWithPadding(
      output,
      encodeUtf16Le(from),
      encodeUtf16Le(to),
      skipSuffix,
    );
  }
  return output;
}

async function rewriteOfficeZipText(
  data: ArrayBuffer,
  shouldRewriteEntry: (name: string) => boolean,
  rewriteText: (text: string, name: string) => string,
) {
  if (!("TextDecoder" in self) || !("TextEncoder" in self)) {
    return data;
  }

  const input = new Uint8Array(data);
  let eocd = -1;
  for (let i = input.length - 22; i >= Math.max(0, input.length - 65558); i--) {
    if (readU32(input, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return data;

  const entryCount = readU16(input, eocd + 10);
  const centralOffset = readU32(input, eocd + 16);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let readOffset = centralOffset;
  let outputOffset = 0;
  let changed = false;

  for (let i = 0; i < entryCount; i++) {
    if (readU32(input, readOffset) !== 0x02014b50) return data;

    const method = readU16(input, readOffset + 10);
    const modTime = readU16(input, readOffset + 12);
    const modDate = readU16(input, readOffset + 14);
    const crc = readU32(input, readOffset + 16);
    const compressedSize = readU32(input, readOffset + 20);
    const uncompressedSize = readU32(input, readOffset + 24);
    const nameLength = readU16(input, readOffset + 28);
    const extraLength = readU16(input, readOffset + 30);
    const commentLength = readU16(input, readOffset + 32);
    const internalAttrs = readU16(input, readOffset + 36);
    const externalAttrs = readU32(input, readOffset + 38);
    const localOffset = readU32(input, readOffset + 42);
    const nameBytes = input.slice(readOffset + 46, readOffset + 46 + nameLength);
    const name = decoder.decode(nameBytes);

    const localNameLength = readU16(input, localOffset + 26);
    const localExtraLength = readU16(input, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = input.slice(dataStart, dataStart + compressedSize);
    let outputData = compressed;
    let outputMethod = method;
    let outputCrc = crc;
    let outputUncompressedSize = uncompressedSize;

    if (shouldRewriteEntry(name)) {
      const uncompressed =
        method === 0
          ? compressed
          : method === 8
            ? await decompressDeflateRaw(compressed)
            : null;
      if (uncompressed) {
        const originalText = decoder.decode(uncompressed);
        const rewrittenText = rewriteText(originalText, name);
        if (rewrittenText !== originalText) {
          const rewritten = encoder.encode(rewrittenText);
          outputData = rewritten;
          outputMethod = 0;
          outputCrc = crc32(rewritten);
          outputUncompressedSize = rewritten.length;
          changed = true;
        }
      }
    }

    const localHeader = new Uint8Array(30 + nameBytes.length);
    writeU32(localHeader, 0, 0x04034b50);
    writeU16(localHeader, 4, 20);
    writeU16(localHeader, 6, 0);
    writeU16(localHeader, 8, outputMethod);
    writeU16(localHeader, 10, modTime);
    writeU16(localHeader, 12, modDate);
    writeU32(localHeader, 14, outputCrc);
    writeU32(localHeader, 18, outputData.length);
    writeU32(localHeader, 22, outputUncompressedSize);
    writeU16(localHeader, 26, nameBytes.length);
    writeU16(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, outputData);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    writeU32(centralHeader, 0, 0x02014b50);
    writeU16(centralHeader, 4, 20);
    writeU16(centralHeader, 6, 20);
    writeU16(centralHeader, 8, 0);
    writeU16(centralHeader, 10, outputMethod);
    writeU16(centralHeader, 12, modTime);
    writeU16(centralHeader, 14, modDate);
    writeU32(centralHeader, 16, outputCrc);
    writeU32(centralHeader, 20, outputData.length);
    writeU32(centralHeader, 24, outputUncompressedSize);
    writeU16(centralHeader, 28, nameBytes.length);
    writeU16(centralHeader, 30, 0);
    writeU16(centralHeader, 32, 0);
    writeU16(centralHeader, 34, 0);
    writeU16(centralHeader, 36, internalAttrs);
    writeU32(centralHeader, 38, externalAttrs);
    writeU32(centralHeader, 42, outputOffset);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    outputOffset += localHeader.length + outputData.length;
    readOffset += 46 + nameLength + extraLength + commentLength;
  }

  const centralDirectory = concatBytes(centralParts);
  const end = new Uint8Array(22);
  writeU32(end, 0, 0x06054b50);
  writeU16(end, 8, entryCount);
  writeU16(end, 10, entryCount);
  writeU32(end, 12, centralDirectory.length);
  writeU32(end, 16, outputOffset);

  return changed ? concatBytes([...localParts, centralDirectory, end]).buffer : data;
}

async function rewriteZipFontNames(
  data: ArrayBuffer,
  fileFrom: string,
  fontAliases?: Record<string, string>,
) {
  if (!fontAliases || !hasOfficeZipExtension(fileFrom)) {
    return data;
  }

  return rewriteOfficeZipText(data, shouldRewriteZipEntry, (text) =>
    rewriteFontText(text, fontAliases),
  );
}

async function sanitizeDocxForX2t(data: ArrayBuffer, fileFrom: string) {
  if (!hasDocxExtension(fileFrom)) {
    return data;
  }

  return rewriteOfficeZipText(
    data,
    shouldRewriteDocxBookmarkEntry,
    normalizeDocxBookmarkIds,
  );
}

/**
 * 上游 x2t.js 把 libheif 留成 `abort("missing function")` 桩。
 * 旧版 ppt 里的 wmf/dib 会走这条探测，一调用就把整个 wasm 打停，
 * 编辑器只弹「未知错误」，后面的页也出不来。
 * 桩改成返回 heif_error_Unsupported_filetype（3），这一张图跳过，其余照常打开。
 * 已经换过桩的脚本里找不到原来的 abort，这里就原样返回。
 */
function softenMissingHeifStubs(scriptSource: string): string {
  const writeError =
    "if(e){HEAP32[e>>2]=3;HEAP32[e+4>>2]=0;HEAP32[e+8>>2]=0;}";
  const sret = [
    "_heif_context_encode_image",
    "_heif_context_get_encoder_for_format",
    "_heif_context_get_primary_image_handle",
    "_heif_context_read_from_file",
    "_heif_context_read_from_memory_without_copy",
    "_heif_context_write_to_file",
    "_heif_decode_image",
    "_heif_image_add_plane",
    "_heif_image_create",
  ];
  const ret0 = [
    "_heif_context_alloc",
    "_heif_image_get_plane",
    "_heif_image_get_plane_readonly",
    "_heif_image_get_primary_height",
    "_heif_image_get_primary_width",
  ];
  const noop = [
    "_heif_context_free",
    "_heif_encoder_release",
    "_heif_image_handle_release",
    "_heif_image_release",
  ];
  let out = scriptSource;
  const replaceStub = (name: string, args: string, body: string) => {
    const old = `function ${name}(){abort("missing function: ${name.slice(1)}")}`;
    if (!out.includes(old)) return;
    out = out.replace(old, `function ${name}(${args}){${body}}`);
  };
  for (const name of sret) replaceStub(name, "e", writeError);
  for (const name of ret0) replaceStub(name, "", "return 0");
  for (const name of noop) replaceStub(name, "", "");
  return out;
}

/**
 * @description 在 Worker 全局作用域执行 Emscripten 脚本；classic worker 用 importScripts，module worker 用间接 eval。
 */
function executeEmscriptenScript(scriptSource: string): void {
  if (typeof importScripts === "function") {
    const scriptBlob = new Blob([scriptSource], {
      type: "application/javascript",
    });
    const scriptBlobUrl = URL.createObjectURL(scriptBlob);
    try {
      importScripts(scriptBlobUrl);
      return;
    } catch (error) {
      const isImportScriptsUnsupported =
        error instanceof TypeError &&
        String(error.message).includes("importScripts");
      if (!isImportScriptsUnsupported) {
        throw error;
      }
    } finally {
      URL.revokeObjectURL(scriptBlobUrl);
    }
  }

  /**
   * @description Module worker 不支持 importScripts，因此使用间接 eval 在 Worker 全局作用域执行。
   */
  (0, eval)(scriptSource);
}

/**
 * @description 在 Worker 上下文初始化 x2t 模块。
 */
function getWorkerStaticResource(params?: X2tConvertParams) {
  return params?.staticResource ?? getStaticResource();
}

async function initX2t(params?: X2tConvertParams): Promise<void> {
  if (x2t) return;

  const staticResource = getWorkerStaticResource(params);
  const x2tBaseUrl = resolveSiteUrl(X2T_ORIGIN, `${staticResource.x2t.root}/`);
  const scriptUrl = resolveSiteUrl(X2T_ORIGIN, staticResource.x2t.script);
  const wasmUrl = resolveSiteUrl(X2T_ORIGIN, staticResource.x2t.wasm);

  const [scriptSource, wasmBinary] = await Promise.all([
    fetchMaybeBrotliScript(scriptUrl),
    fetchMaybeBrotliAsset(wasmUrl),
  ]);

  Object.assign(self, {
    __filename: x2tBaseUrl,
    wasmBinary,
  });

  executeEmscriptenScript(softenMissingHeifStubs(scriptSource));

  x2t = (self as any).Module;

  await new Promise<void>((resolve) => {
    x2t.onRuntimeInitialized = () => resolve();
  });

  try {
    x2t.FS.mkdir("/working");
    x2t.FS.mkdir("/working/media");
    x2t.FS.mkdir("/working/fonts");
    x2t.FS.mkdir("/working/themes");
  } catch (err) {
    console.error("[x2t.worker] mkdir error:", err);
  }

  console.log("[x2t.worker] Initialized successfully");
}

/**
 * @description 转换前确保 x2t 已完成初始化。
 */
async function ensureInit(params?: X2tConvertParams): Promise<void> {
  const staticResource = getWorkerStaticResource(params);
  const resourceKey = JSON.stringify(staticResource.x2t);
  if (x2t && initResourceKey && initResourceKey !== resourceKey) {
    throw new Error("x2t static resource changed after worker initialization");
  }
  initResourceKey = resourceKey;
  if (!initPromise) {
    initPromise = initX2t(params);
  }
  return initPromise;
}

/**
 * @description 转换完成后清理临时文件。
 */
function cleanupFiles(files: string[]): void {
  for (const file of files) {
    try {
      x2t.FS.unlink(file);
    } catch (err) {
      console.error(err);
    }
  }
  cleanMedia();
  cleanFonts();
  cleanThemes();
}

type SerializedWorkerError = {
  name: string;
  message?: string;
  errno?: unknown;
  code?: unknown;
  path?: unknown;
  text: string;
};

type X2tWorkerErrorDetails = {
  stage: string;
  fileFrom?: string;
  fileTo?: string;
  readOutputError?: SerializedWorkerError;
};

class X2tWorkerError extends Error {
  constructor(
    message: string,
    public readonly details: X2tWorkerErrorDetails,
  ) {
    super(message);
    this.name = "X2tWorkerError";
  }
}

function formatSerializedWorkerError(error: Omit<SerializedWorkerError, "text">) {
  const message = error.message ? `: ${error.message}` : "";
  const details = [
    ["errno", error.errno],
    ["code", error.code],
    ["path", error.path],
  ]
    .filter(([, value]) => value != null)
    .map(([key, value]) => `${key}=${String(value)}`)
    .join(", ");
  return details
    ? `${error.name}${message} (${details})`
    : `${error.name}${message}`;
}

function serializeWorkerError(error: unknown): SerializedWorkerError {
  if (error instanceof Error) {
    const serialized = {
      name: error.name || "Error",
      message: error.message,
    };
    return {
      ...serialized,
      text: formatSerializedWorkerError(serialized),
    };
  }

  if (error && typeof error === "object") {
    const record = error as Record<string, unknown>;
    const constructorName =
      "constructor" in record &&
      record.constructor &&
      typeof record.constructor === "function"
        ? record.constructor.name
        : "";
    const name =
      typeof record.name === "string" && record.name
        ? record.name
        : constructorName || Object.prototype.toString.call(error);
    const message =
      typeof record.message === "string" && record.message
        ? record.message
        : undefined;
    const serialized = {
      name,
      message,
      errno: record.errno,
      code: record.code,
      path: record.path,
    };
    return {
      ...serialized,
      text: formatSerializedWorkerError(serialized),
    };
  }

  const text = String(error);
  return {
    name: "Error",
    message: text,
    text,
  };
}

function cleanMedia() {
  try {
    const mediaFiles = x2t.FS.readdir("/working/media/");
    for (const file of mediaFiles) {
      if (file !== "." && file !== "..") {
        x2t.FS.unlink("/working/media/" + file);
      }
    }
  } catch (err) {
    console.error(err);
  }
}

function cleanFonts() {
  try {
    const fontFiles = x2t.FS.readdir("/working/fonts/");
    for (const file of fontFiles) {
      if (file !== "." && file !== "..") {
        x2t.FS.unlink("/working/fonts/" + file);
      }
    }
  } catch (err) {
    console.error(err);
  }
}

function cleanThemes() {
  cleanDirectoryFiles("/working/themes");
}

function cleanDirectoryFiles(dir: string) {
  try {
    const files = x2t.FS.readdir(dir);
    for (const file of files) {
      if (file === "." || file === "..") {
        continue;
      }

      const path = `${dir}/${file}`;
      const stat = x2t.FS.stat(path);
      if (x2t.FS.isDir(stat.mode)) {
        cleanDirectoryFiles(path);
        x2t.FS.rmdir(path);
      } else {
        x2t.FS.unlink(path);
      }
    }
  } catch (err) {
    console.error(err);
  }
}

function ensureDirectory(path: string) {
  const parts = path.split("/").filter(Boolean);
  let current = "";

  for (const part of parts) {
    current += "/" + part;
    try {
      if (!x2t.FS.analyzePath(current).exists) {
        x2t.FS.mkdir(current);
      }
    } catch (err) {
      console.error(err);
    }
  }
}

function writeFileMap(files: { [key: string]: Uint8Array }) {
  for (const [key, value] of Object.entries(files)) {
    try {
      const path = "/working/" + key.replace(/^\/+/, "");
      ensureDirectory(path.split("/").slice(0, -1).join("/"));
      x2t.FS.writeFile(path, value);
    } catch (err) {
      console.error(key, err);
    }
  }
}

function writeFonts(fonts?: { [key: string]: Uint8Array }) {
  cleanFonts();
  if (!fonts) {
    return;
  }

  for (const [key, value] of Object.entries(fonts)) {
    try {
      x2t.FS.writeFile("/working/fonts/" + key, value);
    } catch (err) {
      console.error(key, err);
    }
  }
}

function writePdfBin(pdfBin?: Uint8Array) {
  const pdfPath = "/working/pdf.bin";
  try {
    if (x2t.FS.analyzePath(pdfPath).exists) {
      x2t.FS.unlink(pdfPath);
    }
  } catch (err) {
    console.error(err);
  }

  if (pdfBin?.byteLength) {
    x2t.FS.writeFile(pdfPath, pdfBin);
  }
}

/**
 * @description 从工作目录读取转换生成的媒体文件。
 */
function readMedia(): { [key: string]: Uint8Array } {
  const media: { [key: string]: Uint8Array } = {};
  try {
    const files = x2t.FS.readdir("/working/media/");
    for (const file of files) {
      if (file !== "." && file !== "..") {
        const fileData = x2t.FS.readFile("/working/media/" + file, {
          encoding: "binary",
        });
        media[file] = fileData;
      }
    }
  } catch (e) {
    console.error(e);
  }
  return media;
}

function readDirectoryFiles(root: string, prefix: string) {
  const files: { [key: string]: Uint8Array } = {};

  const visit = (dir: string, relDir: string) => {
    try {
      for (const file of x2t.FS.readdir(dir)) {
        if (file === "." || file === "..") {
          continue;
        }

        const path = `${dir}/${file}`;
        const rel = relDir ? `${relDir}/${file}` : file;
        const stat = x2t.FS.stat(path);
        if (x2t.FS.isDir(stat.mode)) {
          visit(path, rel);
          continue;
        }

        files[`${prefix}/${rel}`] = x2t.FS.readFile(path, {
          encoding: "binary",
        });
      }
    } catch (err) {
      console.error(err);
    }
  };

  visit(root, "");
  return files;
}

const xmlPath = "/working/params.xml";

function writeInputs({
  fileFrom,
  fileTo,
  formatFrom,
  formatTo,
  data,
  media,
  pdfBin,
  themes,
  csvEncoding,
  csvDelimiter,
  csvDelimiterChar,
}: X2tConvertParams) {
  const isCsvSource =
    formatFrom === AvsFileType.AVS_FILE_SPREADSHEET_CSV ||
    fileFrom.toLowerCase().endsWith(".csv");
  const usePdfBinPath = Boolean(pdfBin?.byteLength);

  const params: Record<string, string | number | boolean> = {
    m_sFileFrom: fileFrom,
    m_sThemeDir: "/working/themes",
    m_sFileTo: fileTo,
    m_bIsPDFA: formatTo === AvsFileType.AVS_FILE_CROSSPLATFORM_PDFA,
    m_bIsNoBase64: usePdfBinPath ? false : true,
    m_sFontDir: "/working/fonts/",
  };

  if (!usePdfBinPath) {
    params.m_nFormatFrom = formatFrom ?? 0;
    params.m_nFormatTo = formatTo ?? 0;
  }

  if (isCsvSource) {
    params.m_nCsvTxtEncoding = csvEncoding ?? 46;
    params.m_nCsvDelimiter = csvDelimiter ?? 4;
    if (csvDelimiterChar) {
      params.m_nCsvDelimiterChar = csvDelimiterChar;
    }
  }

  const content = Object.entries(params)
    .filter(([, v]) => v !== undefined && v !== null && v !== "")
    .reduce((a, [k, v]) => a + `<${k}>${v}</${k}>\n`, "");

  const xml = `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
>
${content}
</TaskQueueDataConvert>`;

  x2t.FS.writeFile(xmlPath, xml);
  if (data) {
    x2t.FS.writeFile(fileFrom, new Uint8Array(data));
  }

  if (media) {
    cleanMedia();
    for (const [key, value] of Object.entries(media)) {
      try {
        const data = /\.(emf|svg|webp)$/i.test(key) ? transparentPng : value;
        x2t.FS.writeFile("/working/" + key, data);
      } catch (err) {
        console.error(key, err);
      }
    }
  }

  if (themes) {
    cleanThemes();
    writeFileMap(themes);
  }
}

/**
 * @description 将文档从一种格式转换为另一种格式。
 */
async function convert({
  data,
  fileFrom,
  fileTo,
  formatFrom,
  formatTo,
  media,
  pdfBin,
  fonts,
  fontAliases,
  fontExportAliases,
  themes,
  csvEncoding,
  csvDelimiter,
  csvDelimiterChar,
  staticResource,
}: X2tConvertParams): Promise<X2tConvertResult> {
  let preparedData = data;
  if (preparedData) {
    preparedData = await sanitizeDocxForX2t(preparedData, fileFrom);
    preparedData = await rewriteZipFontNames(preparedData, fileFrom, fontAliases);
  }
  const fromPath = "/working/" + fileFrom;
  const toPath = "/working/" + fileTo;
  const files = [fromPath, toPath, xmlPath];
  if (pdfBin?.byteLength) {
    files.push("/working/pdf.bin");
  }

  const needsPdfFonts =
    Boolean(pdfBin?.byteLength) ||
    formatTo === AvsFileType.AVS_FILE_CROSSPLATFORM_PDF ||
    formatTo === AvsFileType.AVS_FILE_CROSSPLATFORM_PDFA;

  let allFonts = fonts;
  if (needsPdfFonts) {
    const pdfFontsRoot = getWorkerStaticResource({
      staticResource,
    } as X2tConvertParams).x2t.pdfFonts.root;
    const pdfFonts = await loadX2tPdfFonts(
      self.location.origin,
      pdfFontsRoot,
    );
    allFonts = { ...pdfFonts, ...fonts };
  }

  writeFonts(allFonts);
  writePdfBin(pdfBin);

  writeInputs({
    fileFrom: fromPath,
    fileTo: toPath,
    formatFrom,
    formatTo,
    data: preparedData,
    media,
    themes,
    pdfBin,
    csvEncoding,
    csvDelimiter,
    csvDelimiterChar,
  });

  if (
    fileFrom.endsWith(".doc") ||
    formatFrom == AvsFileType.AVS_FILE_DOCUMENT_DOC
  ) {
    const viaPath = fromPath + ".docx";
    try {
      const pathInfo = x2t.FS.analyzePath(viaPath);
      if (pathInfo.exists) {
        x2t.FS.unlink(viaPath);
      }
    } catch (err) {}
    writeInputs({
      fileFrom: fromPath,
      fileTo: viaPath,
      formatFrom: AvsFileType.AVS_FILE_DOCUMENT_DOC,
      formatTo: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
      data: null as never,
    });
    x2t.ccall("main1", ["number"], ["string"], [xmlPath]);
    writeInputs({
      fileFrom: viaPath,
      fileTo: toPath,
      formatFrom: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
      formatTo,
      data: null as never,
    });
    files.push(viaPath);
  }

  try {
    const pathInfo = x2t.FS.analyzePath(toPath);
    if (pathInfo.exists) {
      x2t.FS.unlink(toPath);
    }
  } catch (err) {}

  try {
    x2t.ccall("main1", ["number"], ["string"], [xmlPath]);
  } catch (e) {
    console.error("ccall", e);
  }

  let output: Uint8Array | null = null;
  let readOutputError: SerializedWorkerError | undefined;
  try {
    output = x2t.FS.readFile(toPath);
    if (output && fileTo === "Editor.bin" && fontAliases) {
      output = restoreEditorBinFontNames(output, fontAliases);
    }
    if (output && hasOfficeZipExtension(fileTo) && fontExportAliases) {
      const restored = await rewriteZipFontNames(
        output.slice(0).buffer,
        fileTo,
        fontExportAliases,
      );
      output = new Uint8Array(restored);
    }
  } catch (e) {
    console.error("[x2t.worker] read output failed:", e);
    readOutputError = serializeWorkerError(e);
  }

  if (!output) {
    throw new X2tWorkerError("x2t conversion produced no output", {
      stage: "read-output",
      fileFrom,
      fileTo,
      readOutputError,
    });
  }

  const outputMedia = readMedia();
  const outputThemes = readDirectoryFiles("/working/themes", "themes");

  setTimeout(() => {
    cleanupFiles(files);
  });

  return { output, media: outputMedia, themes: outputThemes };
}

/**
 * @description 主线程发送给 x2t worker 的消息结构。
 */
interface WorkerMessage {
  id?: number;
  type: "convert";
  payload?: any;
}

/**
 * @description 处理主线程发来的 worker 消息。
 */
self.onmessage = async (event: MessageEvent<WorkerMessage>) => {
  const { id, type, payload } = event.data;

  try {
    switch (type) {
      case "convert": {
        await ensureInit(payload);
        const result = await convert(payload);

        /**
         * @description 使用 Transferable 返回大块二进制结果，降低复制成本。
         */
        const transferables: Transferable[] = [];
        if (result.output) {
          transferables.push(result.output.buffer);
        }
        Object.values(result.media).forEach((m) =>
          transferables.push(m.buffer),
        );

        self.postMessage(
          { id, type: "convert:done", payload: result },
          { transfer: transferables },
        );
        break;
      }

      default:
        self.postMessage({
          id,
          type: "error",
          error: `Unknown message type: ${type}`,
        });
    }
  } catch (error) {
    self.postMessage({
      id,
      type: "error",
      error: error instanceof Error ? error.message : String(error),
      errorDetails:
        error instanceof X2tWorkerError ? error.details : undefined,
    });
  }
};

/**
 * @description 通知主线程 worker 已加载。
 */
self.postMessage({ type: "ready" });
