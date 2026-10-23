#!/usr/bin/env node
/**
 * 取格式转换引擎（x2t）的 WebAssembly 产物。
 *
 *   node scripts/fetch-x2t.mjs           # 取到 vendor/x2t/
 *   node scripts/fetch-x2t.mjs --force   # 已经有了也重取
 *
 * ── 这份产物是从哪来的 ──────────────────────────────────────────────────────
 *
 * 官方没有出过 WebAssembly 版的 x2t（core 那边的请求一直开着）。能用的只有社区配方：
 * CryptPad 的 `onlyoffice-x2t-wasm`——一个 Dockerfile 把 ONLYOFFICE core 用
 * emscripten 分三十来个阶段编出来，最后链成一个 wasm 再用 brotli 压一遍。
 *
 * **上游那个组件包里的 x2t 就是这条配方的产物**，不是什么没有源码的黑盒：
 * 它的启动脚本与 CryptPad 仓库里的 `pre-js.js` 逐字一致（连制表符缩进都一样），
 * 并且导出了 `_main1`——那个符号只可能来自 CryptPad 的 `wrap-main.cpp`。
 *
 * 上游那份 9.67 MB，我们这份 6.8 MB，**小 30% 而功能一样**：
 * 上游那趟编译带着 26.5 MB 调试信息、而且没开优化（函数个数 146064 对 76562）。
 *
 * ⚠ **它基于的 core 是 9.3.0.140，而我们的文档服务是 9.4.0.129，差一档。**
 * 等本机有了 docker，`build/x2t/` 那份配方会编一份对齐 9.4.0.129 的顶掉它。
 *
 * ── 判据 ────────────────────────────────────────────────────────────────────
 *
 * 取回来必须对上官方随发布件给的 sha512，对不上就停。这是「拿到的确实是那一份」
 * 的唯一判据——文件名和大小都能对而内容不同。
 */

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

import { X2T } from "../config.mjs";
import { decompressWasm, patchX2tNullCallback } from "./patch-x2t-null-callback.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DEST = path.join(ROOT, "vendor/x2t");

const FORCE = process.argv.includes("--force");
const log = (...a) => console.log("→", ...a);
const die = (m) => {
  console.error("✗ " + m);
  process.exit(1);
};

/** 极简 zip 解包：走中央目录，只认存储与 deflate 两种。 */
function unzip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) die("这不是一个 zip");
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const out = new Map();
  for (let k = 0; k < count; k++) {
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOff = buf.readUInt32LE(p + 42);
    const name = buf.toString("utf8", p + 46, p + 46 + nameLen);
    const method = buf.readUInt16LE(p + 10);
    const csize = buf.readUInt32LE(p + 20);
    const lNameLen = buf.readUInt16LE(localOff + 26);
    const lExtraLen = buf.readUInt16LE(localOff + 28);
    const start = localOff + 30 + lNameLen + lExtraLen;
    let data = buf.subarray(start, start + csize);
    if (method === 8) data = zlib.inflateRawSync(data);
    else if (method !== 0) die(`zip 里 ${name} 用了不认识的压缩方式 ${method}`);
    out.set(name, data);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

/** 数一份 wasm 里各段多大，用来说清「这份产物长什么样」。 */
function describeWasm(raw) {
  let b = raw;
  let brotli = false;
  if (b.readUInt32LE(0) !== 0x6d736100) {
    b = zlib.brotliDecompressSync(b);
    brotli = true;
  }
  if (b.readUInt32LE(0) !== 0x6d736100) die("解开之后还不是 wasm");
  let o = 8;
  const lebu = () => {
    let r = 0;
    let s = 0;
    let by;
    do {
      by = b[o++];
      r |= (by & 0x7f) << s;
      s += 7;
    } while (by & 0x80);
    return r >>> 0;
  };
  let debugBytes = 0;
  let funcCount = 0;
  while (o < b.length) {
    const id = b[o++];
    const size = lebu();
    const end = o + size;
    if (id === 0) {
      const save = o;
      const nlen = lebu();
      const nm = b.toString("utf8", o, o + nlen);
      if (nm === "name" || nm.startsWith(".debug")) debugBytes += size;
      o = save + size;
    } else {
      if (id === 3) {
        const save = o;
        funcCount = lebu();
        o = save;
      }
      o = end;
    }
  }
  return { 压缩: brotli ? "brotli" : "无", 压缩后字节: raw.length, 解开后字节: b.length, 函数个数: funcCount, 调试信息字节: debugBytes };
}

async function main() {
  if (fs.existsSync(DEST) && !FORCE) die(`${DEST} 已经有了。要重取加 --force。`);

  log(`下载 ${X2T.url}`);
  const res = await fetch(X2T.url, { redirect: "follow" });
  if (!res.ok) die(`下载失败：HTTP ${res.status}`);
  const zip = Buffer.from(await res.arrayBuffer());
  log(`  ${zip.length} 字节`);

  const got = crypto.createHash("sha512").update(zip).digest("hex");
  if (got !== X2T.sha512) {
    die(`校验和对不上，拿到的不是那一份。\n  该是：${X2T.sha512}\n  实际：${got}`);
  }
  log("  校验和对上了");

  const files = unzip(zip);
  for (const need of ["x2t.js", "x2t.wasm.br"]) {
    if (!files.has(need)) die(`发布件里没有 ${need}，拿到的东西不对`);
  }

  fs.rmSync(DEST, { recursive: true, force: true });
  fs.mkdirSync(DEST, { recursive: true });

  // 校验和通过之后才改字节。空图片回调会让 Chrome 里的旧版 ppt 停在「未知错误」。
  // 组件取回来自己认魔数：不是 wasm 就当 brotli 解。所以这里存的仍是压缩那一份，
  // 名字仍叫 x2t.wasm——省下载，而组件那边一个字都不用改。
  const patched = patchX2tNullCallback(decompressWasm(files.get("x2t.wasm.br")));
  const patchedBr = zlib.brotliCompressSync(patched, {
    params: { [zlib.constants.BROTLI_PARAM_QUALITY]: 6 },
  });
  fs.writeFileSync(path.join(DEST, "x2t.wasm"), patchedBr);
  fs.writeFileSync(path.join(DEST, "x2t.js"), files.get("x2t.js"));

  const shape = describeWasm(patchedBr);
  const source = {
    说明: "这份 x2t 是从哪来的。别手改——由 scripts/fetch-x2t.mjs 生成。",
    来源: X2T.source,
    发布版本: X2T.version,
    地址: X2T.url,
    仓库: X2T.repo,
    校验和sha512: X2T.sha512,
    基于的core版本: X2T.builtFromCore,
    取的时间: new Date().toISOString(),
    补丁: "null-image-callback：scripts/patch-x2t-null-callback.mjs，函数 62523 的空 call_indirect 改成跳过",
    产物: {
      "x2t.wasm": { ...shape, 说明: "存的是补过空回调之后再 brotli 压缩的那一份，组件自己解" },
      "x2t.js": { 字节: files.get("x2t.js").length },
    },
  };
  fs.writeFileSync(path.join(DEST, "SOURCE.json"), JSON.stringify(source, null, 2) + "\n");

  console.log("\n✓ 取好了：" + DEST);
  console.log(`  x2t.wasm  ${(shape.压缩后字节 / 1048576).toFixed(2)} MB（解开 ${(shape.解开后字节 / 1048576).toFixed(1)} MB，函数 ${shape.函数个数}，调试信息 ${shape.调试信息字节} 字节）`);
  console.log(`  x2t.js    ${(files.get("x2t.js").length / 1024).toFixed(0)} KB`);
}

main();
