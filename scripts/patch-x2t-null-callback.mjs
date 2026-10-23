#!/usr/bin/env node
/**
 * 给 CryptPad v9.3.0+0 的 x2t.wasm 补上空图片回调。
 *
 *   node scripts/patch-x2t-null-callback.mjs
 *   node scripts/patch-x2t-null-callback.mjs vendor/x2t/x2t.wasm
 *
 * 旧版 ppt 转到某一张图时，函数 62523 会 call_indirect 一个空的函数指针。
 * Node 里这次转换能走完，Chrome 会在这里陷阱成 RuntimeError: null function，
 * 编辑器只显示「未知错误」。
 *
 * 这一处把那个 call_indirect 换成：指针是 0 就丢掉两个参数，否则照原样调用。
 * 有效回调仍然执行。字节位置只对解开后 35985703 字节的那一份成立；
 * 对不上就停，不去改别的构建。已经补过的文件原样返回。
 *
 * wasm 本体不进 git。`scripts/fetch-x2t.mjs` 在校验和通过之后调用这里。
 */

import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_WASM = path.join(ROOT, "vendor/x2t/x2t.wasm");

const UPSTREAM_LEN = 35985703;
const PATCHED_LEN = 35985714;
const BODY_SIZE_AT = 21543987;
const CALL_AT = 21544122;
const OLD_BODY_LEN = 184;
const NEW_BODY_LEN = 195;

const BEFORE = Buffer.from([0x28, 0x02, 0x1c]);
const AFTER = Buffer.from([0x20, 0x01, 0x10, 0xc9]);
const CHECK = Buffer.from([
  0x22, 0x03, // local.tee 3
  0x45, // i32.eqz
  0x04, 0x01, // if，类型 1 是 (i32, i32) -> void。空类型 0x40 看不见外层参数。
  0x1a, 0x1a, // drop drop
  0x05, // else
  0x20, 0x03, // local.get 3
  0x11, 0x01, 0x00, // call_indirect 类型 1
  0x0b, // end
]);

function encodeUleb(value) {
  const bytes = [];
  let rest = value >>> 0;
  do {
    let b = rest & 0x7f;
    rest >>>= 7;
    if (rest) b |= 0x80;
    bytes.push(b);
  } while (rest);
  return Buffer.from(bytes);
}

function readUleb(buf, at) {
  let value = 0;
  let shift = 0;
  let o = at;
  let b;
  do {
    b = buf[o++];
    value |= (b & 0x7f) << shift;
    shift += 7;
  } while (b & 0x80);
  return { value: value >>> 0, start: at, len: o - at };
}

function same(buf, at, bytes) {
  return buf.subarray(at, at + bytes.length).equals(bytes);
}

/**
 * @param {Buffer} wasm 解开后的 wasm
 * @returns {Buffer}
 */
export function patchX2tNullCallback(wasm) {
  if (wasm.length === PATCHED_LEN && same(wasm, CALL_AT, CHECK)) {
    return wasm;
  }
  if (wasm.length !== UPSTREAM_LEN) {
    throw new Error(`解开后是 ${wasm.length} 字节，这份补丁只认 ${UPSTREAM_LEN} 字节的 CryptPad v9.3.0+0`);
  }
  if (wasm[BODY_SIZE_AT] !== 0xb8 || wasm[BODY_SIZE_AT + 1] !== 0x01) {
    throw new Error("函数体长度对不上，不去改");
  }
  if (!same(wasm, CALL_AT - BEFORE.length, BEFORE) || wasm[CALL_AT] !== 0x11 || wasm[CALL_AT + 1] !== 0x01 || wasm[CALL_AT + 2] !== 0x00 || !same(wasm, CALL_AT + 3, AFTER)) {
    throw new Error("call_indirect 周围的字节对不上，不去改");
  }

  let o = 8;
  while (o < wasm.length) {
    const id = wasm[o++];
    const size = readUleb(wasm, o);
    o = size.start + size.len;
    const payload = o;
    if (id !== 10) {
      o = payload + size.value;
      continue;
    }
    const extra = CHECK.length - 3;
    const grown = encodeUleb(size.value + extra);
    if (grown.length !== size.len) {
      throw new Error("代码段长度的编码变长了，补丁停在这里");
    }
    const bodyStart = BODY_SIZE_AT + 2;
    const body = Buffer.concat([
      wasm.subarray(bodyStart, CALL_AT),
      CHECK,
      wasm.subarray(CALL_AT + 3, bodyStart + OLD_BODY_LEN),
    ]);
    if (body.length !== NEW_BODY_LEN) {
      throw new Error(`补完的函数体是 ${body.length} 字节`);
    }
    return Buffer.concat([
      wasm.subarray(0, size.start),
      grown,
      wasm.subarray(payload, BODY_SIZE_AT),
      encodeUleb(NEW_BODY_LEN),
      body,
      wasm.subarray(bodyStart + OLD_BODY_LEN),
    ]);
  }
  throw new Error("没有找到代码段");
}

export function decompressWasm(raw) {
  if (raw.length >= 4 && raw.readUInt32LE(0) === 0x6d736100) return raw;
  return zlib.brotliDecompressSync(raw);
}

function main() {
  const target = path.resolve(process.argv[2] || DEFAULT_WASM);
  if (!fs.existsSync(target)) {
    console.error("✗ 没有 " + target + "。先跑 npm run x2t。");
    process.exit(1);
  }
  const raw = fs.readFileSync(target);
  const wasm = decompressWasm(raw);
  const patched = patchX2tNullCallback(wasm);
  if (patched === wasm) {
    console.log("✓ 已经补过：" + target);
    return;
  }
  const br = zlib.brotliCompressSync(patched, {
    params: { [zlib.constants.BROTLI_PARAM_QUALITY]: 6 },
  });
  fs.writeFileSync(target, br);
  console.log(`✓ 已写入 ${target}（解开 ${patched.length} 字节，压缩 ${br.length} 字节）`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
