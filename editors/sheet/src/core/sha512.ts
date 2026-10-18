/**
 * 同步的 SHA-512 / SHA-384（用于校验 Excel 新式的工作表保护密码：加盐后迭代 10 万次，WebCrypto 的异步接口太慢）。
 * 常量不手抄：K = 前 80 个素数立方根的小数部分，IV = 前 8 个素数平方根的小数部分，用 BigInt 现算（测试里和 WebCrypto 对拍）。
 */

const M64 = (1n << 64n) - 1n;

function primes(n: number): number[] {
  const out: number[] = [];
  for (let x = 2; out.length < n; x++) if (out.every((p) => x % p)) out.push(x);
  return out;
}

/** floor(x^(1/k))，x 是 BigInt */
function iroot(x: bigint, k: bigint): bigint {
  let lo = 0n;
  let hi = 1n << (BigInt(x.toString(2).length) / k + 1n);
  while (lo < hi) {
    const mid = (lo + hi + 1n) >> 1n;
    if (mid ** k <= x) lo = mid;
    else hi = mid - 1n;
  }
  return lo;
}

const P = primes(80);
// K[i] = 小数部分 × 2^64 = floor(cbrt(p) × 2^64) mod 2^64 = icbrt(p × 2^192) mod 2^64
const K = P.map((p) => iroot(BigInt(p) << 192n, 3n) & M64);
const IV512 = P.slice(0, 8).map((p) => iroot(BigInt(p) << 128n, 2n) & M64);
// SHA-384 的初始值：第 9 到第 16 个素数的平方根小数部分
const IV384 = primes(16).slice(8).map((p) => iroot(BigInt(p) << 128n, 2n) & M64);

const KH = Int32Array.from(K, (k) => Number(k >> 32n) | 0);
const KL = Int32Array.from(K, (k) => Number(k & 0xffffffffn) | 0);

function digest(data: Uint8Array, iv: bigint[], outLen: number): Uint8Array {
  const H = new Int32Array(16); // hi0 lo0 hi1 lo1 …
  iv.forEach((v, i) => {
    H[i * 2] = Number(v >> 32n) | 0;
    H[i * 2 + 1] = Number(v & 0xffffffffn) | 0;
  });
  // 填充：0x80、补 0、128 位长度（只用低 64 位）
  const total = ((data.length + 17 + 127) >> 7) << 7;
  const msg = new Uint8Array(total);
  msg.set(data);
  msg[data.length] = 0x80;
  const bits = BigInt(data.length) * 8n;
  for (let i = 0; i < 8; i++) msg[total - 1 - i] = Number((bits >> BigInt(8 * i)) & 0xffn);
  const wh = new Int32Array(80);
  const wl = new Int32Array(80);
  const dv = new DataView(msg.buffer);
  for (let off = 0; off < total; off += 128) {
    for (let t = 0; t < 16; t++) {
      wh[t] = dv.getInt32(off + t * 8);
      wl[t] = dv.getInt32(off + t * 8 + 4);
    }
    for (let t = 16; t < 80; t++) {
      const h15 = wh[t - 15], l15 = wl[t - 15], h2 = wh[t - 2], l2 = wl[t - 2];
      // s0 = rotr1 ^ rotr8 ^ shr7
      const s0h = ((h15 >>> 1) | (l15 << 31)) ^ ((h15 >>> 8) | (l15 << 24)) ^ (h15 >>> 7);
      const s0l = ((l15 >>> 1) | (h15 << 31)) ^ ((l15 >>> 8) | (h15 << 24)) ^ ((l15 >>> 7) | (h15 << 25));
      // s1 = rotr19 ^ rotr61 ^ shr6（rotr61 = rotl3 = 交换后 rotr29）
      const s1h = ((h2 >>> 19) | (l2 << 13)) ^ ((l2 >>> 29) | (h2 << 3)) ^ (h2 >>> 6);
      const s1l = ((l2 >>> 19) | (h2 << 13)) ^ ((h2 >>> 29) | (l2 << 3)) ^ ((l2 >>> 6) | (h2 << 26));
      const lo = (wl[t - 16] >>> 0) + (s0l >>> 0) + (wl[t - 7] >>> 0) + (s1l >>> 0);
      wl[t] = lo | 0;
      wh[t] = (wh[t - 16] + s0h + wh[t - 7] + s1h + Math.floor(lo / 4294967296)) | 0;
    }
    let ah = H[0], al = H[1], bh = H[2], bl = H[3], ch = H[4], cl = H[5], dh = H[6], dl = H[7];
    let eh = H[8], el = H[9], fh = H[10], fl = H[11], gh = H[12], gl = H[13], hh = H[14], hl = H[15];
    for (let t = 0; t < 80; t++) {
      // S1(e) = rotr14 ^ rotr18 ^ rotr41（41 = 32 + 9）
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh);
      const chl = (el & fl) ^ (~el & gl);
      const t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (KL[t] >>> 0) + (wl[t] >>> 0);
      const t1h = (hh + S1h + chh + KH[t] + wh[t] + Math.floor(t1l / 4294967296)) | 0;
      // S0(a) = rotr28 ^ rotr34 ^ rotr39（34 = 32 + 2，39 = 32 + 7）
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch);
      const mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
      const t2l = (S0l >>> 0) + (mjl >>> 0);
      const t2h = (S0h + mjh + Math.floor(t2l / 4294967296)) | 0;
      hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
      const nel = (dl >>> 0) + (t1l >>> 0);
      el = nel | 0;
      eh = (dh + t1h + Math.floor(nel / 4294967296)) | 0;
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      const nal = (t1l >>> 0) + (t2l >>> 0);
      al = nal | 0;
      ah = (t1h + t2h + Math.floor(nal / 4294967296)) | 0;
    }
    const add = (i: number, h: number, l: number) => {
      const lo = (H[i * 2 + 1] >>> 0) + (l >>> 0);
      H[i * 2 + 1] = lo | 0;
      H[i * 2] = (H[i * 2] + h + Math.floor(lo / 4294967296)) | 0;
    };
    add(0, ah, al); add(1, bh, bl); add(2, ch, cl); add(3, dh, dl);
    add(4, eh, el); add(5, fh, fl); add(6, gh, gl); add(7, hh, hl);
  }
  const out = new Uint8Array(64);
  const odv = new DataView(out.buffer);
  for (let i = 0; i < 16; i++) odv.setInt32(i * 4, H[i]);
  return out.slice(0, outLen);
}

export const sha512 = (data: Uint8Array) => digest(data, IV512, 64);
export const sha384 = (data: Uint8Array) => digest(data, IV384, 48);
