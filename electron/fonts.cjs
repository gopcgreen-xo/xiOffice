// 系统字体的度量表（Word 编辑器按字体文件算行高）：
// 只从字体文件（.ttf / .otf / .ttc）里取出 name、hhea、OS/2、head 四张表的原始字节（base64）交给前端，
// 解析和 Word 的行高规则在前端 editors/word/src/core/fontfile.ts。只读用到的表，不整读大字体文件；结果在本次运行内缓存。
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

const TAGS = ['name', 'hhea', 'OS/2', 'head'];
/** 单张表的上限（name 表通常几 KB） */
const MAX_TABLE = 4 << 20;

function readAt(fd, pos, len) {
  const buf = Buffer.alloc(len);
  return fs.readSync(fd, buf, 0, len, pos) === len ? buf : null;
}

/** 一个字体文件里每个字体需要的表（表名 → base64） */
function facesOf(file) {
  let fd;
  try {
    fd = fs.openSync(file, 'r');
  } catch {
    return [];
  }
  try {
    const head = readAt(fd, 0, 12);
    if (!head) return [];
    let offsets = [0];
    if (head.toString('latin1', 0, 4) === 'ttcf') {
      const n = Math.min(head.readUInt32BE(8), 256);
      const b = readAt(fd, 12, n * 4);
      if (!b) return [];
      offsets = Array.from({ length: n }, (_, i) => b.readUInt32BE(i * 4));
    }
    const out = [];
    for (const off of offsets) {
      const h = readAt(fd, off, 12);
      if (!h) continue;
      const n = h.readUInt16BE(4);
      const dir = readAt(fd, off + 12, n * 16);
      if (!dir) continue;
      const tables = {};
      for (let i = 0; i < n; i++) {
        const tag = dir.toString('latin1', i * 16, i * 16 + 4);
        if (!TAGS.includes(tag)) continue;
        const start = dir.readUInt32BE(i * 16 + 8);
        const len = dir.readUInt32BE(i * 16 + 12);
        if (len > MAX_TABLE) continue;
        const data = readAt(fd, start, len);
        if (data) tables[tag] = data.toString('base64');
      }
      if (Object.keys(tables).length === TAGS.length) out.push(tables);
    }
    return out;
  } finally {
    fs.closeSync(fd);
  }
}

function fontDirs() {
  const home = os.homedir();
  const dirs = [];
  if (process.env.WINDIR) dirs.push(path.join(process.env.WINDIR, 'Fonts'));
  if (process.env.LOCALAPPDATA) dirs.push(path.join(process.env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts'));
  dirs.push('/System/Library/Fonts', '/Library/Fonts', path.join(home, 'Library', 'Fonts'));
  dirs.push('/usr/share/fonts', '/usr/local/share/fonts', path.join(home, '.local', 'share', 'fonts'), path.join(home, '.fonts'));
  return dirs;
}

function collect(dir, out, depth) {
  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const e of entries) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) {
      if (depth < 3) collect(p, out, depth + 1);
    } else if (/\.(ttf|otf|ttc)$/i.test(e.name)) out.push(...facesOf(p));
  }
}

let cache = null;
/** 所有已安装字体的度量表（第一次调用时读取，之后用缓存） */
function allFaces() {
  if (!cache) {
    cache = [];
    for (const d of fontDirs()) collect(d, cache, 0);
  }
  return cache;
}

module.exports = { allFaces, facesOf };
