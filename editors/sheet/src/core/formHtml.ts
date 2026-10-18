/**
 * 表单的离线填写页：把一个表单生成成一个自包含的 .html（不联网、不依赖本程序），发给别人填写；
 * 填完在页面里“导出填写结果”得到一个 .json，回来用表单面板的“导入填写结果…”写进表格（逐条按列属性校验）。
 * 结果格式见 docs/WBP_FORMAT.md 第 6.2 节（formresult）。
 */
import { COLUMN_TYPE_LABELS, type SelectOption } from '@wz/editor-shared';
import type { ResolvedField } from './forms';

export const FORM_RESULT_VERSION = 1;

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
/** 放进 <script type="application/json"> 的 JSON：< 转义，避免提前结束标签 */
const BS = String.fromCharCode(92);
const jsonForScript = (v: unknown): string =>
  JSON.stringify(v)
    .replace(/</g, `${BS}u003c`)
    .split(String.fromCharCode(0x2028))
    .join(`${BS}u2028`)
    .split(String.fromCharCode(0x2029))
    .join(`${BS}u2029`);

const LONG_TEXT = /备注|说明|描述|原因|意见/;

export interface FormHtmlInput {
  /** 工作包标题（页面小标题） */
  packTitle: string;
  /** 工作包 id（写进结果，导入时核对是同一个包） */
  packId: string;
  formId: string;
  title: string;
  fields: ResolvedField[];
  /** 单选 / 多选 / 复选框的选项（引用区域的已展开成文字） */
  optionsOf: (f: ResolvedField) => SelectOption[];
}

/** 页面里“能填”的字段：隐藏的、自动填的不放进去 */
export function fillableFields(fields: ResolvedField[]): ResolvedField[] {
  return fields.filter((f) => !f.hidden && !f.auto && f.meta.type !== 'autonum' && f.meta.type !== 'created');
}

export function buildFormHtml(input: FormHtmlInput): string {
  const fields = fillableFields(input.fields);
  const spec = fields.map((f) => ({
    id: f.meta.id,
    label: f.label,
    type: f.meta.type,
    required: !!f.meta.required,
    options: f.meta.type === 'select' || f.meta.type === 'multi' ? input.optionsOf(f).map((o) => o.v) : [],
    integer: !!f.meta.integer,
    min: f.meta.min ?? null,
    max: f.meta.max ?? null,
    def: f.meta.default ?? '',
  }));
  const control = (f: ResolvedField): string => {
    const id = esc(f.meta.id);
    const m = f.meta;
    switch (m.type) {
      case 'select':
        return `<select id="f-${id}" data-id="${id}"><option value=""></option>${input
          .optionsOf(f)
          .map((o) => `<option value="${esc(o.v)}">${esc(o.v)}</option>`)
          .join('')}</select>`;
      case 'multi':
        return `<div class="opts" data-id="${id}">${input
          .optionsOf(f)
          .map((o) => `<label><input type="checkbox" value="${esc(o.v)}"> ${esc(o.v)}</label>`)
          .join('')}</div>`;
      case 'check':
        return `<label class="chk"><input type="checkbox" id="f-${id}" data-id="${id}"> 是</label>`;
      case 'date':
        return `<input type="date" id="f-${id}" data-id="${id}">`;
      default:
        return LONG_TEXT.test(m.title) ? `<textarea id="f-${id}" data-id="${id}" rows="3"></textarea>` : `<input type="text" id="f-${id}" data-id="${id}"${m.type === 'number' || m.type === 'money' || m.type === 'percent' ? ' inputmode="decimal"' : ''}>`;
    }
  };
  const rows = fields
    .map(
      (f) =>
        `<div class="field" data-field="${esc(f.meta.id)}"><label for="f-${esc(f.meta.id)}">${esc(f.label)}${f.meta.required ? '<span class="req"> *</span>' : ''} <small>${esc(COLUMN_TYPE_LABELS[f.meta.type])}</small></label>${control(f)}${
          f.hint || f.meta.note ? `<span class="hint">${esc(f.hint || f.meta.note || '')}</span>` : ''
        }<span class="err"></span></div>`,
    )
    .join('\n');
  return `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'">
<title>${esc(input.title)}</title>
<style>
body{margin:0;background:#f5f7fa;color:#1f2329;font:14px/1.6 "Microsoft YaHei UI","Microsoft YaHei",sans-serif}
main{max-width:640px;margin:0 auto;padding:24px 16px 60px}
h1{font-size:20px;margin:0 0 4px}.sub{color:#5c6370;font-size:12px;margin:0 0 16px}
.card{background:#fff;border:1px solid #e3e7ee;border-radius:12px;padding:16px;margin-bottom:16px}
.field{display:flex;flex-direction:column;gap:6px;margin-bottom:14px}label{font-size:13px}small{color:#8a919d}.req{color:#d92d20}
input[type=text],input[type=date],select,textarea{height:34px;border:1px solid #d3d9e3;border-radius:6px;padding:0 10px;font:inherit;background:#fff}
textarea{height:auto;padding:8px 10px}.opts{display:flex;flex-wrap:wrap;gap:6px 14px}.chk{display:flex;gap:6px;align-items:center}
.hint{font-size:12px;color:#5c6370}.err{font-size:12px;color:#b42318;min-height:0}.bad input,.bad select,.bad textarea{border-color:#d92d20}
button{height:34px;padding:0 16px;border-radius:6px;border:1px solid #d3d9e3;background:#fff;font:inherit;cursor:pointer}
button.primary{background:#3f5ef0;border-color:#3f5ef0;color:#fff}.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
table{width:100%;border-collapse:collapse;font-size:12px}td,th{border-bottom:1px solid #eef0f4;padding:4px 6px;text-align:left}
</style>
</head>
<body>
<main>
<h1>${esc(input.title)}</h1>
<p class="sub">${esc(input.packTitle)} · 填好后点“添加这条记录”，可以连续填多条；最后点“导出填写结果”，把下载的文件发回给收集人。</p>
<div class="card" id="form">
${rows}
<div class="bar"><button class="primary" id="add">添加这条记录</button><span id="msg" class="sub"></span></div>
</div>
<div class="card"><div class="bar"><b>已填写 <span id="n">0</span> 条</b><span style="flex:1"></span><button id="export" disabled>导出填写结果</button></div><table id="list"></table></div>
</main>
<script type="application/json" id="spec">${jsonForScript({ v: FORM_RESULT_VERSION, form: input.formId, pack: input.packId, title: input.title, fields: spec })}</script>
<script>
(function () {
  var spec = JSON.parse(document.getElementById('spec').textContent);
  var records = [];
  var $ = function (s, r) { return (r || document).querySelector(s); };
  function value(f) {
    var el = $('[data-id="' + f.id + '"]');
    if (f.type === 'multi') return Array.prototype.map.call(el.querySelectorAll('input:checked'), function (i) { return i.value; }).join(',');
    if (f.type === 'check') return el.checked ? '是' : '否';
    return (el.value || '').trim();
  }
  function check(f, t) {
    if (t === '') return f.required ? '必填' : '';
    var n;
    if (f.type === 'number' || f.type === 'money' || f.type === 'percent') {
      n = Number(t.replace(/[¥￥$,\\s%]/g, ''));
      if (!isFinite(n) || t.replace(/[¥￥$,\\s%]/g, '') === '') return '请输入数字';
      if (f.integer && n % 1 !== 0) return '只允许整数';
      if (f.min !== null && n < f.min) return '不能小于 ' + f.min;
      if (f.max !== null && n > f.max) return '不能大于 ' + f.max;
    } else if (f.type === 'email') { if (!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(t)) return '邮箱格式不对'; }
    else if (f.type === 'phone') { if (!/^[+\\d][\\d\\s\\-()]{4,}$/.test(t)) return '电话格式不对'; }
    else if (f.type === 'link') { if (!/^(https?:\\/\\/|mailto:|tel:)\\S+$/i.test(t)) return '链接要以 http:// 或 https:// 开头'; }
    else if (f.type === 'select') { if (f.options.length && f.options.indexOf(t) < 0) return '不在选项里'; }
    return '';
  }
  function reset() {
    spec.fields.forEach(function (f) {
      var el = $('[data-id="' + f.id + '"]');
      if (f.type === 'multi') Array.prototype.forEach.call(el.querySelectorAll('input'), function (i) { i.checked = false; });
      else if (f.type === 'check') el.checked = f.def === '是';
      else el.value = f.def && f.def.indexOf('{{') < 0 ? f.def : '';
    });
  }
  function render() {
    $('#n').textContent = records.length;
    $('#export').disabled = !records.length;
    var t = $('#list');
    t.innerHTML = '';
    var h = document.createElement('tr');
    spec.fields.forEach(function (f) { var th = document.createElement('th'); th.textContent = f.label; h.appendChild(th); });
    t.appendChild(h);
    records.forEach(function (r) {
      var tr = document.createElement('tr');
      spec.fields.forEach(function (f) { var td = document.createElement('td'); td.textContent = r[f.id] || ''; tr.appendChild(td); });
      t.appendChild(tr);
    });
  }
  $('#add').onclick = function () {
    var rec = {}, bad = 0;
    spec.fields.forEach(function (f) {
      var t = value(f), msg = check(f, t), box = $('[data-field="' + f.id + '"]');
      box.className = 'field' + (msg ? ' bad' : '');
      $('.err', box).textContent = msg;
      if (msg) bad++; else rec[f.id] = t;
    });
    if (bad) { $('#msg').textContent = '有 ' + bad + ' 项没填对'; return; }
    records.push(rec);
    $('#msg').textContent = '已添加';
    reset();
    render();
  };
  $('#export').onclick = function () {
    var out = JSON.stringify({ v: spec.v, form: spec.form, pack: spec.pack, records: records }, null, 2);
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([out], { type: 'application/json' }));
    a.download = spec.title + '-填写结果.json';
    a.click();
  };
  reset();
  render();
})();
</script>
</body>
</html>
`;
}
