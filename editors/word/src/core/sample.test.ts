// 生成排版对照样例：在 Word 与本编辑器中分别打开，对比换行、每页行数与分页。
// 运行：npm run gen:sample
import { writeFileSync } from 'fs';
import { it } from 'vitest';
import { writeDocx } from '../ooxml/docx';
import { FIELD_CHAR, PT_PER_CM, newDocument, newParagraph, newTable, textHeight, textWidth } from './model';

it.runIf(process.env.GEN_SAMPLE)('生成 editors/word/samples/排版对照样例.docx', () => {
  const d = newDocument();
  const s = d.setup;
  // GB/T 9704-2012：天头 37mm、订口 28mm，版心 156×225mm，每面 22 行、每行 28 字，三号仿宋
  Object.assign(s, { marginTop: 3.7 * PT_PER_CM, marginBottom: 3.5 * PT_PER_CM, marginLeft: 2.8 * PT_PER_CM, marginRight: 2.6 * PT_PER_CM, footerDistance: 2.8 * PT_PER_CM });
  d.defaults = { ...d.defaults, fontEA: '仿宋_GB2312', size: 16 };
  s.grid = { type: 'linesAndChars', linePitch: +(textHeight(s) / 22).toFixed(2), charSpace: Math.round((textWidth(s) / 28 - 16) * 4096) };
  const body = { ...d.defaults };
  const title = newParagraph('关于开展2026年第三季度办公用品集中采购工作的通知', { align: 'center', firstLineChars: null }, { ...body, fontEA: '方正小标宋简体', size: 22 });
  const p = (t: string) => newParagraph(t, { firstLineChars: 2 }, body);
  d.body = [
    title,
    newParagraph('', {}, body),
    newParagraph('各单位、各相关部门：', { firstLineChars: null }, body),
    p('为进一步规范办公用品采购管理，降低采购成本，提高保障效率，现就2026年第三季度集中采购工作通知如下。'),
    newParagraph('一、采购范围', { firstLineChars: 2 }, { ...body, fontEA: '黑体' }),
    p('本次集中采购范围包括电脑配件（Laptop charger、USB adapter、Network cable 等）、办公耗材、纸张及安全设备。各部应于9月30日前报送需求计划，逾期未报的，视为无采购需求。'),
    p('行末标点测试：这一行用来检验句号是否悬挂在版心右侧，也就是说标点溢出边界。'),
    p('行首禁则测试：这一段文字的长度经过调整，使右引号和右括号正好落在行首位置」）应被推到上一行或带走前一个字。'),
    newParagraph('二、工作要求', { firstLineChars: 2 }, { ...body, fontEA: '黑体' }),
    p('（一）各部要高度重视，指定专人负责，确保需求计划真实、准确、完整。'),
    p('（二）采购部门应按照“公开、公平、公正”的原则组织询价比价，形成比价表报分管领导审批。'),
    p('附件：主要物资需求汇总表'),
  ];
  const t = newTable(4, 5, textWidth(s), { ...body, size: 12 });
  ['序号', '物资名称', '规格型号', '数量'].forEach((h, i) => (t.rows[0].cells[i].paragraphs[0].runs = [{ text: h, props: { ...body, size: 12, fontEA: '黑体' } }]));
  t.rows[0].header = true;
  t.rows.slice(1).forEach((r, i) => {
    const v = [String(i + 1), ['气缸套', '活塞环', '排气阀座', '燃油喷嘴'][i], '[图号]', '4'];
    r.cells.forEach((c, j) => (c.paragraphs[0].runs = [{ text: v[j], props: { ...body, size: 12 } }]));
    r.height = 28;
    r.heightRule = 'atLeast';
  });
  d.body.push(t, newParagraph('', {}, body));
  for (let i = 0; i < 30; i++) d.body.push(p(`第${i + 1}段：本段用于填充页面以检查分页位置是否与 Word 一致。办公用品集中采购工作涉及需求、询价、比价、合同、验收等环节。`));
  const fp = { ...body, fontEA: '宋体', fontLatin: 'Times New Roman', size: 14 };
  d.footer = [newParagraph('', { align: 'center', firstLineChars: null }, fp)];
  d.footer[0].runs = [{ text: '— ', props: fp }, { text: FIELD_CHAR, props: fp, field: 'PAGE' }, { text: ' —', props: fp }];
  writeFileSync('editors/word/samples/排版对照样例.docx', writeDocx(d));
});
