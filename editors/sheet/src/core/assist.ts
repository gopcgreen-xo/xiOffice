/**
 * 写公式的帮助：自动求和（Alt+=）选哪片区域；输入函数名时的候选列表；光标在函数括号里时显示参数说明。
 */
import { SUPPORTED } from './formula/eval';
import type { Range, SheetData } from './workbook';

/** 函数说明：参数（[ ] 为可选）与一句话用途 */
export const FUNCTION_INFO: Record<string, { args: string[]; desc: string }> = {
  SUM: { args: ['数值1', '[数值2]', '…'], desc: '求和' },
  PRODUCT: { args: ['数值1', '[数值2]', '…'], desc: '乘积' },
  AVERAGE: { args: ['数值1', '[数值2]', '…'], desc: '平均值' },
  MAX: { args: ['数值1', '[数值2]', '…'], desc: '最大值' },
  MIN: { args: ['数值1', '[数值2]', '…'], desc: '最小值' },
  COUNT: { args: ['值1', '[值2]', '…'], desc: '数字的个数' },
  COUNTA: { args: ['值1', '[值2]', '…'], desc: '非空单元格的个数' },
  COUNTBLANK: { args: ['区域'], desc: '空单元格的个数' },
  COUNTIF: { args: ['区域', '条件'], desc: '满足条件的单元格个数' },
  COUNTIFS: { args: ['区域1', '条件1', '[区域2, 条件2]', '…'], desc: '同时满足多个条件的个数' },
  SUMIF: { args: ['条件区域', '条件', '[求和区域]'], desc: '按条件求和' },
  SUMIFS: { args: ['求和区域', '条件区域1', '条件1', '…'], desc: '按多个条件求和' },
  AVERAGEIF: { args: ['条件区域', '条件', '[平均值区域]'], desc: '按条件求平均值' },
  AVERAGEIFS: { args: ['平均值区域', '条件区域1', '条件1', '…'], desc: '按多个条件求平均值' },
  MAXIFS: { args: ['最大值区域', '条件区域1', '条件1', '…'], desc: '按条件求最大值' },
  MINIFS: { args: ['最小值区域', '条件区域1', '条件1', '…'], desc: '按条件求最小值' },
  SUMPRODUCT: { args: ['数组1', '[数组2]', '…'], desc: '对应元素相乘再求和' },
  ROUND: { args: ['数值', '小数位数'], desc: '四舍五入' },
  ROUNDUP: { args: ['数值', '小数位数'], desc: '向上舍入（远离 0）' },
  ROUNDDOWN: { args: ['数值', '小数位数'], desc: '向下舍入（靠近 0）' },
  INT: { args: ['数值'], desc: '向下取整' },
  ABS: { args: ['数值'], desc: '绝对值' },
  MOD: { args: ['被除数', '除数'], desc: '余数' },
  IF: { args: ['条件', '[成立时的值]', '[不成立时的值]'], desc: '按条件取值' },
  IFS: { args: ['条件1', '值1', '[条件2, 值2]', '…'], desc: '取第一个成立的条件对应的值' },
  IFERROR: { args: ['值', '出错时的值'], desc: '出错时换成指定的值' },
  IFNA: { args: ['值', '#N/A 时的值'], desc: '#N/A 时换成指定的值' },
  AND: { args: ['条件1', '[条件2]', '…'], desc: '全部成立' },
  OR: { args: ['条件1', '[条件2]', '…'], desc: '任一成立' },
  NOT: { args: ['条件'], desc: '取反' },
  ISBLANK: { args: ['值'], desc: '是否为空' },
  ISNUMBER: { args: ['值'], desc: '是否为数字' },
  ISTEXT: { args: ['值'], desc: '是否为文字' },
  ISERROR: { args: ['值'], desc: '是否为错误值' },
  ISNA: { args: ['值'], desc: '是否为 #N/A' },
  VLOOKUP: { args: ['查找值', '区域', '列序号', '[近似匹配]'], desc: '在区域首列查找，返回同一行指定列的值' },
  HLOOKUP: { args: ['查找值', '区域', '行序号', '[近似匹配]'], desc: '在区域首行查找，返回同一列指定行的值' },
  INDEX: { args: ['区域', '行号', '[列号]'], desc: '区域里第几行第几列的值' },
  MATCH: { args: ['查找值', '查找区域', '[匹配方式]'], desc: '查找值在区域里的位置' },
  XLOOKUP: { args: ['查找值', '查找区域', '返回区域', '[找不到时]', '[匹配方式]'], desc: '按查找区域找到后返回对应的值' },
  ROW: { args: ['[引用]'], desc: '行号' },
  COLUMN: { args: ['[引用]'], desc: '列号' },
  CHOOSE: { args: ['序号', '值1', '[值2]', '…'], desc: '按序号从列表里取值' },
  LEN: { args: ['文字'], desc: '字数' },
  LEFT: { args: ['文字', '[字数]'], desc: '左边几个字' },
  RIGHT: { args: ['文字', '[字数]'], desc: '右边几个字' },
  MID: { args: ['文字', '开始位置', '字数'], desc: '从中间取几个字' },
  TRIM: { args: ['文字'], desc: '去掉多余的空格' },
  UPPER: { args: ['文字'], desc: '转大写' },
  LOWER: { args: ['文字'], desc: '转小写' },
  CONCATENATE: { args: ['文字1', '[文字2]', '…'], desc: '连接文字' },
  CONCAT: { args: ['文字1', '[文字2]', '…'], desc: '连接文字（可以是区域）' },
  TEXTJOIN: { args: ['分隔符', '忽略空值', '文字1', '…'], desc: '用分隔符连接文字' },
  SUBSTITUTE: { args: ['文字', '旧文字', '新文字', '[第几处]'], desc: '替换文字' },
  FIND: { args: ['要找的文字', '在哪段文字里', '[开始位置]'], desc: '查找位置（区分大小写）' },
  SEARCH: { args: ['要找的文字', '在哪段文字里', '[开始位置]'], desc: '查找位置（不区分大小写，可用通配符）' },
  TEXT: { args: ['数值', '格式代码'], desc: '按格式显示成文字' },
  VALUE: { args: ['文字'], desc: '文字转数字' },
  TODAY: { args: [], desc: '今天的日期' },
  NOW: { args: [], desc: '现在的日期和时间' },
  DATE: { args: ['年', '月', '日'], desc: '组成日期' },
  YEAR: { args: ['日期'], desc: '年份' },
  MONTH: { args: ['日期'], desc: '月份' },
  DAY: { args: ['日期'], desc: '日' },
  DATEDIF: { args: ['开始日期', '结束日期', '单位（Y/M/D/YM/MD/YD）'], desc: '两个日期相差的年 / 月 / 天' },
  SUBTOTAL: { args: ['函数号', '区域1', '[区域2]', '…'], desc: '分类汇总（忽略其中的 SUBTOTAL；101–111 还忽略隐藏行）' },
  RANK: { args: ['数值', '区域', '[升序]'], desc: '排名' },
  EDATE: { args: ['开始日期', '月数'], desc: '几个月之前 / 之后的日期' },
  EOMONTH: { args: ['开始日期', '月数'], desc: '几个月之前 / 之后那个月的最后一天' },
  WEEKDAY: { args: ['日期', '[类型]'], desc: '星期几' },
  NETWORKDAYS: { args: ['开始日期', '结束日期', '[假日]'], desc: '两个日期之间的工作日天数' },
  WORKDAY: { args: ['开始日期', '天数', '[假日]'], desc: '若干个工作日之后的日期' },
  OFFSET: { args: ['引用', '行偏移', '列偏移', '[高度]', '[宽度]'], desc: '从引用偏移得到新的区域' },
  INDIRECT: { args: ['引用文字'], desc: '把文字当作引用' },
  LARGE: { args: ['区域', '第几个'], desc: '第 k 大的值' },
  SMALL: { args: ['区域', '第几个'], desc: '第 k 小的值' },
  MEDIAN: { args: ['数值1', '[数值2]', '…'], desc: '中位数' },
  STDEV: { args: ['数值1', '[数值2]', '…'], desc: '样本标准差' },
  STDEVP: { args: ['数值1', '[数值2]', '…'], desc: '总体标准差' },
  VAR: { args: ['数值1', '[数值2]', '…'], desc: '样本方差' },
  VARP: { args: ['数值1', '[数值2]', '…'], desc: '总体方差' },
  // ---------- 第二批 ----------
  SQRT: { args: ['数值'], desc: '平方根' },
  POWER: { args: ['底数', '指数'], desc: '乘方' },
  SIGN: { args: ['数值'], desc: '正数 1、负数 -1、零 0' },
  LOG: { args: ['数值', '[底数]'], desc: '对数（底数默认 10）' },
  LN: { args: ['数值'], desc: '自然对数' },
  EXP: { args: ['数值'], desc: 'e 的几次方' },
  PI: { args: [], desc: '圆周率' },
  RAND: { args: [], desc: '0 到 1 的随机数' },
  RANDBETWEEN: { args: ['最小值', '最大值'], desc: '两个整数之间的随机整数' },
  TRUNC: { args: ['数值', '[小数位数]'], desc: '直接截掉小数（不四舍五入）' },
  CEILING: { args: ['数值', '倍数'], desc: '向上取整到倍数' },
  FLOOR: { args: ['数值', '倍数'], desc: '向下取整到倍数' },
  MROUND: { args: ['数值', '倍数'], desc: '四舍五入到最接近的倍数' },
  GCD: { args: ['数值1', '[数值2]', '…'], desc: '最大公约数' },
  LCM: { args: ['数值1', '[数值2]', '…'], desc: '最小公倍数' },
  QUOTIENT: { args: ['被除数', '除数'], desc: '除法的整数部分' },
  MODE: { args: ['数值1', '[数值2]', '…'], desc: '出现最多的数' },
  PERCENTILE: { args: ['区域', 'k（0~1）'], desc: '百分位数' },
  QUARTILE: { args: ['区域', '第几个四分位（0~4）'], desc: '四分位数' },
  CORREL: { args: ['区域1', '区域2'], desc: '相关系数' },
  SLOPE: { args: ['已知 y', '已知 x'], desc: '线性回归的斜率' },
  INTERCEPT: { args: ['已知 y', '已知 x'], desc: '线性回归的截距' },
  FORECAST: { args: ['x', '已知 y', '已知 x'], desc: '按线性回归预测' },
  REPT: { args: ['文字', '次数'], desc: '重复文字' },
  CHAR: { args: ['编码'], desc: '编码转字符' },
  CODE: { args: ['文字'], desc: '第一个字符的编码' },
  PROPER: { args: ['文字'], desc: '每个单词首字母大写' },
  CLEAN: { args: ['文字'], desc: '去掉不能打印的字符' },
  EXACT: { args: ['文字1', '文字2'], desc: '两段文字是否完全相同（区分大小写）' },
  REPLACE: { args: ['文字', '开始位置', '字数', '新文字'], desc: '按位置替换文字' },
  TEXTBEFORE: { args: ['文字', '分隔符', '[第几个]'], desc: '分隔符之前的文字' },
  TEXTAFTER: { args: ['文字', '分隔符', '[第几个]'], desc: '分隔符之后的文字' },
  TEXTSPLIT: { args: ['文字', '列分隔符', '[行分隔符]', '[忽略空值]'], desc: '按分隔符拆成多格（溢出）' },
  ISERR: { args: ['值'], desc: '是否是 #N/A 以外的错误' },
  ISEVEN: { args: ['数值'], desc: '是否是偶数' },
  ISODD: { args: ['数值'], desc: '是否是奇数' },
  ISLOGICAL: { args: ['值'], desc: '是否是 TRUE / FALSE' },
  XOR: { args: ['条件1', '[条件2]', '…'], desc: '为真的个数是奇数时返回 TRUE' },
  SWITCH: { args: ['表达式', '值1', '结果1', '…', '[默认值]'], desc: '按值匹配返回结果' },
  ROWS: { args: ['区域'], desc: '行数' },
  COLUMNS: { args: ['区域'], desc: '列数' },
  ADDRESS: { args: ['行号', '列号', '[绝对类型]', '[A1 样式]', '[表名]'], desc: '生成单元格地址文字' },
  LOOKUP: { args: ['查找值', '查找区域', '[结果区域]'], desc: '在已排序的区域里近似查找' },
  XMATCH: { args: ['查找值', '查找区域', '[匹配方式]', '[搜索方向]'], desc: '查找值在区域里的位置（支持反向、近似、通配符）' },
  TRANSPOSE: { args: ['区域'], desc: '行列互换（溢出）' },
  SEQUENCE: { args: ['行数', '[列数]', '[起始值]', '[步长]'], desc: '生成数字序列（溢出）' },
  FILTER: { args: ['区域', '条件', '[没有结果时]'], desc: '按条件筛选出行（溢出）' },
  UNIQUE: { args: ['区域', '[按列]', '[只要出现一次的]'], desc: '去掉重复项（溢出）' },
  SORT: { args: ['区域', '[按第几列]', '[1 升序 / -1 降序]', '[按列]'], desc: '排序（溢出）' },
  SORTBY: { args: ['区域', '排序依据1', '[1 / -1]', '…'], desc: '按另外的区域排序（溢出）' },
  VSTACK: { args: ['区域1', '[区域2]', '…'], desc: '上下拼接（溢出）' },
  HSTACK: { args: ['区域1', '[区域2]', '…'], desc: '左右拼接（溢出）' },
  DAYS: { args: ['结束日期', '开始日期'], desc: '两个日期相差的天数' },
  WEEKNUM: { args: ['日期', '[每周从哪天开始]'], desc: '一年中的第几周' },
  ISOWEEKNUM: { args: ['日期'], desc: 'ISO 周数' },
  YEARFRAC: { args: ['开始日期', '结束日期', '[基准]'], desc: '两个日期之间相当于几年' },
  DATEVALUE: { args: ['日期文字'], desc: '日期文字转序列号' },
  HOUR: { args: ['时间'], desc: '小时' },
  MINUTE: { args: ['时间'], desc: '分钟' },
  SECOND: { args: ['时间'], desc: '秒' },
  TIME: { args: ['时', '分', '秒'], desc: '组成时间' },
  PMT: { args: ['利率', '期数', '现值', '[终值]', '[期初 1 / 期末 0]'], desc: '等额分期付款的每期金额' },
  IPMT: { args: ['利率', '第几期', '期数', '现值', '[终值]', '[类型]'], desc: '某一期付款里的利息' },
  PPMT: { args: ['利率', '第几期', '期数', '现值', '[终值]', '[类型]'], desc: '某一期付款里的本金' },
  FV: { args: ['利率', '期数', '每期付款', '[现值]', '[类型]'], desc: '终值' },
  PV: { args: ['利率', '期数', '每期付款', '[终值]', '[类型]'], desc: '现值' },
  NPER: { args: ['利率', '每期付款', '现值', '[终值]', '[类型]'], desc: '还款期数' },
  RATE: { args: ['期数', '每期付款', '现值', '[终值]', '[类型]', '[猜测值]'], desc: '每期利率' },
  NPV: { args: ['折现率', '现金流1', '[现金流2]', '…'], desc: '净现值' },
  IRR: { args: ['现金流区域', '[猜测值]'], desc: '内部收益率' },
  SLN: { args: ['原值', '残值', '使用年限'], desc: '直线折旧' },
};

/** 已支持计算、有说明的函数，按名称排序 */
export function functionNames(): string[] {
  return Object.keys(FUNCTION_INFO)
    .filter((n) => SUPPORTED.has(n) && FUNCTION_INFO[n].desc)
    .sort();
}

/** 光标前正在输入的函数名（公式里，不在字符串里，前面是运算符或括号） */
export function typingName(text: string, caret: number): { start: number; prefix: string } | null {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  if ((before.match(/"/g)?.length ?? 0) % 2) return null;
  const m = before.match(/(^|[=(,+\-*/^&<>:; ])([A-Za-z][A-Za-z0-9.]*)$/);
  if (!m) return null;
  return { start: caret - m[2].length, prefix: m[2].toUpperCase() };
}

/** 函数名候选：开头相同的在前 */
export function suggest(prefix: string): string[] {
  if (!prefix) return [];
  const names = functionNames();
  return names.filter((n) => n.startsWith(prefix)).concat(names.filter((n) => !n.startsWith(prefix) && n.includes(prefix))).slice(0, 8);
}

/** 光标所在的函数调用：最里层还没闭合的函数名和第几个参数（从 0 开始） */
export function callAt(text: string, caret: number): { name: string; arg: number } | null {
  if (!text.startsWith('=')) return null;
  const stack: { name: string; arg: number }[] = [];
  let quoted = false;
  for (let i = 1; i < caret; i++) {
    const ch = text[i];
    if (ch === '"') quoted = !quoted;
    if (quoted) continue;
    if (ch === '(') {
      const m = text.slice(0, i).match(/([A-Za-z][A-Za-z0-9.]*)\s*$/);
      stack.push({ name: m ? m[1].toUpperCase() : '', arg: 0 });
    } else if (ch === ')') stack.pop();
    else if (ch === ',' && stack.length) stack[stack.length - 1].arg++;
  }
  const top = stack[stack.length - 1];
  return top && FUNCTION_INFO[top.name] ? top : null;
}

/**
 * 自动求和（同 Excel 的 Alt+=）：活动单元格上方是数字就向上取连续的数字区（中间的空格子之前停），
 * 否则看左边；都没有返回 null（只插入 =SUM()）。
 */
export function autoSumRange(sheet: SheetData, r: number, c: number): Range | null {
  const num = (rr: number, cc: number) => typeof sheet.cells[rr]?.[cc]?.v === 'number';
  if (r > 0 && num(r - 1, c)) {
    let r0 = r - 1;
    while (r0 > 0 && num(r0 - 1, c)) r0--;
    return { r0, c0: c, r1: r - 1, c1: c };
  }
  if (c > 0 && num(r, c - 1)) {
    let c0 = c - 1;
    while (c0 > 0 && num(r, c0 - 1)) c0--;
    return { r0: r, c0, r1: r, c1: c - 1 };
  }
  return null;
}
