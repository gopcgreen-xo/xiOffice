/**
 * 数据透视表的 Excel 真部件：pivotCacheDefinition / pivotCacheRecords / pivotTable。
 *
 * 把 `SheetData.pivot`（编辑器里创建 / 刷新的透视表）写成 Excel 能刷新、能拖拽字段的真透视表。结构照 Excel 自己保存的
 * “表格形式、无分类汇总”的透视表来写（`compact="0" outline="0"`，行字段 `defaultSubtotal="0"`），单元格内容就是 pivot.ts 排好的
 * 同样布局，所以打开时看到的和 Excel 刷新后一样；缓存里带 `refreshOnLoad="1"`，Excel 打开时会按源数据重新汇总一次。
 *
 * 支持：任意个行字段、最多一个列字段、1 个以上值字段（求和 / 计数 / 数值计数 / 平均 / 最大 / 最小）、总计行 / 列开关。
 * 不写：筛选、切片器、日期分组、计算字段、值显示方式。文件里原有的透视表不动。
 */
import { strToU8 } from 'fflate';
import { colName } from '@wz/editor-shared/xlsx';
import { addOverride, addRel, CT, esc, nextPart, REL, type Files, type Pkg } from './parts';
import { analyzePivot, cmpLabel, effectiveValues, keyId, labelOf, readSource, valueName, type Agg, type Label, type PivotConfig, type PivotDef, type PivotParts } from './pivot';
import type { Scalar } from './formula/eval';
import type { SheetData, Workbook } from './workbook';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const SUBTOTAL: Partial<Record<Agg, string>> = { avg: 'average', max: 'max', min: 'min', counta: 'count', count: 'countNums' };

/** 能写成真透视表吗（列字段最多一个，字段号都在数据区域里） */
export function canWritePivot(def: PivotDef, width: number): boolean {
  const c = def.config;
  const values = effectiveValues(c, width);
  const axis = new Set([...c.rows, ...c.cols]);
  // 值字段不能同时是行 / 列字段（Excel 不允许）
  return c.cols.length <= 1 && [...axis, ...values.map((v) => v.field)].every((f) => f >= 0 && f < width) && values.every((v) => !axis.has(v.field));
}

interface FieldInfo {
  items: Label[];
  /** 值是否写成共享项（字符串、或用在行 / 列上的字段） */
  shared: boolean;
  hasBlank: boolean;
  hasString: boolean;
  hasNumber: boolean;
  allInt: boolean;
  min: number;
  max: number;
}

function fieldInfo(records: Scalar[][], f: number, axis: boolean): FieldInfo {
  const info: FieldInfo = { items: [], shared: false, hasBlank: false, hasString: false, hasNumber: false, allInt: true, min: Infinity, max: -Infinity };
  const seen = new Map<string, Label>();
  for (const row of records) {
    const v = row[f] ?? null;
    if (v === null || v === '') {
      info.hasBlank = true;
      continue;
    }
    if (typeof v === 'number') {
      info.hasNumber = true;
      info.allInt &&= Number.isInteger(v);
      info.min = Math.min(info.min, v);
      info.max = Math.max(info.max, v);
    } else info.hasString = true;
    seen.set(`${typeof v}:${v}`, v as Label);
  }
  info.shared = info.hasString || axis;
  info.items = [...seen.values()].sort((a, b) => cmpLabel(a, b));
  return info;
}

function sharedItemsXml(i: FieldInfo): string {
  const num = i.hasNumber ? ` containsNumber="1"${i.allInt ? ' containsInteger="1"' : ''} minValue="${i.min}" maxValue="${i.max}"` : '';
  const blank = i.hasBlank ? ' containsBlank="1"' : '';
  if (!i.shared) {
    // 纯数字且不在行 / 列上：只记范围，不列项
    const semi = i.hasBlank ? '' : ' containsSemiMixedTypes="0"';
    return `<sharedItems${semi} containsString="0"${blank}${num}/>`;
  }
  const items = i.items.map((v) => (typeof v === 'number' ? `<n v="${v}"/>` : `<s v="${esc(String(v))}"/>`)).join('') + (i.hasBlank ? '<m/>' : '');
  const count = i.items.length + (i.hasBlank ? 1 : 0);
  if (i.hasString && i.hasNumber) return `<sharedItems containsMixedTypes="1"${blank}${num} count="${count}">${items}</sharedItems>`;
  if (i.hasNumber) return `<sharedItems${i.hasBlank ? '' : ' containsSemiMixedTypes="0"'} containsString="0"${blank}${num} count="${count}">${items}</sharedItems>`;
  return `<sharedItems${blank} count="${count}">${items}</sharedItems>`;
}

/** 三个部件的 XML 文字 */
export function pivotXml(src: SheetData, def: PivotDef, extent: string, name: string, cacheId: number, now: Date = new Date()): { cacheDef: string; cacheRec: string; table: string } {
  const g = def.source.range;
  const data = readSource(src, g);
  const cfg: PivotConfig = def.config;
  const an = analyzePivot(data, cfg);
  const names = an.fieldNames;
  const nf = names.length;
  const R = cfg.rows.length;
  const hasCols = cfg.cols.length > 0;
  const values = effectiveValues(cfg, nf);
  const K = values.length;
  const axisFields = new Set([...cfg.rows, ...cfg.cols]);
  const infos = names.map((_, f) => fieldInfo(an.records, f, axisFields.has(f)));
  const pos = infos.map((i) => new Map(i.items.map((v, k) => [keyId([labelOf(v)]), k])));
  const blankPos = (f: number) => infos[f].items.length;
  const idxOf = (f: number, v: Scalar): number => (v === null || v === '' ? blankPos(f) : (pos[f].get(keyId([v as Label])) ?? -1));

  // ---- 缓存定义 ----
  const cacheDef =
    `${HEAD}<pivotCacheDefinition xmlns="${NS}" xmlns:r="${NS_R}" r:id="rId1" refreshOnLoad="1" refreshedBy="xiOffice" refreshedDate="${25569 + now.getTime() / 86400000}" createdVersion="3" refreshedVersion="6" minRefreshableVersion="3" recordCount="${an.records.length}">` +
    `<cacheSource type="worksheet"><worksheetSource ref="${colName(g.c0)}${g.r0 + 1}:${colName(g.c1)}${g.r1 + 1}" sheet="${esc(src.name)}"/></cacheSource>` +
    `<cacheFields count="${nf}">${names.map((n, f) => `<cacheField name="${esc(n)}" numFmtId="0">${sharedItemsXml(infos[f])}</cacheField>`).join('')}</cacheFields></pivotCacheDefinition>`;

  // ---- 缓存记录 ----
  const rec = (row: Scalar[]) =>
    `<r>${names
      .map((_, f) => {
        const v = row[f] ?? null;
        if (infos[f].shared) return `<x v="${idxOf(f, v)}"/>`;
        return v === null || v === '' ? '<m/>' : `<n v="${v}"/>`;
      })
      .join('')}</r>`;
  const cacheRec = `${HEAD}<pivotCacheRecords xmlns="${NS}" xmlns:r="${NS_R}" count="${an.records.length}">${an.records.map(rec).join('')}</pivotCacheRecords>`;

  // ---- 透视表 ----
  const role = (f: number) => (cfg.rows.includes(f) ? 'axisRow' : cfg.cols.includes(f) ? 'axisCol' : '');
  const dataFieldSet = new Set(values.map((v) => v.field));
  const pf = names
    .map((_, f) => {
      const r = role(f);
      const items = infos[f].items.length + (infos[f].hasBlank ? 1 : 0);
      const itemList = Array.from({ length: items }, (_, k) => `<item x="${k}"/>`).join('');
      const attrs = `${r ? ` axis="${r}"` : ''}${dataFieldSet.has(f) ? ' dataField="1"' : ''} compact="0" outline="0" showAll="0"${r === 'axisRow' ? ' defaultSubtotal="0"' : ''}`;
      if (!r) return `<pivotField${attrs}/>`;
      return `<pivotField${attrs}><items count="${items + (r === 'axisCol' ? 1 : 0)}">${itemList}${r === 'axisCol' ? '<item t="default"/>' : ''}</items></pivotField>`;
    })
    .join('');
  const x = (v: number) => (v ? `<x v="${v}"/>` : '<x/>');

  // 行项目：每行的各级位置；和上一行相同的前几级用 r 省略
  let rowItems = '';
  let nRow = 0;
  if (R === 0) {
    rowItems = '<i/>';
    nRow = 1;
  } else {
    let prev: number[] = [];
    for (const rk of an.rowKeys) {
      const p = rk.map((l, j) => idxOf(cfg.rows[j], l === '(空白)' ? null : l));
      let same = 0;
      while (same < p.length - 1 && prev[same] === p[same]) same++;
      rowItems += `<i${same ? ` r="${same}"` : ''}>${p.slice(same).map(x).join('')}</i>`;
      nRow++;
      prev = p;
    }
    if (cfg.totals.rows) {
      rowItems += '<i t="grand"><x/></i>';
      nRow++;
    }
  }
  // 列项目
  let colFields = '';
  let colItems = '';
  let nCol = 0;
  const showTotCol = hasCols && cfg.totals.cols;
  if (!hasCols) {
    if (K > 1) {
      colFields = '<colFields count="1"><field x="-2"/></colFields>';
      for (let k = 0; k < K; k++) colItems += `<i${k ? ` i="${k}"` : ''}>${x(k)}</i>`;
      nCol = K;
    } else {
      colItems = '<i/>';
      nCol = 1;
    }
  } else {
    colFields = `<colFields count="${K > 1 ? 2 : 1}"><field x="${cfg.cols[0]}"/>${K > 1 ? '<field x="-2"/>' : ''}</colFields>`;
    for (const ck of an.colKeys) {
      const j = idxOf(cfg.cols[0], ck[0] === '(空白)' ? null : ck[0]);
      if (K === 1) {
        colItems += `<i>${x(j)}</i>`;
        nCol++;
      } else {
        for (let k = 0; k < K; k++) {
          colItems += k === 0 ? `<i>${x(j)}<x/></i>` : `<i r="1" i="${k}">${x(k)}</i>`;
          nCol++;
        }
      }
    }
    if (showTotCol) {
      for (let k = 0; k < K; k++) {
        colItems += k === 0 ? '<i t="grand"><x/></i>' : `<i t="grand" i="${k}"><x/></i>`;
        nCol++;
      }
    }
  }
  const dataFields = values
    .map((v) => `<dataField name="${esc(valueName(v, names))}" fld="${v.field}"${SUBTOTAL[v.agg] ? ` subtotal="${SUBTOTAL[v.agg]}"` : ''} baseField="0" baseItem="0"/>`)
    .join('');
  const headerRows = (hasCols ? 1 : 0) + (K > 1 ? 1 : 0) + 1;
  // rowGrandTotals = 底部的总计行，colGrandTotals = 右边的总计列（Excel 的命名，对拍确认过）
  const flags = `${R > 0 && !cfg.totals.rows ? ' rowGrandTotals="0"' : ''}${hasCols && !cfg.totals.cols ? ' colGrandTotals="0"' : ''}`;
  const table =
    `${HEAD}<pivotTableDefinition xmlns="${NS}" name="${esc(name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="值" updatedVersion="6" minRefreshableVersion="3" showCalcMbrs="0" useAutoFormatting="1" itemPrintTitles="1" createdVersion="3" indent="0" compact="0" compactData="0" multipleFieldFilters="0"${flags}>` +
    `<location ref="${extent}" firstHeaderRow="1" firstDataRow="${headerRows}" firstDataCol="${R === 0 && hasCols && K === 1 ? 1 : R}"/>` +
    `<pivotFields count="${nf}">${pf}</pivotFields>` +
    (R ? `<rowFields count="${R}">${cfg.rows.map((f) => `<field x="${f}"/>`).join('')}</rowFields>` : '') +
    `<rowItems count="${nRow}">${rowItems}</rowItems>${colFields}<colItems count="${nCol}">${colItems}</colItems>` +
    `<dataFields count="${K}">${dataFields}</dataFields>` +
    `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
  return { cacheDef, cacheRec, table };
}

const a1 = (g: { r0: number; c0: number; r1: number; c1: number }) => `${colName(g.c0)}${g.r0 + 1}${g.r0 === g.r1 && g.c0 === g.c1 ? '' : `:${colName(g.c1)}${g.r1 + 1}`}`;

/**
 * 把这张表的透视表写进包里。已有部件（本次会话保存过）原地覆盖，没有就新建并建好关系。
 * 返回部件位置（保存成功后记到模型上）；isNew = 要在 workbook.xml 里登记新缓存。
 */
export function syncPivot(pkg: Pkg, sheetPath: string, sheet: SheetData, book: Workbook, usedCacheIds: Set<number>): { parts: PivotParts; isNew: boolean } | null {
  const def = sheet.pivot;
  if (!def) return null;
  const src = book.sheets.find((s) => s.name === def.source.sheet);
  if (!src || !canWritePivot(def, def.source.range.c1 - def.source.range.c0 + 1)) return null;
  const files: Files = pkg.files;
  let parts = def.parts;
  const isNew = !(parts && files[parts.table] && files[parts.cacheDef]);
  if (isNew) {
    let id = 1;
    while (usedCacheIds.has(id)) id++;
    parts = {
      table: nextPart(files, 'xl/pivotTables', 'pivotTable', '.xml'),
      cacheDef: nextPart(files, 'xl/pivotCache', 'pivotCacheDefinition', '.xml'),
      cacheRec: nextPart(files, 'xl/pivotCache', 'pivotCacheRecords', '.xml'),
      cacheId: id,
    };
    usedCacheIds.add(id);
  }
  const p = parts!;
  const name = `数据透视表${p.cacheId}`;
  const xml = pivotXml(src, def, a1(def.extent), name, p.cacheId);
  files[p.cacheDef] = strToU8(xml.cacheDef);
  files[p.cacheRec] = strToU8(xml.cacheRec);
  files[p.table] = strToU8(xml.table);
  if (isNew) {
    addOverride(pkg, p.table, CT.pivotTable);
    addOverride(pkg, p.cacheDef, CT.pivotCacheDef);
    addOverride(pkg, p.cacheRec, CT.pivotCacheRec);
    addRel(files, p.cacheDef, REL.pivotCacheRec, p.cacheRec);
    addRel(files, p.table, REL.pivotCacheDef, p.cacheDef);
    addRel(files, sheetPath, REL.pivotTable, p.table);
  }
  return { parts: p, isNew };
}
