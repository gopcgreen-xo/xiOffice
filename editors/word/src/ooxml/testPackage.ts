/** 测试用：手写最小 docx（document.xml + styles.xml + 关系），不依赖任何真实文档 */
import { strToU8, zipSync } from 'fflate';

export const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
export const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface PackageOptions {
  /** document.xml 末尾的 sectPr 内容（缺省 A4 纸 + 页边距） */
  sect?: string;
  /** styles.xml 里 docDefaults 之后的样式 */
  styles?: string;
  /** 额外的关系（document.xml.rels 里） */
  rels?: string;
  /** 额外的部件 */
  extra?: Record<string, string>;
  /** 额外的内容类型 Override */
  types?: string;
}

export function pkg(body: string, o: PackageOptions = {}): Uint8Array {
  const sect = o.sect ?? '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
  const styles =
    o.styles ??
    '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>';
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        `${o.types ?? ''}</Types>`,
    ),
    '_rels/.rels': strToU8(`${XML}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${OD}/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/_rels/document.xml.rels': strToU8(`${XML}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${OD}/styles" Target="styles.xml"/>${o.rels ?? ''}</Relationships>`),
    'word/styles.xml': strToU8(
      `${XML}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>${styles}</w:styles>`,
    ),
    'word/document.xml': strToU8(`${XML}<w:document ${W} ${R}><w:body>${body}<w:sectPr>${sect}</w:sectPr></w:body></w:document>`),
  };
  for (const [k, v] of Object.entries(o.extra ?? {})) files[k] = strToU8(v);
  return zipSync(files);
}
