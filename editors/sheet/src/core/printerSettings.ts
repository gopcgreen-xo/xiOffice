import { strFromU8, strToU8 } from 'fflate';
import { readRels } from './parts';
import { xmlAttr } from '@wz/editor-shared/xlsx';

/** Recognize the legacy binary → UTF-8 → binary corruption, not arbitrary
 * vendor printer formats. Healthy and unrecognized binary parts stay intact. */
function utf8DamagedDevMode(bytes: Uint8Array): boolean {
  if (bytes.length < 72) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // DEVMODEW: 32 UTF-16 device-name characters, then version and size fields.
  if (view.getUint16(64, true) !== 0x0401) return false;
  const size = view.getUint16(68, true), extra = view.getUint16(70, true);
  if (size >= 72 && size + extra === bytes.length) return false;
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes).includes('\uFFFD');
  } catch {
    return false;
  }
}

/** Keep paper, margins and scaling. Only detach an unusable printer-driver
 * reference; its opaque bytes and all unrelated package parts remain intact. */
export function detachDamagedPrinterSettings(files: Record<string, Uint8Array>): number {
  let repaired = 0;
  for (const path of Object.keys(files).filter(path => /^xl\/worksheets\/[^/]+\.xml$/.test(path))) {
    const badIds = new Set(readRels(files, path)
      .filter(rel => !rel.external && rel.type.endsWith('/printerSettings') && files[rel.path] && utf8DamagedDevMode(files[rel.path]))
      .map(rel => rel.id));
    if (!badIds.size) continue;
    const xml = strFromU8(files[path]);
    const next = xml.replace(/<pageSetup\b[^>]*>/g, tag => {
      const id = xmlAttr(tag, 'r:id');
      if (!id || !badIds.has(id)) return tag;
      repaired++;
      return tag.replace(/\s+r:id="[^"]*"/, '');
    });
    if (next !== xml) files[path] = strToU8(next);
  }
  return repaired;
}
