// Мінімальний генератор .xlsx (один аркуш): клітинки зі значеннями/формулами, стилі, об'єднання, ширини колонок.
// Без зовнішніх бібліотек: ZIP без стиснення + SpreadsheetML.

const enc = new TextEncoder();
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b) { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function zip(files) {             // files: [[name, string]]
  const parts = [], central = []; let off = 0;
  const d = new Date(), dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(), tm = (d.getHours() << 11) | (d.getMinutes() << 5);
  for (const [name, text] of files) {
    const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
    h.setUint16(10, tm, true); h.setUint16(12, dt, true); h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
    h.setUint16(26, nb.length, true); h.setUint16(28, 0, true);
    parts.push(new Uint8Array(h.buffer), nb, data);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
    c.setUint16(12, tm, true); c.setUint16(14, dt, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
    c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
    central.push(new Uint8Array(c.buffer), nb);
    off += 30 + nb.length + data.length;
  }
  const csize = central.reduce((a, x) => a + x.length, 0);
  const e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, csize, true); e.setUint32(16, off, true);
  return new Blob([...parts, ...central, new Uint8Array(e.buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
const x = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
export const col = (n) => { let s = ""; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };   // 0 → A
export const ref = (c, r) => col(c) + (r + 1);

// стилі: шрифт Times New Roman; індекси використовуються в cell.s
// розміри — як у реєстрі ЦЗ, який друкується на A4 альбомно (масштаб «вмістити по ширині»)
// розміри шрифту під друк A4 альбомно майже без зменшення (масштаб «по ширині» ≈ 85–100 %)
const FONTS = [["11", true, false], ["13", true, false], ["8", false, true], ["10", true, false], ["11", true, false], ["12", true, false], ["9", true, false], ["9", false, false]];
//         [font, border, h-align, v-align, wrap]
export const ST = { def: 0, title: 1, cap: 2, head: 3, txt: 4, num: 5, ctr: 6, totN: 7, totT: 8, plain: 9, num2: 10, sign: 11, titleL: 12, num1: 13, totN2: 14,
  headV: 15, headS: 16, ctrS: 17, numS: 18, totS: 19, titleR: 20, note: 21, headV10: 22 };
//         [font, border, h-align, v-align, wrap, numFmtId]   164 = «0.0#» (залишок на початок), 2 = «0.00» (витрачено, залишок)
const XF = [[0, 0, "", "", 0], [1, 0, "center", "center", 0], [2, 0, "center", "top", 0], [3, 1, "center", "center", 1], [0, 1, "center", "center", 1],
  [0, 1, "center", "center", 0], [0, 1, "center", "center", 1], [4, 1, "center", "center", 0], [4, 1, "center", "center", 0], [0, 0, "left", "top", 1], [0, 1, "center", "center", 0, 2],
  [5, 0, "left", "center", 0], [1, 0, "left", "center", 0], [0, 1, "center", "center", 0, 164], [4, 1, "center", "center", 0, 2],
  // реєстр «Додаток 6» (звичайний підрозділ): шрифт 9 — заголовок вертикальний / звичайний, текст і числа з переносом, підсумок жирний; «Додаток № 6» праворуч
  [7, 1, "center", "center", 1, 0, 90], [7, 1, "center", "center", 1], [7, 1, "center", "center", 1], [7, 1, "center", "center", 0], [6, 1, "center", "center", 1], [5, 0, "right", "center", 0], [7, 0, "left", "center", 0], [3, 1, "center", "center", 1, 0, 90]];
function styles() {
  const fonts = FONTS.map(([sz, b, i]) => `<font>${b ? "<b/>" : ""}${i ? "<i/>" : ""}<sz val="${sz}"/><name val="Times New Roman"/><family val="1"/><charset val="204"/></font>`).join("");
  const xfs = XF.map(([f, b, h, v, w, nf = 0, rot = 0]) => `<xf numFmtId="${nf}" fontId="${f}" fillId="0" borderId="${b}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"${nf ? ` applyNumberFormat="1"` : ""}><alignment${h ? ` horizontal="${h}"` : ""}${v ? ` vertical="${v}"` : ""}${w ? ` wrapText="1"` : ""}${rot ? ` textRotation="${rot}"` : ""}/></xf>`).join("");
  const thin = `<left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0#"/></numFmts><fonts count="${FONTS.length}">${fonts}</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>${thin}</border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${XF.length}">${xfs}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}

/**
 * sheet: {name, cells: Map("r,c" → {v, f, s}), merges: ["A1:F1"], widths: [..], heights: {r: pt}, landscape, fitWidth, repeatRows: [r1, r2]}
 */
export function buildSheet(sheet) {
  const rows = new Map();
  for (const [k, cell] of sheet.cells) { const [r, c] = k.split(",").map(Number); if (!rows.has(r)) rows.set(r, []); rows.get(r).push([c, cell]); }
  const rowXml = [...rows.keys()].sort((a, b) => a - b).map((r) => {
    const cs = rows.get(r).sort((a, b) => a[0] - b[0]).map(([c, cell]) => {
      const a = ` r="${ref(c, r)}" s="${cell.s || 0}"`;
      if (cell.f) return `<c${a}><f>${x(cell.f)}</f>${typeof cell.v === "number" ? `<v>${cell.v}</v>` : ""}</c>`;
      if (typeof cell.v === "number" && isFinite(cell.v)) return `<c${a}><v>${cell.v}</v></c>`;
      if (cell.v === undefined || cell.v === null || cell.v === "") return `<c${a}/>`;
      return `<c${a} t="inlineStr"><is><t xml:space="preserve">${x(cell.v)}</t></is></c>`;
    }).join("");
    const h = sheet.heights && sheet.heights[r];
    return `<row r="${r + 1}"${h ? ` ht="${h}" customHeight="1"` : ""}>${cs}</row>`;
  }).join("");
  const cols = (sheet.widths || []).map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("");
  const merges = (sheet.merges || []).length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : "";
  const ws = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView workbookViewId="0" zoomScale="100">${sheet.freeze ? `<pane xSplit="${sheet.freeze[0]}" ySplit="${sheet.freeze[1]}" topLeftCell="${ref(sheet.freeze[0], sheet.freeze[1])}" activePane="bottomRight" state="frozen"/>` : ""}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>
${cols ? `<cols>${cols}</cols>` : ""}<sheetData>${rowXml}</sheetData>${merges}
<printOptions horizontalCentered="1"/><pageMargins left="0.236" right="0.236" top="0.748" bottom="0.748" header="0.315" footer="0.315"/>
<pageSetup paperSize="9" orientation="${sheet.landscape ? "landscape" : "portrait"}" fitToWidth="1" fitToHeight="0"/></worksheet>`;
  const name = x(sheet.name || "Лист1");
  const pt = sheet.repeatRows ? `<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">'${name}'!$${sheet.repeatRows[0] + 1}:$${sheet.repeatRows[1] + 1}</definedName></definedNames>` : "";
  return zip([
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="${name}" sheetId="1" r:id="rId1"/></sheets>${pt}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ["xl/worksheets/sheet1.xml", ws],
    ["xl/styles.xml", styles()],
  ]);
}
