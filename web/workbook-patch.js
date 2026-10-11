// Writes appended rows straight into the workbook's XML parts so everything the
// browser does not touch (styles, validation, conditional formatting, tables,
// charts, macros) survives byte for byte. SheetJS is only used for reading.
import {
  strFromU8,
  strToU8,
  unzipSync,
  zipSync,
} from "https://cdn.jsdelivr.net/npm/fflate@0.8.2/esm/browser.js";

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_DOC_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_CONTENT_TYPES =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const WORKSHEET_REL_TYPE = `${NS_DOC_REL}/worksheet`;
const CALC_CHAIN_REL_TYPE = `${NS_DOC_REL}/calcChain`;
const WORKSHEET_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const BUILTIN_DATE_FORMAT_ID = 14;
const BUILTIN_DATE_FORMAT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
]);
const HIGHLIGHT_RGB = "FFFFF2CC";
// Elements that must follow <calcPr> inside <workbook>, per the schema order.
const AFTER_CALC_PR = [
  "oleSize",
  "customWorkbookViews",
  "pivotCaches",
  "smartTagPr",
  "smartTagTypes",
  "webPublishing",
  "fileRecoveryPr",
  "webPublishObjects",
  "extLst",
];
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

/**
 * @param {ArrayBuffer|Uint8Array} fileBytes original .xlsx / .xlsm
 * @param {object} patch
 * @param {string} patch.sheetName target sheet
 * @param {Array<{row:number, template:number|null, cells:Array}>} patch.rows
 *   1-based rows; cells are {col, kind: "number"|"date"|"string"|"formula", value}
 * @param {boolean} patch.highlight fill newly written value cells
 * @param {string} patch.logSheetName sheet that records every change
 * @param {Array<Array>} patch.logHeader header row used when the log is created
 * @param {Array<Array>} patch.logRows rows appended to the log sheet
 * @returns {Uint8Array}
 */
export function patchWorkbook(fileBytes, patch) {
  const files = unzipSync(new Uint8Array(fileBytes));
  const pkg = new Package(files);
  const workbook = pkg.xml("xl/workbook.xml");
  const workbookRels = pkg.xml("xl/_rels/workbook.xml.rels");
  const date1904 = ["1", "true"].includes(
    first(workbook, "workbookPr")?.getAttribute("date1904"),
  );
  const styles = files["xl/styles.xml"]
    ? new StyleBook(pkg.xml("xl/styles.xml"))
    : null;

  const sheetPath = worksheetPath(workbook, workbookRels, patch.sheetName);
  if (!sheetPath) {
    throw new Error(`Sheet ${patch.sheetName} was not found in the workbook.`);
  }
  writeRows(pkg.xml(sheetPath), patch.rows, {
    styles,
    date1904,
    highlight: patch.highlight,
  });

  if (patch.logRows.length) {
    let logPath = worksheetPath(workbook, workbookRels, patch.logSheetName);
    let logRows = patch.logRows;
    if (!logPath) {
      logPath = addWorksheet(pkg, workbook, workbookRels, patch.logSheetName);
      logRows = [patch.logHeader, ...patch.logRows];
    }
    const logSheet = pkg.xml(logPath);
    const start = lastRowNumber(logSheet) + 1;
    writeRows(
      logSheet,
      logRows.map((values, index) => ({
        row: start + index,
        template: null,
        cells: values.map((value, col) => ({
          col: col + 1,
          kind: typeof value === "number" ? "number" : "string",
          value,
        })),
      })),
      { styles: null, date1904, highlight: false },
    );
  }

  dropCalcChain(pkg, workbookRels);
  requestFullRecalc(workbook);
  return pkg.zip();
}

class Package {
  constructor(files) {
    this.files = files;
    this.docs = new Map();
  }

  xml(path) {
    if (!this.docs.has(path)) {
      const bytes = this.files[path];
      if (!bytes) {
        throw new Error(`The workbook is missing ${path}.`);
      }
      const doc = new DOMParser().parseFromString(
        strFromU8(bytes),
        "application/xml",
      );
      if (doc.getElementsByTagName("parsererror").length) {
        throw new Error(`The workbook part ${path} could not be read.`);
      }
      this.docs.set(path, doc);
    }
    return this.docs.get(path);
  }

  create(path, source) {
    this.files[path] = strToU8(source);
  }

  remove(path) {
    delete this.files[path];
    this.docs.delete(path);
  }

  zip() {
    const serializer = new XMLSerializer();
    for (const [path, doc] of this.docs) {
      let text = serializer.serializeToString(doc);
      if (!text.startsWith("<?xml")) {
        text = XML_DECLARATION + text;
      }
      this.files[path] = strToU8(text);
    }
    return zipSync(this.files, { level: 6 });
  }
}

class StyleBook {
  constructor(doc) {
    this.doc = doc;
    this.cellXfs = first(doc, "cellXfs");
    this.fills = first(doc, "fills");
    this.customFormats = new Map(
      [...doc.getElementsByTagNameNS(NS_MAIN, "numFmt")].map((format) => [
        Number(format.getAttribute("numFmtId")),
        format.getAttribute("formatCode") || "",
      ]),
    );
    this.cache = new Map();
    this.origins = new Map();
    this.highlightFillId = null;
  }

  xf(index) {
    return children(this.cellXfs, "xf")[index] || null;
  }

  isDateStyle(index) {
    const formatId = Number(this.xf(index)?.getAttribute("numFmtId") || 0);
    if (BUILTIN_DATE_FORMAT_IDS.has(formatId)) {
      return true;
    }
    const code = this.customFormats.get(formatId);
    if (!code) {
      return false;
    }
    const stripped = code
      .replace(/"[^"]*"/g, "")
      .replace(/\[[^\]]*\]/g, "")
      .replace(/\\./g, "");
    return /[dmy]/i.test(stripped);
  }

  /** Returns a style index based on `base` with a highlight fill and/or date format. */
  derive(base, { highlight, date }) {
    const wantsDate = date && !this.isDateStyle(base);
    if (!highlight && !wantsDate) {
      return base;
    }
    const key = `${base}|${highlight ? 1 : 0}|${wantsDate ? 1 : 0}`;
    if (this.cache.has(key)) {
      return this.cache.get(key);
    }

    const source = this.xf(base) || this.xf(0);
    const xf = source
      ? source.cloneNode(true)
      : this.doc.createElementNS(NS_MAIN, "xf");
    if (!source) {
      for (const name of ["numFmtId", "fontId", "fillId", "borderId", "xfId"]) {
        xf.setAttribute(name, "0");
      }
    }
    if (highlight) {
      xf.setAttribute("fillId", String(this.highlightFill()));
      xf.setAttribute("applyFill", "1");
    }
    if (wantsDate) {
      xf.setAttribute("numFmtId", String(BUILTIN_DATE_FORMAT_ID));
      xf.setAttribute("applyNumberFormat", "1");
    }
    this.cellXfs.appendChild(xf);
    const index = children(this.cellXfs, "xf").length - 1;
    this.cellXfs.setAttribute("count", String(index + 1));
    this.cache.set(key, index);
    this.origins.set(index, base);
    return index;
  }

  /** The style a derived index was made from, so highlights are not copied down. */
  origin(index) {
    return this.origins.get(index) ?? index;
  }

  highlightFill() {
    if (this.highlightFillId === null) {
      const fill = element(this.doc, "fill");
      const pattern = element(this.doc, "patternFill", { patternType: "solid" });
      pattern.appendChild(element(this.doc, "fgColor", { rgb: HIGHLIGHT_RGB }));
      pattern.appendChild(element(this.doc, "bgColor", { indexed: "64" }));
      fill.appendChild(pattern);
      this.fills.appendChild(fill);
      this.highlightFillId = children(this.fills, "fill").length - 1;
      this.fills.setAttribute("count", String(this.highlightFillId + 1));
    }
    return this.highlightFillId;
  }
}

function writeRows(sheet, rows, { styles, date1904, highlight }) {
  const sheetData = first(sheet, "sheetData");
  const rowIndex = new Map(
    children(sheetData, "row").map((row) => [Number(row.getAttribute("r")), row]),
  );
  let maxRow = 0;
  let maxCol = 0;

  for (const plan of rows) {
    const template = plan.template ? rowIndex.get(plan.template) : null;
    let row = rowIndex.get(plan.row);
    if (!row) {
      row = element(sheet, "row", { r: String(plan.row) });
      for (const name of ["ht", "customHeight", "s", "customFormat"]) {
        if (template?.hasAttribute(name)) {
          row.setAttribute(name, template.getAttribute(name));
        }
      }
      insertOrdered(sheetData, row, plan.row, (node) => Number(node.getAttribute("r")));
      rowIndex.set(plan.row, row);
    }
    row.removeAttribute("spans");

    // Carry the template row's formatting onto cells we do not otherwise touch.
    const templateStyles = new Map();
    for (const cell of template ? children(template, "c") : []) {
      const col = columnNumber(cell.getAttribute("r"));
      const raw = cell.getAttribute("s");
      const style = raw !== null && styles ? String(styles.origin(Number(raw))) : raw;
      if (style !== null) {
        templateStyles.set(col, style);
        if (!findCell(row, col)) {
          const styled = element(sheet, "c", {
            r: cellRef(col, plan.row),
            s: style,
          });
          insertOrdered(row, styled, col, cellColumn);
        }
      }
    }

    for (const write of plan.cells) {
      let cell = findCell(row, write.col);
      if (!cell) {
        cell = element(sheet, "c", { r: cellRef(write.col, plan.row) });
        insertOrdered(row, cell, write.col, cellColumn);
      }
      const base = Number(
        cell.getAttribute("s") ?? templateStyles.get(write.col) ?? row.getAttribute("s") ?? 0,
      );
      setCellValue(sheet, cell, write, { styles, base, date1904, highlight });
      maxCol = Math.max(maxCol, write.col);
    }
    maxRow = Math.max(maxRow, plan.row);
  }

  extendDimension(sheet, maxRow, maxCol);
}

function setCellValue(sheet, cell, write, { styles, base, date1904, highlight }) {
  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }
  cell.removeAttribute("t");

  let style = base;
  if (write.kind === "formula") {
    cell.appendChild(element(sheet, "f", {}, String(write.value)));
  } else if (write.kind === "number") {
    cell.appendChild(element(sheet, "v", {}, String(write.value)));
    style = styles ? styles.derive(base, { highlight }) : base;
  } else if (write.kind === "date") {
    cell.appendChild(element(sheet, "v", {}, String(dateSerial(write.value, date1904))));
    style = styles ? styles.derive(base, { highlight, date: true }) : base;
  } else {
    cell.setAttribute("t", "inlineStr");
    const inline = element(sheet, "is");
    const text = element(sheet, "t", {}, String(write.value).replace(INVALID_XML_CHARS, ""));
    text.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    inline.appendChild(text);
    cell.appendChild(inline);
    style = styles ? styles.derive(base, { highlight }) : base;
  }

  if (style) {
    cell.setAttribute("s", String(style));
  } else {
    cell.removeAttribute("s");
  }
}

/** Excel serial day number for a YYYY-MM-DD string. */
export function dateSerial(isoDate, date1904 = false) {
  const match = String(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return null;
  }
  const [year, month, day] = match.slice(1).map(Number);
  const utc = Date.UTC(year, month - 1, day);
  const check = new Date(utc);
  if (check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) {
    return null;
  }
  const serial = utc / 86400000 + 25569;
  return date1904 ? serial - 1462 : serial;
}

function worksheetPath(workbook, workbookRels, sheetName) {
  const sheet = [...workbook.getElementsByTagNameNS(NS_MAIN, "sheet")].find(
    (node) => node.getAttribute("name") === sheetName,
  );
  const relId = sheet?.getAttributeNS(NS_DOC_REL, "id");
  const rel = [...workbookRels.getElementsByTagNameNS(NS_PKG_REL, "Relationship")].find(
    (node) => node.getAttribute("Id") === relId,
  );
  return rel ? resolvePartPath(rel.getAttribute("Target")) : null;
}

function resolvePartPath(target) {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const parts = ["xl"];
  for (const segment of target.split("/")) {
    if (segment === "..") {
      parts.pop();
    } else if (segment && segment !== ".") {
      parts.push(segment);
    }
  }
  return parts.join("/");
}

function addWorksheet(pkg, workbook, workbookRels, sheetName) {
  let number = 1;
  while (pkg.files[`xl/worksheets/sheet${number}.xml`]) {
    number += 1;
  }
  const path = `xl/worksheets/sheet${number}.xml`;
  pkg.create(
    path,
    `${XML_DECLARATION}<worksheet xmlns="${NS_MAIN}"><sheetData/></worksheet>`,
  );

  const relationships = workbookRels.documentElement;
  const relIds = new Set(
    children(relationships, "Relationship", NS_PKG_REL).map((rel) => rel.getAttribute("Id")),
  );
  let relNumber = relIds.size + 1;
  while (relIds.has(`rId${relNumber}`)) {
    relNumber += 1;
  }
  const relId = `rId${relNumber}`;
  const rel = workbookRels.createElementNS(NS_PKG_REL, "Relationship");
  rel.setAttribute("Id", relId);
  rel.setAttribute("Type", WORKSHEET_REL_TYPE);
  rel.setAttribute("Target", `worksheets/sheet${number}.xml`);
  relationships.appendChild(rel);

  const sheets = first(workbook, "sheets");
  const sheetIds = children(sheets, "sheet").map((sheet) => Number(sheet.getAttribute("sheetId")));
  const sheet = element(workbook, "sheet", {
    name: sheetName,
    sheetId: String(Math.max(0, ...sheetIds) + 1),
  });
  sheet.setAttributeNS(NS_DOC_REL, "r:id", relId);
  sheets.appendChild(sheet);

  const types = pkg.xml("[Content_Types].xml");
  const override = types.createElementNS(NS_CONTENT_TYPES, "Override");
  override.setAttribute("PartName", `/${path}`);
  override.setAttribute("ContentType", WORKSHEET_CONTENT_TYPE);
  types.documentElement.appendChild(override);
  return path;
}

// Excel rebuilds the calculation chain on open; a stale one can trigger repairs.
function dropCalcChain(pkg, workbookRels) {
  if (!pkg.files["xl/calcChain.xml"]) {
    return;
  }
  pkg.remove("xl/calcChain.xml");
  for (const rel of [...workbookRels.getElementsByTagNameNS(NS_PKG_REL, "Relationship")]) {
    if (rel.getAttribute("Type") === CALC_CHAIN_REL_TYPE) {
      rel.parentNode.removeChild(rel);
    }
  }
  const types = pkg.xml("[Content_Types].xml");
  for (const override of [...types.getElementsByTagNameNS(NS_CONTENT_TYPES, "Override")]) {
    if (override.getAttribute("PartName") === "/xl/calcChain.xml") {
      override.parentNode.removeChild(override);
    }
  }
}

function requestFullRecalc(workbook) {
  let calcPr = first(workbook, "calcPr");
  if (!calcPr) {
    calcPr = element(workbook, "calcPr");
    const root = workbook.documentElement;
    const next = [...root.children].find((node) => AFTER_CALC_PR.includes(node.localName));
    root.insertBefore(calcPr, next || null);
  }
  calcPr.setAttribute("fullCalcOnLoad", "1");
}

function extendDimension(sheet, maxRow, maxCol) {
  const dimension = first(sheet, "dimension");
  if (!dimension || !maxRow) {
    return;
  }
  const [start, end = start] = (dimension.getAttribute("ref") || "A1").split(":");
  const endCol = Math.max(columnNumber(end), maxCol);
  const endRow = Math.max(rowNumber(end), maxRow);
  dimension.setAttribute("ref", `${start}:${cellRef(endCol, endRow)}`);
}

function lastRowNumber(sheet) {
  const rows = children(first(sheet, "sheetData"), "row");
  return Math.max(0, ...rows.map((row) => Number(row.getAttribute("r")) || 0));
}

function findCell(row, col) {
  return children(row, "c").find((cell) => cellColumn(cell) === col) || null;
}

function cellColumn(cell) {
  return columnNumber(cell.getAttribute("r"));
}

function insertOrdered(parent, node, key, keyOf) {
  const next = [...parent.children].find(
    (child) => child.localName === node.localName && keyOf(child) > key,
  );
  parent.insertBefore(node, next || null);
}

function columnNumber(ref) {
  const letters = String(ref || "").match(/^\$?([A-Z]+)/i)?.[1] || "A";
  return [...letters.toUpperCase()].reduce(
    (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
    0,
  );
}

function rowNumber(ref) {
  return Number(String(ref || "").match(/(\d+)$/)?.[1] || 1);
}

function cellRef(col, row) {
  let letters = "";
  let remaining = col;
  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + offset) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return `${letters}${row}`;
}

function first(doc, localName) {
  return doc.getElementsByTagNameNS(NS_MAIN, localName)[0] || null;
}

function children(parent, localName, namespace = NS_MAIN) {
  return [...(parent?.children || [])].filter(
    (node) => node.localName === localName && node.namespaceURI === namespace,
  );
}

function element(doc, name, attributes = {}, text) {
  const node = doc.createElementNS(NS_MAIN, name);
  for (const [key, value] of Object.entries(attributes)) {
    node.setAttribute(key, value);
  }
  if (text !== undefined) {
    node.textContent = text;
  }
  return node;
}
