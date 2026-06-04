import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.0.227/build/pdf.mjs";
import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.0.227/build/pdf.worker.mjs";

const INVALID_SHEET_TITLE_CHARS = /[\[\]:*?/\\]/g;
const MAX_SHEET_TITLE_LENGTH = 31;
const COLUMN_TOLERANCE = 18;

const state = {
  file: null,
  pdf: null,
  sheets: [],
  busy: false,
  token: 0,
};

const elements = {
  dropZone: document.querySelector("#dropZone"),
  fileInput: document.querySelector("#fileInput"),
  fileBadge: document.querySelector("#fileBadge"),
  fileName: document.querySelector("#fileName"),
  fileSize: document.querySelector("#fileSize"),
  pageCount: document.querySelector("#pageCount"),
  textFallback: document.querySelector("#textFallback"),
  convertButton: document.querySelector("#convertButton"),
  downloadButton: document.querySelector("#downloadButton"),
  resetButton: document.querySelector("#resetButton"),
  progressBar: document.querySelector("#progressBar"),
  statusText: document.querySelector("#statusText"),
  previewBadge: document.querySelector("#previewBadge"),
  previewCanvas: document.querySelector("#previewCanvas"),
  emptyPreview: document.querySelector("#emptyPreview"),
  sheetBadge: document.querySelector("#sheetBadge"),
  summary: document.querySelector("#summary"),
  sheetList: document.querySelector("#sheetList"),
};

elements.fileInput.addEventListener("change", () => {
  const [file] = elements.fileInput.files;
  if (file) {
    void selectFile(file);
  }
});

elements.dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  elements.dropZone.classList.add("dragging");
});

elements.dropZone.addEventListener("dragleave", () => {
  elements.dropZone.classList.remove("dragging");
});

elements.dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  elements.dropZone.classList.remove("dragging");
  const [file] = event.dataTransfer.files;
  if (file) {
    void selectFile(file);
  }
});

elements.convertButton.addEventListener("click", () => {
  void convertSelectedPdf();
});

elements.downloadButton.addEventListener("click", () => {
  downloadWorkbook();
});

elements.resetButton.addEventListener("click", () => {
  resetApp();
});

async function selectFile(file) {
  const token = state.token + 1;
  state.token = token;

  if (!isPdfFile(file)) {
    resetApp();
    showError("Choose a PDF file.");
    return;
  }

  if (state.pdf) {
    await state.pdf.destroy();
  }

  state.file = file;
  state.pdf = null;
  state.sheets = [];
  elements.fileInput.value = "";

  setBadge(elements.fileBadge, "PDF", "good");
  elements.fileName.textContent = file.name;
  elements.fileSize.textContent = formatBytes(file.size);
  elements.pageCount.textContent = "-";
  elements.summary.textContent = "No workbook yet.";
  elements.sheetList.innerHTML = "";
  setBadge(elements.sheetBadge, "0");
  setProgress(0);
  clearPreview();
  setStatus("Reading PDF...");
  updateControls();

  try {
    setBusy(true);
    const pdf = await loadPdfDocument(file);
    if (token !== state.token) {
      await pdf.destroy();
      return;
    }
    state.pdf = pdf;
    elements.pageCount.textContent = String(pdf.numPages);
    setBadge(
      elements.previewBadge,
      `${pdf.numPages} page${pdf.numPages === 1 ? "" : "s"}`,
      "good",
    );
    await renderFirstPage(pdf);
    setStatus("Ready to convert.");
  } catch (error) {
    if (token === state.token) {
      state.file = null;
      state.pdf = null;
      state.sheets = [];
      elements.fileName.textContent = "None";
      elements.fileSize.textContent = "-";
      elements.pageCount.textContent = "-";
      renderSheets();
      clearPreview();
      showError(error instanceof Error ? error.message : "The PDF could not be read.");
    }
  } finally {
    if (token === state.token) {
      setBusy(false);
      updateControls();
    }
  }
}

async function convertSelectedPdf() {
  if (!state.file) {
    return;
  }

  setBusy(true);
  setStatus("Extracting text...");
  setProgress(0);
  state.sheets = [];
  renderSheets();
  updateControls();

  try {
    const pdf = state.pdf || (await loadPdfDocument(state.file));
    state.pdf = pdf;
    const extracted = [];

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const textContent = await page.getTextContent();
      const sheet = extractSheetFromPage(
        textContent,
        pageNumber,
        elements.textFallback.checked,
      );
      if (sheet) {
        extracted.push(sheet);
      }
      setProgress((pageNumber / pdf.numPages) * 100);
    }

    state.sheets = withUniqueSheetTitles(extracted);
    renderSheets();

    if (!state.sheets.length) {
      showError("No readable text was found. Run OCR on scanned PDFs first.");
      return;
    }

    const rowCount = state.sheets.reduce(
      (total, sheet) => total + sheet.rows.length,
      0,
    );
    setStatus(
      `Workbook ready: ${state.sheets.length} sheet${
        state.sheets.length === 1 ? "" : "s"
      }, ${rowCount} row${rowCount === 1 ? "" : "s"}.`,
    );
  } catch (error) {
    showError(error instanceof Error ? error.message : "Conversion failed.");
  } finally {
    setBusy(false);
    updateControls();
  }
}

function downloadWorkbook() {
  if (!state.sheets.length || !state.file) {
    return;
  }

  const workbook = XLSX.utils.book_new();
  for (const sheet of state.sheets) {
    const worksheet = XLSX.utils.aoa_to_sheet(sheet.rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.title);
  }

  const filename = outputFilename(state.file.name);
  if (typeof XLSX.writeFileXLSX === "function") {
    XLSX.writeFileXLSX(workbook, filename);
  } else {
    XLSX.writeFile(workbook, filename);
  }
  setStatus(`Downloaded ${filename}.`);
}

async function loadPdfDocument(file) {
  const data = await file.arrayBuffer();
  return pdfjsLib.getDocument({ data }).promise;
}

async function renderFirstPage(pdf) {
  const page = await pdf.getPage(1);
  const unscaled = page.getViewport({ scale: 1 });
  const frame = elements.previewCanvas.parentElement;
  const availableWidth = Math.max(frame.clientWidth - 28, 260);
  const cssScale = Math.min(1.25, availableWidth / unscaled.width);
  const pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
  const viewport = page.getViewport({ scale: cssScale * pixelRatio });
  const canvas = elements.previewCanvas;
  const context = canvas.getContext("2d");

  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  canvas.style.width = `${Math.ceil(viewport.width / pixelRatio)}px`;
  canvas.style.height = `${Math.ceil(viewport.height / pixelRatio)}px`;

  elements.emptyPreview.hidden = true;
  canvas.hidden = false;

  await page.render({ canvasContext: context, viewport }).promise;
}

function extractSheetFromPage(textContent, pageNumber, textFallback) {
  const lines = groupTextIntoLines(textContent.items);
  if (!lines.length) {
    return null;
  }

  const tableRows = extractTableRows(lines);
  if (tableRows.length) {
    return {
      pageNumber,
      source: "table",
      rows: tableRows,
    };
  }

  if (!textFallback) {
    return null;
  }

  const rows = lines
    .map((line) => [cleanCell(line.items.map((item) => item.text).join(" "))])
    .filter((row) => row[0]);

  return rows.length
    ? {
        pageNumber,
        source: "text",
        rows,
      }
    : null;
}

function groupTextIntoLines(items) {
  const positioned = items
    .map((item) => {
      const text = cleanCell(item.str);
      const transform = item.transform || [1, 0, 0, 1, 0, 0];
      const x = Number(transform[4]) || 0;
      const y = Number(transform[5]) || 0;
      const height =
        Math.hypot(Number(transform[2]) || 0, Number(transform[3]) || 0) ||
        Number(item.height) ||
        10;
      const width = Number(item.width) || text.length * height * 0.5;

      return {
        text,
        x,
        y,
        endX: x + width,
        height,
      };
    })
    .filter((item) => item.text);

  positioned.sort((left, right) => right.y - left.y || left.x - right.x);

  const lines = [];
  for (const item of positioned) {
    const tolerance = Math.max(2.5, Math.min(8, item.height * 0.38));
    const line = lines.find(
      (candidate) => Math.abs(candidate.y - item.y) <= tolerance,
    );

    if (!line) {
      lines.push({ y: item.y, height: item.height, items: [item] });
      continue;
    }

    line.items.push(item);
    line.y = average(line.items.map((candidate) => candidate.y));
    line.height = Math.max(line.height, item.height);
  }

  return lines
    .sort((left, right) => right.y - left.y)
    .map((line) => ({
      ...line,
      items: line.items.sort((left, right) => left.x - right.x),
    }));
}

function extractTableRows(lines) {
  const rows = lines
    .map((line) => ({ y: line.y, cells: mergeLineCells(line.items) }))
    .filter((row) => row.cells.length);

  const tableLikeRows = rows.filter((row) => row.cells.length >= 2);
  if (tableLikeRows.length < 2) {
    return [];
  }

  const columns = clusterColumns(tableLikeRows.flatMap((row) => row.cells));
  if (columns.length < 2) {
    return [];
  }

  const shapedRows = rows
    .filter((row) => row.cells.length >= 2)
    .map((row) => placeCellsInColumns(row.cells, columns))
    .filter((row) => row.some(Boolean));

  const supportedRows = shapedRows.filter((row) => row.filter(Boolean).length >= 2);
  if (supportedRows.length < 2) {
    return [];
  }

  return trimEmptyColumns(shapedRows);
}

function mergeLineCells(items) {
  if (!items.length) {
    return [];
  }

  const cells = [];
  let current = { ...items[0] };

  for (const item of items.slice(1)) {
    const gap = item.x - current.endX;
    const threshold = Math.max(
      4,
      Math.min(16, Math.max(current.height, item.height) * 0.72),
    );

    if (gap <= threshold) {
      current.text = cleanCell(`${current.text} ${item.text}`);
      current.endX = Math.max(current.endX, item.endX);
      current.height = Math.max(current.height, item.height);
    } else {
      cells.push(current);
      current = { ...item };
    }
  }

  cells.push(current);
  return cells.filter((cell) => cell.text);
}

function clusterColumns(cells) {
  const clusters = [];
  const sortedCells = [...cells].sort((left, right) => left.x - right.x);

  for (const cell of sortedCells) {
    const cluster = clusters.find(
      (candidate) => Math.abs(candidate.center - cell.x) <= COLUMN_TOLERANCE,
    );

    if (cluster) {
      cluster.center = (cluster.center * cluster.count + cell.x) / (cluster.count + 1);
      cluster.count += 1;
    } else {
      clusters.push({ center: cell.x, count: 1 });
    }
  }

  const minSupport = Math.max(2, Math.ceil(cells.length * 0.08));
  const supported = clusters.filter((cluster) => cluster.count >= minSupport);
  return (supported.length >= 2 ? supported : clusters).sort(
    (left, right) => left.center - right.center,
  );
}

function placeCellsInColumns(cells, columns) {
  const row = Array.from({ length: columns.length }, () => "");

  for (const cell of cells) {
    const index = nearestColumnIndex(cell.x, columns);
    row[index] = cleanCell(row[index] ? `${row[index]} ${cell.text}` : cell.text);
  }

  return row;
}

function nearestColumnIndex(x, columns) {
  let bestIndex = 0;
  let bestDistance = Number.POSITIVE_INFINITY;

  columns.forEach((column, index) => {
    const distance = Math.abs(column.center - x);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestIndex = index;
    }
  });

  return bestIndex;
}

function trimEmptyColumns(rows) {
  const keep = rows[0].map((_, index) => rows.some((row) => row[index]));
  return rows.map((row) => row.filter((_, index) => keep[index]));
}

function withUniqueSheetTitles(sheets) {
  const existing = new Set();
  return sheets.map((sheet) => {
    const title = uniqueSheetTitle(
      existing,
      `Page ${sheet.pageNumber} ${capitalize(sheet.source)} 1`,
    );
    return { ...sheet, title };
  });
}

function uniqueSheetTitle(existing, title) {
  const base = sanitizeSheetTitle(title);
  let candidate = base;
  let counter = 2;

  while (existing.has(candidate)) {
    const suffix = ` ${counter}`;
    candidate = `${base.slice(0, MAX_SHEET_TITLE_LENGTH - suffix.length)}${suffix}`;
    counter += 1;
  }

  existing.add(candidate);
  return candidate;
}

function sanitizeSheetTitle(title) {
  const sanitized = title.replace(INVALID_SHEET_TITLE_CHARS, "_").trim();
  return (sanitized || "Sheet").slice(0, MAX_SHEET_TITLE_LENGTH);
}

function renderSheets() {
  elements.sheetList.innerHTML = "";
  setBadge(
    elements.sheetBadge,
    String(state.sheets.length),
    state.sheets.length ? "good" : "",
  );

  if (!state.sheets.length) {
    elements.summary.textContent = "No workbook yet.";
    return;
  }

  const rowCount = state.sheets.reduce(
    (total, sheet) => total + sheet.rows.length,
    0,
  );
  elements.summary.textContent = `${state.sheets.length} sheet${
    state.sheets.length === 1 ? "" : "s"
  } ready, ${rowCount} row${rowCount === 1 ? "" : "s"} total.`;

  for (const sheet of state.sheets) {
    const item = document.createElement("article");
    item.className = "sheet-item";

    const titleRow = document.createElement("div");
    titleRow.className = "sheet-title-row";

    const title = document.createElement("h3");
    title.textContent = sheet.title;

    const source = document.createElement("span");
    source.className = "source";
    source.textContent = sheet.source;

    const meta = document.createElement("p");
    const columnCount = Math.max(...sheet.rows.map((row) => row.length), 0);
    meta.textContent = `${sheet.rows.length} row${
      sheet.rows.length === 1 ? "" : "s"
    } / ${columnCount} column${columnCount === 1 ? "" : "s"}`;

    const grid = previewGrid(sheet.rows);

    titleRow.append(title, source);
    item.append(titleRow, meta, grid);
    elements.sheetList.append(item);
  }
}

function previewGrid(rows) {
  const grid = document.createElement("div");
  const previewRows = rows.slice(0, 5);
  const columnCount = Math.min(
    Math.max(...previewRows.map((row) => row.length), 1),
    4,
  );

  grid.className = "sheet-grid";
  grid.style.gridTemplateColumns = `repeat(${columnCount}, minmax(0, 1fr))`;

  for (let rowIndex = 0; rowIndex < previewRows.length; rowIndex += 1) {
    const row = previewRows[rowIndex];
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const cell = document.createElement("span");
      cell.className = `sheet-cell${rowIndex === 0 ? " header" : ""}`;
      cell.textContent = row[columnIndex] || "";
      grid.append(cell);
    }
  }

  return grid;
}

function resetApp() {
  const token = state.token + 1;
  state.token = token;
  if (state.pdf) {
    void state.pdf.destroy();
  }
  state.file = null;
  state.pdf = null;
  state.sheets = [];
  state.busy = false;
  elements.fileInput.value = "";
  elements.fileName.textContent = "None";
  elements.fileSize.textContent = "-";
  elements.pageCount.textContent = "-";
  elements.summary.textContent = "No workbook yet.";
  elements.sheetList.innerHTML = "";
  setBadge(elements.fileBadge, "No file");
  setBadge(elements.sheetBadge, "0");
  setBadge(elements.previewBadge, "Waiting");
  setStatus("Choose a PDF to begin.");
  setProgress(0);
  clearPreview();
  updateControls();
}

function clearPreview() {
  const canvas = elements.previewCanvas;
  const context = canvas.getContext("2d");
  context.clearRect(0, 0, canvas.width, canvas.height);
  canvas.hidden = true;
  elements.emptyPreview.hidden = false;
}

function showError(message) {
  setStatus(message);
  setBadge(elements.fileBadge, state.file ? "PDF" : "No file", "error");
  setBadge(elements.previewBadge, "Error", "error");
  setProgress(0);
}

function updateControls() {
  elements.convertButton.disabled = !state.file || state.busy;
  elements.downloadButton.disabled = !state.sheets.length || state.busy;
  elements.resetButton.disabled = (!state.file && !state.sheets.length) || state.busy;
}

function setBusy(isBusy) {
  state.busy = isBusy;
  updateControls();
}

function setStatus(message) {
  elements.statusText.textContent = message;
}

function setProgress(percent) {
  elements.progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
}

function setBadge(element, text, tone = "") {
  element.textContent = text;
  element.className = `badge${tone ? ` ${tone}` : ""}`;
}

function isPdfFile(file) {
  return file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
}

function outputFilename(name) {
  const base = name.replace(/\.pdf$/i, "").trim() || "converted";
  return `${base}.xlsx`;
}

function formatBytes(bytes) {
  if (!bytes) {
    return "0 B";
  }
  const units = ["B", "KB", "MB", "GB"];
  const exponent = Math.min(
    Math.floor(Math.log(bytes) / Math.log(1024)),
    units.length - 1,
  );
  const value = bytes / 1024 ** exponent;
  return `${value.toFixed(value >= 10 || exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}

function cleanCell(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function average(values) {
  return values.reduce((total, value) => total + value, 0) / values.length;
}

function capitalize(value) {
  return value.charAt(0).toUpperCase() + value.slice(1);
}
