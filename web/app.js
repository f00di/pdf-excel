import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.0.227/build/pdf.mjs";
import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.0.227/build/pdf.worker.mjs";

const INVALID_SHEET_TITLE_CHARS = /[\[\]:*?/\\]/g;
const MAX_SHEET_TITLE_LENGTH = 31;
const COLUMN_TOLERANCE = 18;
const LOG_SHEET_NAME = "_PDF_Append_Log";

const APPEND_FIELDS = [
  "transaction_date",
  "description",
  "debit",
  "credit",
  "amount",
  "balance",
  "reference_number",
  "notes",
];

const FIELD_LABELS = {
  transaction_date: "Date",
  description: "Description",
  debit: "Debit",
  credit: "Credit",
  amount: "Amount",
  balance: "Balance",
  reference_number: "Reference",
  notes: "Notes",
};

const FIELD_ALIASES = {
  transaction_date: ["date", "transaction date", "trans date", "posting date"],
  description: ["description", "details", "memo", "payee", "transaction"],
  debit: ["debit", "withdrawal", "withdrawals", "paid out", "charge", "spent"],
  credit: ["credit", "deposit", "deposits", "paid in", "received"],
  amount: ["amount", "transaction amount", "value"],
  balance: ["balance", "running balance", "available balance"],
  reference_number: ["reference", "reference number", "ref", "check", "trace"],
  notes: ["notes", "note", "comments", "remarks", "details"],
};

const TABLE_FIELD_ALIASES = {
  transaction_date: ["date", "transaction date", "posting date", "value date"],
  status: ["status"],
  description: ["description", "details", "memo", "transaction"],
  debit: ["debit", "withdrawal", "withdrawals", "paid out", "charge"],
  credit: ["credit", "deposit", "deposits", "paid in", "receipt"],
  amount: ["amount", "transaction amount", "value"],
  balance: ["balance", "available balance", "statement balance"],
  reference_number: ["ref", "reference", "confirmation", "trace", "check"],
};

const BOA_SECTION_TYPES = {
  "deposits and other additions": "credit",
  "atm and debit card subtractions": "debit",
  "other subtractions": "debit",
  "service fees": "debit",
};

const DATE_PATTERN =
  /\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{2,4})\b/i;
const AMOUNT_PATTERN =
  /-?\(?\$?\s*(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})\)?/g;
const BOA_TRANSACTION_LINE_PATTERN =
  /^(\d{2}\/\d{2}\/\d{2,4})\s+(.+?)\s+(-?\$?\d[\d,]*\.\d{2})$/;

const state = {
  file: null,
  files: [],
  pdf: null,
  sheets: [],
  transactions: [],
  rawText: "",
  workbookFile: null,
  workbook: null,
  workbookSheets: {},
  targetSheetName: "",
  mapping: {},
  appendedWorkbook: null,
  appendedFilename: "",
  appendSummary: null,
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
  workbookInput: document.querySelector("#workbookInput"),
  workbookPickText: document.querySelector("#workbookPickText"),
  workbookName: document.querySelector("#workbookName"),
  workbookSheetCount: document.querySelector("#workbookSheetCount"),
  targetSheet: document.querySelector("#targetSheet"),
  writeMode: document.querySelector("#writeMode"),
  mappingPreview: document.querySelector("#mappingPreview"),
  textFallback: document.querySelector("#textFallback"),
  convertButton: document.querySelector("#convertButton"),
  appendButton: document.querySelector("#appendButton"),
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
  const files = Array.from(elements.fileInput.files);
  if (files.length) {
    void selectPdfFiles(files);
  }
});

elements.workbookInput.addEventListener("change", () => {
  const [file] = elements.workbookInput.files;
  if (file) {
    void selectWorkbookFile(file);
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
  const files = Array.from(event.dataTransfer.files);
  if (files.length) {
    void selectPdfFiles(files);
  }
});

elements.targetSheet.addEventListener("change", () => {
  state.targetSheetName = elements.targetSheet.value;
  state.mapping = {};
  renderMappingPreview();
  updateControls();
});

elements.writeMode.addEventListener("change", () => {
  state.appendedWorkbook = null;
  state.appendSummary = null;
  renderSheets();
  updateControls();
});

elements.convertButton.addEventListener("click", () => {
  void convertSelectedPdf();
});

elements.appendButton.addEventListener("click", () => {
  void appendToWorkbook();
});

elements.downloadButton.addEventListener("click", () => {
  downloadWorkbook();
});

elements.resetButton.addEventListener("click", () => {
  resetApp();
});

async function selectPdfFiles(files) {
  const token = state.token + 1;
  state.token = token;

  if (!files.every(isPdfFile)) {
    resetPdfState();
    showError("Choose PDF files only.");
    return;
  }

  if (state.pdf) {
    await state.pdf.destroy();
  }

  const [file] = files;
  state.file = file;
  state.files = files;
  state.pdf = null;
  state.sheets = [];
  state.transactions = [];
  state.rawText = "";
  state.appendedWorkbook = null;
  state.appendSummary = null;
  elements.fileInput.value = "";

  setBadge(elements.fileBadge, "PDF", "good");
  elements.fileName.textContent =
    files.length === 1 ? file.name : `${file.name} + ${files.length - 1} more`;
  elements.fileSize.textContent = formatBytes(
    files.reduce((total, candidate) => total + candidate.size, 0),
  );
  elements.pageCount.textContent = "-";
  setProgress(0);
  clearPreview();
  renderSheets();
  renderMappingPreview();
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
    setStatus(
      files.length === 1
        ? "Ready to convert or append."
        : `${files.length} PDFs ready to convert or append.`,
    );
  } catch (error) {
    if (token === state.token) {
      resetPdfState();
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

async function selectWorkbookFile(file) {
  if (!isWorkbookFile(file)) {
    resetWorkbookState();
    showError("Choose an .xlsx or .xlsm workbook.");
    return;
  }

  setBusy(true);
  setStatus("Reading workbook...");
  updateControls();

  try {
    const workbook = await readWorkbookFile(file);
    state.workbookFile = file;
    state.workbook = workbook;
    state.workbookSheets = analyzeWorkbook(workbook);
    state.targetSheetName = defaultTargetSheet(workbook.SheetNames);
    state.mapping = {};
    state.appendedWorkbook = null;
    state.appendSummary = null;
    renderWorkbookControls();
    renderMappingPreview();
    renderSheets();
    setStatus(`Workbook ready: ${file.name}.`);
  } catch (error) {
    resetWorkbookState();
    showError(error instanceof Error ? error.message : "The workbook could not be read.");
  } finally {
    setBusy(false);
    updateControls();
  }
}

async function convertSelectedPdf() {
  if (!state.file) {
    return;
  }

  setBusy(true);
  setStatus("Extracting PDF data...");
  setProgress(0);
  state.appendedWorkbook = null;
  state.appendSummary = null;
  updateControls();

  try {
    await extractSelectedPdfData();
    renderSheets();

    if (!state.sheets.length) {
      showError("No readable text was found. Run OCR on scanned PDFs first.");
      return;
    }

    const rowCount = state.sheets.reduce(
      (total, sheet) => total + sheet.rows.length,
      0,
    );
    const transactionText = state.transactions.length
      ? ` ${state.transactions.length} transaction row${
          state.transactions.length === 1 ? "" : "s"
        } detected.`
      : "";
    setStatus(
      `Workbook ready: ${state.sheets.length} sheet${
        state.sheets.length === 1 ? "" : "s"
      }, ${rowCount} row${rowCount === 1 ? "" : "s"}.${transactionText}`,
    );
  } catch (error) {
    showError(error instanceof Error ? error.message : "Conversion failed.");
  } finally {
    setBusy(false);
    updateControls();
  }
}

async function appendToWorkbook() {
  if (!state.file || !state.workbookFile || !state.targetSheetName) {
    return;
  }

  setBusy(true);
  setStatus("Preparing append...");
  setProgress(0);
  updateControls();

  try {
    await extractSelectedPdfData();
    if (!state.transactions.length) {
      showError("No transaction rows were detected in the PDF.");
      return;
    }

    const workbook = await readWorkbookFile(state.workbookFile);
    const analyses = analyzeWorkbook(workbook);
    const sheet = analyses[state.targetSheetName];
    if (!sheet) {
      showError("The selected sheet was not found in the workbook.");
      return;
    }

    const mapping = normalizedMapping(sheet);
    const validation = validateAppendPlan(state.transactions, sheet, mapping);
    const errors = validation.filter((issue) => issue.severity === "error");
    if (errors.length) {
      showError(errors.map((issue) => issue.message).join(" "));
      renderMappingPreview(validation);
      return;
    }

    const summary = appendTransactionsToWorkbook(
      workbook,
      sheet,
      state.transactions,
      mapping,
      elements.writeMode.value,
    );

    state.workbook = workbook;
    state.workbookSheets = analyzeWorkbook(workbook);
    state.appendedWorkbook = workbook;
    state.appendedFilename = updatedWorkbookFilename(state.workbookFile.name);
    state.appendSummary = { ...summary, validation };
    renderWorkbookControls();
    renderMappingPreview(validation);
    renderSheets();
    setProgress(100);
    setStatus(
      `Updated workbook ready: ${summary.rowsAdded} row${
        summary.rowsAdded === 1 ? "" : "s"
      } added, ${summary.rowsSkipped} skipped.`,
    );
  } catch (error) {
    showError(error instanceof Error ? error.message : "Append failed.");
  } finally {
    setBusy(false);
    updateControls();
  }
}

async function extractSelectedPdfData() {
  const extractedSheets = [];
  const allRawText = [];
  const allTransactions = [];
  const files = state.files.length ? state.files : [state.file];

  for (let fileIndex = 0; fileIndex < files.length; fileIndex += 1) {
    const file = files[fileIndex];
    const pdf =
      file === state.file && state.pdf ? state.pdf : await loadPdfDocument(file);
    if (file === state.file) {
      state.pdf = pdf;
    }

    const pageTables = [];
    const rawTextParts = [];

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const textContent = await page.getTextContent();
      const lines = groupTextIntoLines(textContent.items);
      const rawLines = lines.map((line) =>
        cleanCell(line.items.map((item) => item.text).join(" ")),
      );
      const tableRows = extractTableRows(lines);
      const sheet = sheetFromPageRows(
        tableRows,
        rawLines,
        pageNumber,
        elements.textFallback.checked,
      );

      rawTextParts.push(`--- Page ${pageNumber} ---\n${rawLines.join("\n")}`);
      if (tableRows.length) {
        pageTables.push({ page: pageNumber, rows: tableRows });
      }
      if (sheet) {
        extractedSheets.push({
          ...sheet,
          fileName: file.name,
        });
      }
      setProgress(((fileIndex + pageNumber / pdf.numPages) / files.length) * 100);
    }

    const fileRawText = rawTextParts.join("\n\n").trim();
    allRawText.push(`--- File ${file.name} ---\n${fileRawText}`);
    allTransactions.push(...extractTransactions(file.name, fileRawText, pageTables));

    if (file !== state.file) {
      await pdf.destroy();
    }
  }

  state.sheets = withUniqueSheetTitles(extractedSheets);
  state.rawText = allRawText.join("\n\n").trim();
  state.transactions = deduplicateTransactions(allTransactions);
  renderMappingPreview();
}

function downloadWorkbook() {
  if (state.appendedWorkbook) {
    writeWorkbookFile(state.appendedWorkbook, state.appendedFilename);
    setStatus(`Downloaded ${state.appendedFilename}.`);
    return;
  }

  if (!state.sheets.length || !state.file) {
    return;
  }

  const workbook = XLSX.utils.book_new();
  for (const sheet of state.sheets) {
    const worksheet = XLSX.utils.aoa_to_sheet(sheet.rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.title);
  }

  const filename =
    state.files.length > 1 ? "converted_pdfs.xlsx" : outputFilename(state.file.name);
  writeWorkbookFile(workbook, filename);
  setStatus(`Downloaded ${filename}.`);
}

function writeWorkbookFile(workbook, filename) {
  if (typeof XLSX.writeFileXLSX === "function") {
    XLSX.writeFileXLSX(workbook, filename);
  } else {
    XLSX.writeFile(workbook, filename);
  }
}

async function loadPdfDocument(file) {
  const data = await file.arrayBuffer();
  return pdfjsLib.getDocument({ data }).promise;
}

async function readWorkbookFile(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, {
    type: "array",
    cellDates: true,
    cellStyles: true,
    cellFormula: true,
  });
  if (!workbook.SheetNames.length) {
    throw new Error("The workbook does not contain any sheets.");
  }
  return workbook;
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

function sheetFromPageRows(tableRows, rawLines, pageNumber, textFallback) {
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

  const rows = rawLines.map((line) => [cleanCell(line)]).filter((row) => row[0]);
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

function extractTransactions(fileName, rawText, tables) {
  const metadata = extractStatementMetadata(rawText);
  const boaTransactions = transactionsFromBoaStatement(fileName, rawText, metadata);
  if (boaTransactions.length) {
    return deduplicateTransactions(boaTransactions);
  }

  const tableTransactions = transactionsFromTables(fileName, tables, metadata);
  if (tableTransactions.length) {
    return deduplicateTransactions(tableTransactions);
  }

  return deduplicateTransactions(transactionsFromLines(fileName, rawText, metadata));
}

function extractStatementMetadata(rawText) {
  const metadata = {};
  const condensed = cleanCell(rawText);
  const accountMatch =
    condensed.match(/Account\s+number:\s*([0-9 ]{6,})/i) ||
    condensed.match(/Account\s*#\s*([0-9 ]{6,})/i) ||
    condensed.match(/(?:account|acct)\s*(?:number|no\.?|#)?\s*[:\-]?\s*([*Xx0-9\- ]{4,})/i);

  if (accountMatch) {
    metadata.account_number = cleanCell(accountMatch[1]);
  }

  const periodMatch = condensed.match(
    /\bfor\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})\s+to\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})\b/i,
  );
  if (periodMatch) {
    metadata.statement_period = `${periodMatch[1]} to ${periodMatch[2]}`;
    metadata.statement_start = normalizeDate(periodMatch[1]);
    metadata.statement_end = normalizeDate(periodMatch[2]);
  }

  const openingMatch = condensed.match(
    /(?:opening|beginning)\s+balance(?:\s+on\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})?\s*[:\-]?\s*(\$?\(?-?[\d,]+\.\d{2}\)?)/i,
  );
  if (openingMatch) {
    metadata.opening_balance = parseAmount(openingMatch[1]);
  }

  const closingMatch = condensed.match(
    /(?:closing|ending)\s+balance(?:\s+on\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})?\s*[:\-]?\s*(\$?\(?-?[\d,]+\.\d{2}\)?)/i,
  );
  if (closingMatch) {
    metadata.closing_balance = parseAmount(closingMatch[1]);
  }

  return metadata;
}

function transactionsFromBoaStatement(fileName, rawText, metadata) {
  const normalized = rawText.toLowerCase();
  if (
    !normalized.includes("bank of america") ||
    !normalized.includes("deposits and other additions")
  ) {
    return [];
  }

  const rows = [];
  const totals = {};
  let currentPage = null;
  let currentSection = "";
  let activeTable = false;
  let pending = null;

  const finishPending = () => {
    if (pending) {
      rows.push(pending);
      pending = null;
    }
  };

  for (const rawLine of rawText.split(/\r?\n/)) {
    const line = cleanCell(rawLine);
    if (!line) {
      continue;
    }

    const pageMatch = line.match(/^--- Page (\d+)/);
    if (pageMatch) {
      currentPage = Number(pageMatch[1]);
      continue;
    }

    const totalMatch = line.match(/^Total\s+(.+?)\s+(-?\$?[\d,]+\.\d{2})$/i);
    if (totalMatch) {
      finishPending();
      const totalSection = normalizeBoaSection(totalMatch[1]);
      if (BOA_SECTION_TYPES[totalSection]) {
        const total = parseAmount(totalMatch[2]);
        if (total !== null) {
          totals[totalSection] = Math.abs(total);
        }
      }
      activeTable = false;
      continue;
    }

    const section = normalizeBoaSection(line);
    if (BOA_SECTION_TYPES[section]) {
      finishPending();
      currentSection = section;
      activeTable = false;
      continue;
    }

    if (
      currentSection &&
      ["date description amount", "date transaction description amount"].includes(
        line.toLowerCase(),
      )
    ) {
      activeTable = true;
      continue;
    }

    if (!activeTable) {
      continue;
    }

    if (isBoaNoiseLine(line)) {
      finishPending();
      activeTable = line.toLowerCase().includes("continued on the next page");
      continue;
    }

    const transactionMatch = line.match(BOA_TRANSACTION_LINE_PATTERN);
    if (transactionMatch) {
      finishPending();
      pending = {
        dateText: transactionMatch[1],
        descriptionParts: [transactionMatch[2]],
        amountText: transactionMatch[3],
        amount: parseAmount(transactionMatch[3]),
        section: currentSection,
        page: currentPage,
      };
      continue;
    }

    if (pending) {
      pending.descriptionParts.push(line);
    }
  }

  finishPending();
  if (!rows.length) {
    return [];
  }

  const sectionSums = {};
  for (const row of rows) {
    if (row.amount === null) {
      continue;
    }
    sectionSums[row.section] = (sectionSums[row.section] || 0) + Math.abs(row.amount);
  }

  return rows
    .map((row) => {
      const amount = row.amount;
      if (amount === null) {
        return null;
      }
      const sectionType = BOA_SECTION_TYPES[row.section] || "";
      const descriptionBody = cleanCell(row.descriptionParts.join(" "));
      const amountText = row.amountText;
      const totalVerified =
        totals[row.section] !== undefined &&
        Math.abs(round2(sectionSums[row.section]) - totals[row.section]) <= 0.02;
      const debit = amount < 0 || sectionType === "debit" ? Math.abs(amount) : null;
      const credit = sectionType === "credit" && amount >= 0 ? Math.abs(amount) : null;

      return {
        transaction_date: normalizeDate(row.dateText),
        description: cleanCell(`${row.dateText} ${descriptionBody} ${amountText}`),
        debit,
        credit,
        amount,
        balance: null,
        reference_number: extractReferenceNumber(descriptionBody),
        account_number: metadata.account_number || "",
        statement_period: metadata.statement_period || "",
        opening_balance: metadata.opening_balance ?? null,
        closing_balance: metadata.closing_balance ?? null,
        notes: `BoA section: ${row.section}${
          totalVerified ? "; section total verified" : ""
        }`,
        source_pdf: fileName,
        source_page: row.page,
        raw_text: cleanCell(`${row.dateText} ${descriptionBody} ${amountText}`),
        confidence: totalVerified ? 0.96 : 0.91,
      };
    })
    .filter(Boolean);
}

function transactionsFromTables(fileName, tables, metadata) {
  const transactions = [];
  for (const table of tables) {
    const [headerIndex, mapping] = detectTableHeader(table.rows);
    if (headerIndex === null) {
      continue;
    }

    for (const row of table.rows.slice(headerIndex + 1)) {
      const transaction = transactionFromTableRow(
        row,
        mapping,
        fileName,
        table.page,
        metadata,
      );
      if (transaction) {
        transactions.push(transaction);
      }
    }
  }
  return transactions;
}

function detectTableHeader(rows) {
  let bestIndex = null;
  let bestMapping = {};
  let bestScore = 0;

  rows.slice(0, 8).forEach((row, rowIndex) => {
    const mapping = {};
    row.forEach((value, colIndex) => {
      const normalized = normalizeLabel(value);
      Object.entries(TABLE_FIELD_ALIASES).forEach(([field, aliases]) => {
        if (!mapping[field] && aliases.some((alias) => normalized.includes(alias))) {
          mapping[field] = colIndex;
        }
      });
    });

    let score = Object.keys(mapping).length;
    if (mapping.transaction_date !== undefined) {
      score += 2;
    }
    if (mapping.description !== undefined) {
      score += 2;
    }
    if (score > bestScore) {
      bestIndex = rowIndex;
      bestMapping = mapping;
      bestScore = score;
    }
  });

  return bestScore >= 3 ? [bestIndex, bestMapping] : [null, {}];
}

function transactionFromTableRow(row, mapping, fileName, page, metadata) {
  const value = (field) => {
    const index = mapping[field];
    return index === undefined || index >= row.length ? "" : cleanCell(row[index]);
  };

  const date = normalizeDate(value("transaction_date"));
  const description = value("description");
  const amount = parseAmount(value("amount"));
  let debit = parseAmount(value("debit"));
  let credit = parseAmount(value("credit"));
  const balance = parseAmount(value("balance"));

  if (debit === null && credit === null && amount !== null) {
    if (amount < 0) {
      debit = Math.abs(amount);
    } else {
      credit = Math.abs(amount);
    }
  }

  if (!date || !description || isBalanceOnlyDescription(description)) {
    return null;
  }
  if (debit === null && credit === null && amount === null && balance === null) {
    return null;
  }
  if (credit !== null && credit < 0) {
    debit = Math.abs(credit);
    credit = null;
  }
  if (debit !== null && debit < 0) {
    debit = Math.abs(debit);
  }

  return {
    transaction_date: date,
    description,
    debit,
    credit,
    amount,
    balance,
    status: value("status"),
    reference_number: value("reference_number") || extractReferenceNumber(description),
    account_number: metadata.account_number || "",
    statement_period: metadata.statement_period || "",
    opening_balance: metadata.opening_balance ?? null,
    closing_balance: metadata.closing_balance ?? null,
    notes: "",
    source_pdf: fileName,
    source_page: page,
    raw_text: row.join(" | "),
    confidence: 0.86,
  };
}

function transactionsFromLines(fileName, rawText, metadata) {
  const transactions = [];
  let currentPage = null;

  for (const rawLine of rawText.split(/\r?\n/)) {
    const pageMatch = rawLine.match(/^--- Page (\d+)/);
    if (pageMatch) {
      currentPage = Number(pageMatch[1]);
      continue;
    }

    const line = cleanCell(rawLine);
    if (!line) {
      continue;
    }

    const dateMatch = line.match(DATE_PATTERN);
    const amountMatches = [...line.matchAll(AMOUNT_PATTERN)];
    if (!dateMatch || !amountMatches.length) {
      continue;
    }

    const dateText = dateMatch[1];
    const amountIndex = amountMatches.length >= 2 ? amountMatches.length - 2 : 0;
    const amountText = amountMatches[amountIndex][0];
    const amount = parseAmount(amountText);
    if (amount === null) {
      continue;
    }

    const balance =
      amountMatches.length >= 2
        ? parseAmount(amountMatches[amountMatches.length - 1][0])
        : null;
    const description = cleanCell(
      line.slice(dateMatch.index + dateText.length, amountMatches[amountIndex].index),
    ).replace(/^[-:\s]+|[-:\s]+$/g, "");

    if (!description || isBalanceOnlyDescription(description)) {
      continue;
    }

    const upperLine = line.toUpperCase();
    const debit =
      amount < 0 || amountTextIsNegative(amountText) || debitLikeLine(upperLine)
        ? Math.abs(amount)
        : null;
    const credit = debit === null ? Math.abs(amount) : null;

    transactions.push({
      transaction_date: normalizeDate(dateText),
      description,
      debit,
      credit,
      amount,
      balance,
      reference_number: extractReferenceNumber(description),
      account_number: metadata.account_number || "",
      statement_period: metadata.statement_period || "",
      opening_balance: metadata.opening_balance ?? null,
      closing_balance: metadata.closing_balance ?? null,
      notes: "",
      source_pdf: fileName,
      source_page: currentPage,
      raw_text: line,
      confidence: 0.68,
    });
  }

  return transactions;
}

function analyzeWorkbook(workbook) {
  const analyses = {};
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false,
    });
    const headerRowIndex = detectHeaderRow(rows);
    const headers = extractHeaders(rows, headerRowIndex);
    const formulas = detectFormulaColumns(worksheet);

    analyses[sheetName] = {
      sheetName,
      rows,
      headerRow: headerRowIndex + 1,
      headerRowIndex,
      headers,
      formulas,
      maxRow: rows.length,
      maxColumn: Math.max(...rows.map((row) => row.length), 1),
      nextRow: findNextAppendRow(rows, headers, headerRowIndex, formulas),
    };
  });
  return analyses;
}

function detectHeaderRow(rows) {
  let bestIndex = 0;
  let bestScore = -1;

  rows.slice(0, 25).forEach((row, index) => {
    const values = row.map(cleanCell).filter(Boolean);
    if (!values.length) {
      return;
    }

    const keywordHits = values.filter((value) =>
      Object.values(FIELD_ALIASES).some((aliases) =>
        aliases.some((alias) => normalizeLabel(value).includes(alias)),
      ),
    ).length;
    const textish = values.filter((value) => Number.isNaN(Number(value))).length;
    const nextRow = rows[index + 1] || [];
    const nextDataScore = nextRow.filter((value) => cleanCell(value)).length;
    const uniqueRatio = new Set(values).size / Math.max(values.length, 1);
    const score =
      values.length +
      keywordHits * 3 +
      textish * 0.5 +
      uniqueRatio +
      Math.min(nextDataScore, values.length) * 0.35;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestIndex;
}

function extractHeaders(rows, headerRowIndex) {
  const headerRow = rows[headerRowIndex] || [];
  const headers = {};
  headerRow.forEach((value, index) => {
    const header = cleanCell(value);
    if (header) {
      headers[indexToColumn(index)] = header;
    }
  });
  return primaryHeaderBlock(headers);
}

function primaryHeaderBlock(headers) {
  const ordered = Object.keys(headers).sort(
    (left, right) => columnToIndex(left) - columnToIndex(right),
  );
  const primary = {};
  let previousIndex = null;
  for (const column of ordered) {
    const index = columnToIndex(column);
    if (previousIndex !== null && index > previousIndex + 1) {
      break;
    }
    primary[column] = headers[column];
    previousIndex = index;
  }
  return primary;
}

function detectFormulaColumns(worksheet) {
  const formulas = {};
  const range = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : null;
  if (!range) {
    return formulas;
  }

  for (let col = range.s.c; col <= range.e.c; col += 1) {
    let count = 0;
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
      if (cell?.f) {
        count += 1;
      }
    }
    formulas[indexToColumn(col)] = count;
  }
  return formulas;
}

function findNextAppendRow(rows, headers, headerRowIndex, formulas) {
  const dateColumn = Object.entries(headers).find(([, header]) => {
    const normalized = normalizeLabel(header);
    return normalized === "date" || normalized.includes("transaction date");
  })?.[0];

  if (dateColumn) {
    const dateIndex = columnToIndex(dateColumn);
    let lastDateRowIndex = headerRowIndex;
    for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
      if (cleanCell(rows[rowIndex]?.[dateIndex])) {
        lastDateRowIndex = rowIndex;
      }
    }
    if (lastDateRowIndex > headerRowIndex) {
      return lastDateRowIndex + 2;
    }
  }

  const entryIndexes = Object.keys(headers)
    .filter((column) => !(formulas[column] >= 3))
    .map(columnToIndex);
  let lastDataRowIndex = headerRowIndex;
  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    if (entryIndexes.some((index) => cleanCell(rows[rowIndex]?.[index]))) {
      lastDataRowIndex = rowIndex;
    }
  }
  return lastDataRowIndex + 2;
}

function suggestMapping(sheet) {
  const mapping = {};
  const usedColumns = new Set();

  APPEND_FIELDS.forEach((field) => {
    let bestColumn = "";
    let bestScore = 0;
    Object.entries(sheet.headers).forEach(([column, header]) => {
      if (usedColumns.has(column)) {
        return;
      }
      if (field === "balance" && sheet.formulas[column] > 0) {
        return;
      }
      const score = scoreHeader(field, header);
      if (score > bestScore) {
        bestColumn = column;
        bestScore = score;
      }
    });

    const minimumScore = field === "notes" || field === "reference_number" ? 0.6 : 0.45;
    if (bestColumn && bestScore >= minimumScore) {
      mapping[field] = bestColumn;
      usedColumns.add(bestColumn);
    }
  });

  return mapping;
}

function scoreHeader(field, header) {
  const normalizedHeader = normalizeLabel(header);
  const aliases = FIELD_ALIASES[field].map(normalizeLabel);
  if (aliases.includes(normalizedHeader)) {
    return 1;
  }
  if (aliases.some((alias) => normalizedHeader.includes(alias))) {
    return 0.88;
  }

  const headerTokens = new Set(normalizedHeader.split(" ").filter(Boolean));
  return aliases.reduce((best, alias) => {
    const aliasTokens = alias.split(" ").filter(Boolean);
    const overlap = aliasTokens.filter((token) => headerTokens.has(token)).length;
    return Math.max(best, overlap / Math.max(aliasTokens.length, 1));
  }, 0);
}

function normalizedMapping(sheet) {
  if (!Object.keys(state.mapping).length) {
    state.mapping = suggestMapping(sheet);
  }
  return Object.fromEntries(
    Object.entries(state.mapping).filter(([, column]) => column && sheet.headers[column]),
  );
}

function validateAppendPlan(transactions, sheet, mapping) {
  const issues = [];
  if (!mapping.transaction_date) {
    issues.push({ severity: "error", message: "Map the Date field before appending." });
  }
  if (!mapping.description) {
    issues.push({
      severity: "error",
      message: "Map the Description field before appending.",
    });
  }
  if (!mapping.amount && !mapping.debit && !mapping.credit) {
    issues.push({
      severity: "error",
      message: "Map Amount, Debit, or Credit before appending.",
    });
  }

  Object.entries(mapping).forEach(([field, column]) => {
    if (!sheet.headers[column]) {
      issues.push({
        severity: "error",
        message: `${FIELD_LABELS[field]} is mapped to a missing column.`,
      });
    }
  });

  transactions.forEach((transaction, index) => {
    if (!transaction.transaction_date) {
      issues.push({
        severity: "warning",
        message: `Row ${index + 1} has no transaction date.`,
      });
    }
    if (!transaction.description) {
      issues.push({
        severity: "warning",
        message: `Row ${index + 1} has no description.`,
      });
    }
  });

  return issues;
}

function appendTransactionsToWorkbook(workbook, sheet, transactions, mapping, writeMode) {
  const worksheet = workbook.Sheets[sheet.sheetName];
  const range = worksheet["!ref"]
    ? XLSX.utils.decode_range(worksheet["!ref"])
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  const mappedColumns = new Set(Object.values(mapping));
  const existingKeys =
    writeMode === "skip_duplicates"
      ? existingTransactionKeys(worksheet, sheet, mapping)
      : new Set();
  const changes = [];
  let rowsAdded = 0;
  let rowsSkipped = 0;
  let targetRowIndex = Math.max(sheet.nextRow - 1, sheet.headerRowIndex + 1);

  transactions.forEach((transaction) => {
    const transactionKey = transactionKeyFromObject(transaction);
    if (writeMode === "skip_duplicates" && existingKeys.has(transactionKey)) {
      rowsSkipped += 1;
      return;
    }

    if (writeMode !== "overwrite") {
      while (rowHasMappedValues(worksheet, targetRowIndex, mappedColumns)) {
        targetRowIndex += 1;
      }
    }

    copyRowTemplate(
      worksheet,
      Math.max(sheet.headerRowIndex + 1, targetRowIndex - 1),
      targetRowIndex,
      mappedColumns,
    );

    let wroteAny = false;
    Object.entries(mapping).forEach(([field, column]) => {
      const value = transactionFieldValue(transaction, field);
      if (value === null || value === undefined || value === "") {
        return;
      }

      const colIndex = columnToIndex(column);
      const address = XLSX.utils.encode_cell({ r: targetRowIndex, c: colIndex });
      const previous = worksheet[address] || {};
      worksheet[address] = {
        ...previous,
        ...cellForValue(field, value),
        s: highlightedStyle(previous.s),
      };
      changes.push({
        sheet: sheet.sheetName,
        cell: address,
        field,
        value,
        source_pdf: transaction.source_pdf,
      });
      wroteAny = true;
      range.e.r = Math.max(range.e.r, targetRowIndex);
      range.e.c = Math.max(range.e.c, colIndex);
    });

    if (wroteAny) {
      rowsAdded += 1;
      existingKeys.add(transactionKey);
      targetRowIndex += 1;
    } else {
      rowsSkipped += 1;
    }
  });

  worksheet["!ref"] = XLSX.utils.encode_range(range);
  appendLogSheet(workbook, changes);
  return { rowsAdded, rowsSkipped, changes };
}

function copyRowTemplate(worksheet, sourceRowIndex, targetRowIndex, mappedColumns) {
  if (sourceRowIndex < 0 || sourceRowIndex === targetRowIndex) {
    return;
  }

  const range = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : null;
  if (!range) {
    return;
  }

  for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
    const column = indexToColumn(colIndex);
    const sourceAddress = XLSX.utils.encode_cell({ r: sourceRowIndex, c: colIndex });
    const targetAddress = XLSX.utils.encode_cell({ r: targetRowIndex, c: colIndex });
    const sourceCell = worksheet[sourceAddress];
    if (!sourceCell) {
      continue;
    }

    const targetCell = worksheet[targetAddress] || {};
    if (sourceCell.s) {
      targetCell.s = { ...sourceCell.s };
    }
    if (!mappedColumns.has(column) && sourceCell.f) {
      targetCell.f = translateFormula(
        sourceCell.f,
        sourceRowIndex + 1,
        targetRowIndex + 1,
      );
      targetCell.t = sourceCell.t;
    }
    worksheet[targetAddress] = targetCell;
  }
}

function appendLogSheet(workbook, changes) {
  if (!changes.length) {
    return;
  }

  let worksheet = workbook.Sheets[LOG_SHEET_NAME];
  let rows = [];
  if (worksheet) {
    rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  } else {
    rows = [["Timestamp", "Source PDF", "Target Sheet", "Cell", "Field", "Value"]];
    workbook.SheetNames.push(LOG_SHEET_NAME);
  }

  const timestamp = new Date().toISOString().slice(0, 19);
  changes.forEach((change) => {
    rows.push([
      timestamp,
      change.source_pdf || "",
      change.sheet,
      change.cell,
      FIELD_LABELS[change.field] || change.field,
      change.value,
    ]);
  });
  workbook.Sheets[LOG_SHEET_NAME] = XLSX.utils.aoa_to_sheet(rows);
}

function rowHasMappedValues(worksheet, rowIndex, mappedColumns) {
  for (const column of mappedColumns) {
    const address = XLSX.utils.encode_cell({
      r: rowIndex,
      c: columnToIndex(column),
    });
    const cell = worksheet[address];
    if (cell?.v !== undefined && cell.v !== "") {
      return true;
    }
  }
  return false;
}

function existingTransactionKeys(worksheet, sheet, mapping) {
  const keys = new Set();
  const range = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : null;
  if (!range) {
    return keys;
  }

  for (let rowIndex = sheet.headerRowIndex + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const payload = {};
    ["transaction_date", "description", "debit", "credit", "amount"].forEach((field) => {
      const column = mapping[field];
      if (!column) {
        return;
      }
      const address = XLSX.utils.encode_cell({
        r: rowIndex,
        c: columnToIndex(column),
      });
      payload[field] = worksheet[address]?.v ?? "";
    });
    const key = transactionKeyFromObject(payload);
    if (key) {
      keys.add(key);
    }
  }
  return keys;
}

function transactionKeyFromObject(transaction) {
  const date = normalizeDate(transaction.transaction_date) || cleanCell(transaction.transaction_date);
  const description = cleanCell(transaction.description).toLowerCase();
  const debit = amountKey(transaction.debit);
  const credit = amountKey(transaction.credit);
  const amount = amountKey(transaction.amount);
  if (!date && !description && !debit && !credit && !amount) {
    return "";
  }
  return [date, description, debit, credit, amount].join("|");
}

function transactionFieldValue(transaction, field) {
  if (field === "notes") {
    return transaction.notes || transaction.status || "";
  }
  return transaction[field] ?? "";
}

function cellForValue(field, value) {
  if (["debit", "credit", "amount", "balance"].includes(field)) {
    const numeric = Number(value);
    if (!Number.isNaN(numeric)) {
      return { t: "n", v: numeric };
    }
  }
  return { t: "s", v: String(value) };
}

function highlightedStyle(style = {}) {
  return {
    ...style,
    fill: {
      patternType: "solid",
      fgColor: { rgb: "FFF2CC" },
    },
  };
}

function translateFormula(formula, sourceRow, targetRow) {
  const delta = targetRow - sourceRow;
  return String(formula).replace(
    /(\$?[A-Z]{1,3})(\$?)(\d+)/g,
    (match, column, absoluteRow, rowText) => {
      if (absoluteRow) {
        return match;
      }
      return `${column}${Number(rowText) + delta}`;
    },
  );
}

function renderWorkbookControls() {
  const sheetNames = state.workbook?.SheetNames || [];
  elements.workbookPickText.textContent = state.workbookFile
    ? "Change Excel file"
    : "Select Excel file";
  elements.workbookName.textContent = state.workbookFile?.name || "None";
  elements.workbookSheetCount.textContent = sheetNames.length
    ? String(sheetNames.length)
    : "-";
  elements.targetSheet.innerHTML = "";
  sheetNames.forEach((sheetName) => {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    elements.targetSheet.append(option);
  });
  elements.targetSheet.disabled = !sheetNames.length;
  if (state.targetSheetName) {
    elements.targetSheet.value = state.targetSheetName;
  }
}

function renderMappingPreview(issues = []) {
  elements.mappingPreview.innerHTML = "";
  const sheet = state.workbookSheets[state.targetSheetName];
  if (!sheet) {
    elements.mappingPreview.textContent = "";
    return;
  }

  if (!Object.keys(state.mapping).length) {
    state.mapping = suggestMapping(sheet);
  }

  const title = document.createElement("div");
  title.className = "mapping-title";
  title.textContent = `Header row ${sheet.headerRow}, next row ${sheet.nextRow}`;
  elements.mappingPreview.append(title);

  const fields = document.createElement("div");
  fields.className = "mapping-grid";
  const columns = Object.entries(sheet.headers);
  APPEND_FIELDS.forEach((field) => {
    const label = document.createElement("label");
    const text = document.createElement("span");
    const select = document.createElement("select");

    text.textContent = FIELD_LABELS[field];
    select.dataset.field = field;

    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "Do not write";
    select.append(empty);

    columns.forEach(([column, header]) => {
      const option = document.createElement("option");
      option.value = column;
      option.textContent = `${column} - ${header}`;
      select.append(option);
    });

    select.value = state.mapping[field] || "";
    select.addEventListener("change", () => {
      state.mapping[field] = select.value;
      state.appendedWorkbook = null;
      state.appendSummary = null;
      updateControls();
    });

    label.append(text, select);
    fields.append(label);
  });
  elements.mappingPreview.append(fields);

  const warnings = issues.filter((issue) => issue.severity === "warning");
  if (warnings.length) {
    const warning = document.createElement("div");
    warning.className = "mapping-warning";
    warning.textContent = `${warnings.length} warning${
      warnings.length === 1 ? "" : "s"
    } found.`;
    elements.mappingPreview.append(warning);
  }
}

function renderSheets() {
  elements.sheetList.innerHTML = "";
  setBadge(
    elements.sheetBadge,
    state.appendSummary
      ? String(state.appendSummary.rowsAdded)
      : String(state.sheets.length),
    state.appendSummary || state.sheets.length ? "good" : "",
  );

  if (state.appendSummary) {
    renderAppendSummary();
    return;
  }

  if (!state.sheets.length) {
    elements.summary.textContent = "No workbook yet.";
    return;
  }

  const rowCount = state.sheets.reduce(
    (total, sheet) => total + sheet.rows.length,
    0,
  );
  const transactionText = state.transactions.length
    ? ` ${state.transactions.length} transaction row${
        state.transactions.length === 1 ? "" : "s"
      } detected.`
    : "";
  elements.summary.textContent = `${state.sheets.length} sheet${
    state.sheets.length === 1 ? "" : "s"
  } ready, ${rowCount} row${rowCount === 1 ? "" : "s"} total.${transactionText}`;

  for (const sheet of state.sheets) {
    elements.sheetList.append(sheetPreviewItem(sheet));
  }
}

function renderAppendSummary() {
  const summary = state.appendSummary;
  const validationWarnings = summary.validation.filter(
    (issue) => issue.severity === "warning",
  ).length;
  elements.summary.textContent = `${summary.rowsAdded} row${
    summary.rowsAdded === 1 ? "" : "s"
  } added to ${state.targetSheetName}; ${summary.rowsSkipped} skipped${
    validationWarnings ? `; ${validationWarnings} warning(s)` : ""
  }.`;

  const transactions = {
    title: "Extracted transactions",
    source: "append",
    rows: [
      ["Date", "Description", "Debit", "Credit", "Amount"],
      ...state.transactions.slice(0, 12).map((transaction) => [
        transaction.transaction_date,
        transaction.description,
        transaction.debit ?? "",
        transaction.credit ?? "",
        transaction.amount ?? "",
      ]),
    ],
  };
  elements.sheetList.append(sheetPreviewItem(transactions));

  const changes = {
    title: "Changed cells",
    source: "log",
    rows: [
      ["Sheet", "Cell", "Field", "Value"],
      ...summary.changes.slice(0, 12).map((change) => [
        change.sheet,
        change.cell,
        FIELD_LABELS[change.field] || change.field,
        change.value,
      ]),
    ],
  };
  elements.sheetList.append(sheetPreviewItem(changes));
}

function sheetPreviewItem(sheet) {
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
  return item;
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
      cell.textContent = row[columnIndex] ?? "";
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
  resetPdfState();
  resetWorkbookState();
  state.busy = false;
  setStatus("Choose a PDF to begin.");
  setProgress(0);
  clearPreview();
  renderSheets();
  updateControls();
}

function resetPdfState() {
  state.file = null;
  state.files = [];
  state.pdf = null;
  state.sheets = [];
  state.transactions = [];
  state.rawText = "";
  state.appendedWorkbook = null;
  state.appendedFilename = "";
  state.appendSummary = null;
  elements.fileInput.value = "";
  elements.fileName.textContent = "None";
  elements.fileSize.textContent = "-";
  elements.pageCount.textContent = "-";
  setBadge(elements.fileBadge, "No file");
  setBadge(elements.sheetBadge, "0");
  setBadge(elements.previewBadge, "Waiting");
}

function resetWorkbookState() {
  state.workbookFile = null;
  state.workbook = null;
  state.workbookSheets = {};
  state.targetSheetName = "";
  state.mapping = {};
  state.appendedWorkbook = null;
  state.appendedFilename = "";
  state.appendSummary = null;
  elements.workbookInput.value = "";
  renderWorkbookControls();
  renderMappingPreview();
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
  elements.appendButton.disabled =
    !state.file || !state.workbookFile || !state.targetSheetName || state.busy;
  elements.downloadButton.disabled =
    (!state.sheets.length && !state.appendedWorkbook) || state.busy;
  elements.resetButton.disabled =
    (!state.file &&
      !state.sheets.length &&
      !state.workbookFile &&
      !state.appendedWorkbook) ||
    state.busy;
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

function withUniqueSheetTitles(sheets) {
  const existing = new Set();
  return sheets.map((sheet) => {
    const prefix =
      state.files.length > 1 && sheet.fileName
        ? `${sheet.fileName.replace(/\.pdf$/i, "")} Page`
        : "Page";
    const title = uniqueSheetTitle(
      existing,
      `${prefix} ${sheet.pageNumber} ${capitalize(sheet.source)} 1`,
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

function defaultTargetSheet(sheetNames) {
  return sheetNames.includes("BoA") ? "BoA" : sheetNames[0] || "";
}

function isPdfFile(file) {
  return file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
}

function isWorkbookFile(file) {
  return /\.(xlsx|xlsm)$/i.test(file.name);
}

function outputFilename(name) {
  const base = name.replace(/\.pdf$/i, "").trim() || "converted";
  return `${base}.xlsx`;
}

function updatedWorkbookFilename(name) {
  const match = name.match(/^(.*?)(\.(xlsx|xlsm))$/i);
  if (!match) {
    return "updated_workbook.xlsx";
  }
  return `${match[1]}_updated${match[2]}`;
}

function parseAmount(value) {
  if (value === null || value === undefined) {
    return null;
  }
  let text = cleanCell(value);
  if (!text || text === "-" || text === "--") {
    return null;
  }
  const negative = amountTextIsNegative(text);
  text = text
    .replace(/\b(CR|DR)\b/gi, "")
    .replace(/[$,()]/g, "")
    .trim();
  const amount = Number(text);
  if (Number.isNaN(amount)) {
    return null;
  }
  return negative ? -Math.abs(amount) : amount;
}

function normalizeDate(value) {
  const text = cleanCell(value);
  if (!text) {
    return "";
  }

  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    return `${iso[1]}-${pad2(iso[2])}-${pad2(iso[3])}`;
  }

  const slash = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slash) {
    const year = slash[3].length === 2 ? `20${slash[3]}` : slash[3];
    return `${year}-${pad2(slash[1])}-${pad2(slash[2])}`;
  }

  const parsed = new Date(text.replace(/\./g, ""));
  if (!Number.isNaN(parsed.valueOf())) {
    return parsed.toISOString().slice(0, 10);
  }
  return "";
}

function amountTextIsNegative(text) {
  const stripped = cleanCell(text);
  return stripped.startsWith("-") || (stripped.startsWith("(") && stripped.endsWith(")"));
}

function debitLikeLine(upperLine) {
  return [
    " DR",
    "DEBIT",
    "WITHDRAWAL",
    "PURCHASE",
    "PAYMENT TO",
    "FEE",
    "ATM",
    "POS",
    "CHECKCARD",
    "MOBILE PURCHASE",
    "ZELLE PAYMENT TO",
  ].some((token) => upperLine.includes(token));
}

function isBalanceOnlyDescription(description) {
  return /^ending bal|^beginning bal|^opening bal|^closing bal/i.test(description);
}

function normalizeBoaSection(line) {
  return cleanCell(line).replace(/\s+-\s+continued$/i, "").toLowerCase();
}

function isBoaNoiseLine(line) {
  const lower = line.toLowerCase();
  return (
    lower.startsWith("page ") ||
    lower.startsWith("available in ") ||
    lower === "continued on the next page" ||
    lower.startsWith("make bank transfers") ||
    lower.startsWith("use our app") ||
    lower.startsWith("scan the code") ||
    lower.startsWith("bankofamerica.com")
  );
}

function extractReferenceNumber(description) {
  const conf = description.match(/\bConf#\s*([A-Z0-9]+)/i);
  if (conf) {
    return conf[1];
  }
  const confirmation = description.match(/\b(?:ID|Confirmation#)\s*:?\s*([A-Z0-9]+)/i);
  if (confirmation) {
    return confirmation[1];
  }
  const longNumber = description.match(/\b(\d{10,})\b/);
  return longNumber ? longNumber[1] : "";
}

function deduplicateTransactions(transactions) {
  const seen = new Set();
  const unique = [];
  transactions.forEach((transaction) => {
    const key = transactionKeyFromObject(transaction);
    if (!key || seen.has(key)) {
      return;
    }
    seen.add(key);
    unique.push(transaction);
  });
  return unique;
}

function amountKey(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  const amount = Number(value);
  return Number.isNaN(amount) ? cleanCell(value) : amount.toFixed(2);
}

function indexToColumn(index) {
  let number = index + 1;
  let column = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    number = Math.floor((number - 1) / 26);
  }
  return column;
}

function columnToIndex(column) {
  return cleanCell(column)
    .toUpperCase()
    .split("")
    .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function normalizeLabel(value) {
  return cleanCell(value).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
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

function round2(value) {
  return Math.round(value * 100) / 100;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function capitalize(value) {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

renderWorkbookControls();
renderMappingPreview();
updateControls();
