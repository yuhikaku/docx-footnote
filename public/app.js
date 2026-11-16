const DOCX_EXTENSION = ".docx";
const OUTPUT_SUFFIX = "_脚注付";
const MISSING_SUFFIX = "_欠番";
const MISSING_FALLBACK_TEXT = "（Excel中になし）";
const ERROR_MESSAGES = {
  403: "このネットワーク（IPアドレス）からは利用できません。許可されたネットワークからアクセスしてください。",
  413: "ファイルが大きすぎます。分割してお試しください。",
  429: "アクセスが集中しています。しばらく待ってからお試しください。",
};
const ACCESS_DENIED_STATUS = 403;
const EXCEL_EXTENSIONS = [".xlsx", ".xls"];

const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const footnotesInput = document.getElementById("footnotes-input");
const runButton = document.getElementById("run-button");
const statusEl = document.getElementById("status");
const downloadLink = document.getElementById("download");
const missingLink = document.getElementById("download-missing");
const accessBanner = document.getElementById("access-banner");
const accessBannerMessage = document.getElementById("access-banner-message");
const tipsCallout = document.getElementById("tips-callout");
const excelDropZone = document.getElementById("excel-drop-zone");
const excelInput = document.getElementById("excel-input");

let selectedFile = null;
let downloadUrl = null;
let missingUrl = null;

/**
 * Show or hide the access denied banner.
 */
const showAccessBanner = (message) => {
  accessBannerMessage.textContent = message;
  accessBanner.hidden = false;
};

const hideAccessBanner = () => {
  accessBanner.hidden = true;
};

/**
 * Show or hide the tips callout.
 */
const showTips = () => {
  tipsCallout.hidden = false;
};

const hideTips = () => {
  tipsCallout.hidden = true;
};

/**
 * Show a status message. kind is "", "error" or "warn".
 */
const setStatus = (message, kind = "") => {
  statusEl.textContent = message;
  statusEl.className = kind;
};

const clearDownload = () => {
  hideAccessBanner();
  hideTips();
  if (downloadUrl) {
    URL.revokeObjectURL(downloadUrl);
    downloadUrl = null;
  }
  if (missingUrl) {
    URL.revokeObjectURL(missingUrl);
    missingUrl = null;
  }
  downloadLink.style.display = "none";
  missingLink.style.display = "none";
};

/**
 * Convert a cell value to TSV text.
 * Cells containing tab, newline or quote are wrapped in double quotes.
 */
const toTsvCell = (value) => {
  const text = String(value ?? "");
  if (!/[\t\r\n"]/.test(text)) {
    return text;
  }
  return `"${text.replace(/"/g, '""')}"`;
};

/**
 * Convert the first sheet of a workbook to TSV text.
 */
const sheetToTsv = (workbook) => {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    blankrows: false,
    defval: "",
  });
  return rows
    .map((row) => row.slice(0, 2).map(toTsvCell).join("\t"))
    .join("\n");
};

/**
 * Read the first sheet of an Excel file into the footnotes textarea.
 */
const loadExcel = async (file) => {
  const isExcel = EXCEL_EXTENSIONS.some((ext) =>
    file.name.toLowerCase().endsWith(ext),
  );
  if (!isExcel) {
    setStatus("Only .xlsx / .xls files are supported.", "error");
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    footnotesInput.value = sheetToTsv(workbook);
    excelDropZone.textContent = file.name;
    excelDropZone.classList.add("has-file");
    clearDownload();
    setStatus("");
    updateRunButton();
  } catch (error) {
    setStatus(`Failed to read Excel: ${error.message}`, "error");
  }
};

excelDropZone.addEventListener("click", () => excelInput.click());
excelDropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    excelInput.click();
  }
});
excelInput.addEventListener("change", () => {
  const file = excelInput.files[0];
  if (file) {
    loadExcel(file);
  }
  excelInput.value = "";
});
excelDropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  excelDropZone.classList.add("dragover");
});
excelDropZone.addEventListener("dragleave", () => {
  excelDropZone.classList.remove("dragover");
});
excelDropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  excelDropZone.classList.remove("dragover");
  const file = event.dataTransfer.files[0];
  if (file) {
    loadExcel(file);
  }
});

const updateRunButton = () => {
  runButton.disabled = !(selectedFile && footnotesInput.value.trim());
};

const setFile = (file) => {
  if (!file.name.toLowerCase().endsWith(DOCX_EXTENSION)) {
    setStatus("Only .docx files are supported.", "error");
    return;
  }
  selectedFile = file;
  dropZone.textContent = file.name;
  dropZone.classList.add("has-file");
  clearDownload();
  setStatus("");
  updateRunButton();
};

/**
 * Parse Excel clipboard text (TSV with quoted cells) into rows of cells.
 * A cell wrapped in double quotes may contain tabs and newlines;
 * a doubled quote ("") inside it is an escaped quote.
 */
const parseTsv = (raw) => {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  let index = 0;

  const endCell = () => {
    row.push(cell);
    cell = "";
  };
  const endRow = () => {
    endCell();
    rows.push(row);
    row = [];
  };

  while (index < raw.length) {
    const char = raw[index];

    if (inQuotes) {
      if (char !== '"') {
        cell += char;
        index += 1;
        continue;
      }
      if (raw[index + 1] === '"') {
        cell += '"';
        index += 2;
        continue;
      }
      inQuotes = false;
      index += 1;
      continue;
    }

    if (char === '"' && cell === "") {
      inQuotes = true;
      index += 1;
      continue;
    }
    if (char === "\t") {
      endCell();
      index += 1;
      continue;
    }
    if (char === "\r" && raw[index + 1] === "\n") {
      endRow();
      index += 2;
      continue;
    }
    if (char === "\n" || char === "\r") {
      endRow();
      index += 1;
      continue;
    }
    cell += char;
    index += 1;
  }

  if (cell !== "" || row.length !== 0) {
    endRow();
  }
  return rows;
};

/**
 * Collapse line breaks inside a cell into a single space.
 */
const normalizeCell = (cell) => cell.replace(/\s*[\r\n]+\s*/g, " ").trim();

/**
 * Parse pasted text into [[number, text], ...].
 * Throws an Error with a row number when a row is invalid.
 */
const parseFootnotes = (raw) => {
  const pairs = [];

  for (const [index, columns] of parseTsv(raw).entries()) {
    const isEmptyRow = columns.every((cell) => !cell.trim());
    if (isEmptyRow) {
      continue;
    }

    const number = normalizeCell(columns[0] ?? "");
    const text = normalizeCell(columns.slice(1).join("\t"));

    if (!/^[0-9]+$/.test(number)) {
      throw new Error(`Row ${index + 1}: the first column must be a number.`);
    }
    if (!text) {
      throw new Error(`Row ${index + 1}: the second column is empty.`);
    }
    pairs.push([number, text]);
  }

  if (pairs.length === 0) {
    throw new Error("No footnotes found.");
  }
  return pairs;
};

/**
 * Build the output file name: input.docx -> input_脚注付.docx
 */
const buildOutputName = (inputName) => {
  const base = inputName.slice(0, -DOCX_EXTENSION.length);
  return `${base}${OUTPUT_SUFFIX}${DOCX_EXTENSION}`;
};

/**
 * Build the missing report name: input.docx -> input_脚注付_欠番.txt
 */
const buildMissingName = (inputName) => {
  const base = inputName.slice(0, -DOCX_EXTENSION.length);
  return `${base}${OUTPUT_SUFFIX}${MISSING_SUFFIX}.txt`;
};

/**
 * Build the missing numbers report text.
 * Same format as the CLI: "■number\ntext\n" joined by newlines.
 */
const buildMissingReport = (missingNumbers, pairs) => {
  const textByNumber = new Map(pairs);
  return missingNumbers
    .map(
      (num) => `■${num}\n${textByNumber.get(num) ?? MISSING_FALLBACK_TEXT}\n`,
    )
    .join("\n");
};

/**
 * Show the missing report download link.
 */
const showMissingReport = (missingNumbers, pairs) => {
  const report = buildMissingReport(missingNumbers, pairs);
  const blob = new Blob([report], { type: "text/plain;charset=utf-8" });
  missingUrl = URL.createObjectURL(blob);
  missingLink.href = missingUrl;
  missingLink.download = buildMissingName(selectedFile.name);
  missingLink.style.display = "inline";
};

/**
 * Read the error detail from a failed response.
 */
const readErrorDetail = async (response) => {
  try {
    const body = await response.json();
    if (typeof body.detail === "string") {
      return body.detail;
    }
  } catch {
    // fall through to the status text
  }
  return response.statusText || "Unknown error";
};

const run = async () => {
  clearDownload();

  let pairs;
  try {
    pairs = parseFootnotes(footnotesInput.value);
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }

  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("footnotes", JSON.stringify(pairs));

  runButton.disabled = true;
  setStatus("Converting...");

  try {
    const response = await fetch("/api/footnotes", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      if (response.status === ACCESS_DENIED_STATUS) {
        showAccessBanner(ERROR_MESSAGES[ACCESS_DENIED_STATUS]);
        return;
      }
      const known = ERROR_MESSAGES[response.status];
      const detail = known ?? (await readErrorDetail(response));
      setStatus(`Error ${response.status}: ${detail}`, "error");
      return;
    }

    const blob = await response.blob();
    downloadUrl = URL.createObjectURL(blob);
    downloadLink.href = downloadUrl;
    downloadLink.download = buildOutputName(selectedFile.name);
    downloadLink.style.display = "inline";
    showTips();

    const count = response.headers.get("X-Footnote-Count");
    const missingHeader = response.headers.get("X-Missing-Numbers");
    if (missingHeader) {
      const missingNumbers = missingHeader.split(",");
      showMissingReport(missingNumbers, pairs);
      setStatus(
        `${count}件の脚注を追加しました。\n欠番${missingNumbers.length}件（${missingHeader}）`,
        "warn",
      );
      return;
    }
    setStatus(`${count}件の脚注を追加しました。`);
  } catch (error) {
    setStatus(`Request failed: ${error.message}`, "error");
  } finally {
    updateRunButton();
  }
};

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (file) {
    setFile(file);
  }
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragover");
});
dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});
dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragover");
  const file = event.dataTransfer.files[0];
  if (file) {
    setFile(file);
  }
});

footnotesInput.addEventListener("input", updateRunButton);
runButton.addEventListener("click", run);
