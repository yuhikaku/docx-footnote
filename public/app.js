const DOCX_EXTENSION = ".docx";

const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const footnotesInput = document.getElementById("footnotes-input");
const runButton = document.getElementById("run-button");
const statusEl = document.getElementById("status");
const downloadLink = document.getElementById("download");

let selectedFile = null;
let downloadUrl = null;

/**
 * Show a status message. kind is "", "error" or "warn".
 */
const setStatus = (message, kind = "") => {
  statusEl.textContent = message;
  statusEl.className = kind;
};

const clearDownload = () => {
  if (downloadUrl) {
    URL.revokeObjectURL(downloadUrl);
    downloadUrl = null;
  }
  downloadLink.style.display = "none";
};

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
 * Build the output file name: input.docx -> input_footnotes.docx
 */
const buildOutputName = (inputName) => {
  const base = inputName.slice(0, -DOCX_EXTENSION.length);
  return `${base}_footnotes${DOCX_EXTENSION}`;
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
      const detail = await readErrorDetail(response);
      setStatus(`Error ${response.status}: ${detail}`, "error");
      return;
    }

    const blob = await response.blob();
    downloadUrl = URL.createObjectURL(blob);
    downloadLink.href = downloadUrl;
    downloadLink.download = buildOutputName(selectedFile.name);
    downloadLink.style.display = "inline";

    const count = response.headers.get("X-Footnote-Count");
    const missing = response.headers.get("X-Missing-Numbers");
    if (missing) {
      setStatus(
        `Done. ${count} footnote(s) inserted.\nMissing numbers: ${missing}`,
        "warn",
      );
      return;
    }
    setStatus(`Done. ${count} footnote(s) inserted.`);
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
