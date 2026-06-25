"use strict";

const TEMPLATE_HEADERS = [
  "ID",
  "External ID",
  "Name",
  "Customer Type",
  "Address1",
  "Address2",
  "Address3",
  "Address4",
  "Post Code",
  "Customer Telephone",
  "Tags",
  "Account Number",
  "VAT Number",
  "Reference Number",
  "Customer Job Type",
  "Customer Job Category",
  "Enable Billing Address",
  "Billing Address Name",
  "Billing Address 1",
  "Billing Address 2",
  "Billing Address 3",
  "Billing Address 4",
  "Billing Postcode",
  "Billing Telephone",
  "Billing Email Address",
  "Billing Account Number",
  "Billing other Emails",
  "Warning 1",
  "Warning 1 Show",
  "Warning 2",
  "Warning 2 Show",
  "Warning 3",
  "Warning 3 Show",
  "Invoice Payment Due",
  "Invoice Payment Terms",
  "Account Manager",
  "Job Order Number",
  "Selling Rate",
  "Contact 1 First Name",
  "Contact 1 Last Name",
  "Contact 1 Telephone",
  "Contact 1 Secondary Telephone",
  "Contact 1 Email Address",
  "Contact 1 Position",
  "Contact 2 First Name",
  "Contact 2 Last Name",
  "Contact 2 Telephone",
  "Contact 2 Secondary Telephone",
  "Contact 2 Email Address",
  "Contact 2 Position",
  "Contact 3 First Name",
  "Contact 3 Last Name",
  "Contact 3 Telephone",
  "Contact 3 Secondary Telephone",
  "Contact 3 Email Address",
  "Contact 3 Position",
  "Contact 4 First Name",
  "Contact 4 Last Name",
  "Contact 4 Telephone",
  "Contact 4 Secondary Telephone",
  "Contact 4 Email Address",
  "Contact 4 Position",
  "Contact 5 First Name",
  "Contact 5 Last Name",
  "Contact 5 Telephone",
  "Contact 5 Secondary Telephone",
  "Contact 5 Email Address",
  "Contact 5 Position",
  "Auto-Generate Sites",
  "Copy Data from Customer to Site",
  "Notes",
  "Note Types",
  "Latitude",
  "Longitude",
  "Suspend / Inactivate Customer",
  "Line Type",
  "Nullable Column / Field Value delete",
];

const TEMPLATE_INSTRUCTIONS = [
  "* 7 characters max\n* Required (Update)\n* Alphanumeric (A-Z, a-z, 0-9)",
  "* 255 Characters max Mapped With EDI Reference\n* (Only Use in Insert Case)",
  "* 255 characters max\n* Required",
  "* If a value is entered it must exist in 'Customer Type' library",
  "* 255 characters max",
  "* 255 characters max",
  "* 255 characters max",
  "* 255 characters max",
  "",
  "* 60 characters max",
  "* Should exist in 'Tag(s)' list - separate tags with comma",
  "",
  "* 32 characters max",
  "* 255 characters max",
  "* Separate multiple with comma",
  "* Separate multiple with comma",
  "* true/false",
  "* 255 characters max\n* Required if Enable billing address is true",
  "* 255 characters max\n* Required if Enable billing address is true",
  "* 255 characters max",
  "* 255 characters max",
  "* 255 characters max",
  "* 50 characters max\n* Required if Enable billing address is true",
  "* 60 characters max",
  "* 255 characters max",
  "* 255 characters max",
  "",
  "* 8000 characters max",
  "* true/false",
  "* 8000 characters max",
  "* true/false",
  "* 8000 characters max",
  "* true/false",
  "* In days\n* Number must be greater than 0 and less than 438,000",
  "* 1000 characters max",
  "* If a value is entered it must exist in 'User'",
  "* Is Mandatory (yes/no)",
  "",
  "* 64 characters max",
  "* 64 characters max",
  "",
  "",
  "* 255 characters max\n* Must be a valid system format",
  "* 64 characters max",
  "* 64 characters max",
  "* 64 characters max",
  "",
  "",
  "* 255 characters max\n* Must be a valid system format",
  "* 64 characters max",
  "* 64 characters max",
  "* 64 characters max",
  "",
  "",
  "* 255 characters max\n* Must be a valid system format",
  "* 64 characters max",
  "* 64 characters max",
  "* 64 characters max",
  "",
  "",
  "* 255 characters max\n* Must be a valid system format",
  "* 64 characters max",
  "* 64 characters max",
  "* 64 characters max",
  "",
  "",
  "* 255 characters max\n* Must be a valid system format",
  "* 64 characters max",
  "*Yes/No\n*Default/Blank would be No",
  "* Yes/No\n* Default/Blank will be No\n* Data to be copied: Telephone, All Contacts, Reference Number, and Complete Billing Information.",
  "* Add '<Note_Separator>' in between for more than one notes",
  "1: Public\n2: Private\n3: Private and show on mobile\n* Separate multiple with commas.",
  "* Insert and Update",
  "* Insert and Update",
  "*Yes/No\n*Default/Blank would be No",
  '* Required\n* Value must be "Insert" or "Update"\n* "Insert" to import new line\n* "Update" to modify existing data (Customer ID must exist)\n* Update action should only update the cell values which are provided for update and shouldn\'t change non-populated fields\n* "Delete" to soft delete the line.',
  "Update against the Customer Ids provided",
];

const ADDRESS_HEADERS = ["Address1", "Address2", "Address3", "Address4", "Post Code"];
const PREVIEW_HEADERS = [
  "Name",
  "Customer Type",
  "Address1",
  "Address2",
  "Address3",
  "Address4",
  "Post Code",
  "Invoice Payment Due",
  "Notes",
  "Note Types",
  "Tags",
];

const HEADER_ALIASES = new Map([["ablecolumnfieldvaluedelete", "nullablecolumnfieldvaluedelete"]]);

const state = {
  cleanedRows: [],
  typeCounts: new Map(),
  stats: {
    rows: 0,
    nameFixes: 0,
    addressFixes: 0,
    tagFixes: 0,
    paymentFixes: 0,
    noteFixes: 0,
  },
  outputFileName: "cleaned_customers.xlsx",
};

const elements = {};

document.addEventListener("DOMContentLoaded", () => {
  bindElements();
  bindEvents();
  routeView();
  refreshIcons();
});

function bindElements() {
  elements.pageTitle = document.querySelector("#pageTitle");
  elements.homeButton = document.querySelector("#homeButton");
  elements.homeView = document.querySelector("#homeView");
  elements.customerView = document.querySelector("#customerView");
  elements.moduleCards = document.querySelectorAll(".module-card");
  elements.fileInput = document.querySelector("#fileInput");
  elements.dropZone = document.querySelector("#dropZone");
  elements.downloadButton = document.querySelector("#downloadButton");
  elements.fileName = document.querySelector("#fileName");
  elements.statusBadge = document.querySelector("#statusBadge");
  elements.rowCount = document.querySelector("#rowCount");
  elements.nameFixes = document.querySelector("#nameFixes");
  elements.addressFixes = document.querySelector("#addressFixes");
  elements.tagFixes = document.querySelector("#tagFixes");
  elements.paymentFixes = document.querySelector("#paymentFixes");
  elements.noteFixes = document.querySelector("#noteFixes");
  elements.typeCount = document.querySelector("#typeCount");
  elements.typeList = document.querySelector("#typeList");
  elements.previewCount = document.querySelector("#previewCount");
  elements.previewBody = document.querySelector("#previewBody");
}

function bindEvents() {
  elements.homeButton.addEventListener("click", () => {
    window.location.hash = "";
  });

  elements.moduleCards.forEach((card) => {
    card.addEventListener("click", () => {
      if (card.dataset.module === "customer") {
        window.location.hash = "customer";
      }
    });
  });

  window.addEventListener("hashchange", routeView);

  elements.fileInput.addEventListener("change", (event) => {
    const [file] = event.target.files;
    if (file) {
      processFile(file);
    }
  });

  elements.dropZone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      elements.fileInput.click();
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
      elements.fileInput.files = event.dataTransfer.files;
      processFile(file);
    }
  });

  elements.downloadButton.addEventListener("click", () => {
    downloadCleanedWorkbook();
  });
}

function routeView() {
  const isCustomer = window.location.hash.replace("#", "").toLowerCase() === "customer";
  elements.homeView.classList.toggle("view-hidden", isCustomer);
  elements.customerView.classList.toggle("view-hidden", !isCustomer);
  elements.pageTitle.textContent = isCustomer ? "Customer Cleaner" : "Import Cleaner";
  elements.homeButton.hidden = !isCustomer;
  elements.statusBadge.hidden = !isCustomer;
  if (isCustomer) {
    setStatus(state.cleanedRows.length ? "Cleaned" : "Ready", state.cleanedRows.length ? "good" : "");
  }
  refreshIcons();
}

async function processFile(file) {
  try {
    setStatus("Reading", "");
    elements.fileName.textContent = file.name;
    elements.downloadButton.disabled = true;

    await waitForXlsx();
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellDates: false,
      raw: false,
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false,
    });

    const result = cleanWorkbookRows(rows);
    Object.assign(state, result);
    state.outputFileName = makeOutputName(file.name);

    renderResults();
    setStatus("Cleaned", "good");
    elements.downloadButton.disabled = state.cleanedRows.length === 0;
  } catch (error) {
    console.error(error);
    resetResults();
    setStatus(error.message || "Could not clean file", "bad");
  }
}

function waitForXlsx() {
  if (window.XLSX) return Promise.resolve();

  return new Promise((resolve, reject) => {
    const startedAt = Date.now();
    const timer = window.setInterval(() => {
      if (window.XLSX) {
        window.clearInterval(timer);
        resolve();
      } else if (Date.now() - startedAt > 8000) {
        window.clearInterval(timer);
        reject(new Error("Spreadsheet library unavailable"));
      }
    }, 80);
  });
}

function cleanWorkbookRows(rows) {
  const headerIndex = findHeaderIndex(rows);
  if (headerIndex < 0) {
    throw new Error("Missing customer headers");
  }

  const headers = rows[headerIndex].map((value) => toText(value));
  const headerMap = buildHeaderMap(headers);
  const dataRows = rows.slice(headerIndex + 1).filter((row, index) => {
    if (index === 0 && looksLikeInstructionRow(row)) return false;
    return row.some((cell) => toText(cell).trim() !== "");
  });

  const typeCounts = new Map();
  const stats = {
    rows: 0,
    nameFixes: 0,
    addressFixes: 0,
    tagFixes: 0,
    paymentFixes: 0,
    noteFixes: 0,
  };
  const cleanedRows = dataRows.map((row) => {
    const sourceName = getSourceValue(row, headerMap, "Name");
    const cleanNameValue = normalizeSpaces(sourceName);
    const sourceTags = getSourceValue(row, headerMap, "Tags");
    const cleanTagsValue = cleanTags(sourceTags);
    const sourceCustomerType = getSourceValue(row, headerMap, "Customer Type");
    const cleanCustomerType = normalizeSpaces(sourceCustomerType);
    const cleanAddress = cleanAddressFields(row, headerMap);
    const sourcePaymentDue = getSourceValue(row, headerMap, "Invoice Payment Due");
    const cleanPaymentDue = cleanInvoicePaymentDue(sourcePaymentDue);
    const noteResult = cleanNotesAndTypes(
      getSourceValue(row, headerMap, "Notes"),
      getSourceValue(row, headerMap, "Note Types"),
    );

    const output = TEMPLATE_HEADERS.map((header) => {
      if (header === "ID") return "";
      if (header === "External ID") {
        return preserveCell(
          getSourceValue(row, headerMap, "External ID") || getSourceValue(row, headerMap, "ID"),
        );
      }
      if (header === "Name") return cleanNameValue;
      if (header === "Customer Type") return cleanCustomerType;
      if (header === "Tags") return cleanTagsValue;
      if (header in cleanAddress) return cleanAddress[header];
      if (header === "Invoice Payment Due") return cleanPaymentDue;
      if (header === "Notes") return noteResult.notes;
      if (header === "Note Types") return noteResult.noteTypes;
      return preserveCell(getSourceValue(row, headerMap, header));
    });

    stats.rows += 1;
    if (toText(sourceName) !== cleanNameValue) stats.nameFixes += 1;
    if (toText(sourceTags) !== cleanTagsValue) stats.tagFixes += 1;
    if (addressChanged(row, headerMap, cleanAddress)) stats.addressFixes += 1;
    if (toText(sourcePaymentDue) !== toText(cleanPaymentDue)) stats.paymentFixes += 1;
    if (noteResult.changed) stats.noteFixes += 1;
    if (cleanCustomerType) {
      typeCounts.set(cleanCustomerType, (typeCounts.get(cleanCustomerType) || 0) + 1);
    }

    return output;
  });

  return { cleanedRows, typeCounts, stats };
}

function findHeaderIndex(rows) {
  return rows.findIndex((row) => {
    const names = new Set(row.map((cell) => normalizeHeader(cell)));
    return names.has("name") && names.has("customertype");
  });
}

function buildHeaderMap(headers) {
  const map = new Map();
  headers.forEach((header, index) => {
    const key = normalizeHeader(header);
    if (key && !map.has(key)) {
      map.set(key, index);
    }
  });
  return map;
}

function normalizeHeader(value) {
  const key = toText(value).toLowerCase().replace(/[^a-z0-9]/g, "");
  return HEADER_ALIASES.get(key) || key;
}

function looksLikeInstructionRow(row) {
  const populated = row.map((cell) => toText(cell).trim()).filter(Boolean);
  if (!populated.length) return false;
  const instructionHits = populated.filter((cell) =>
    /(^\*|characters max|required|default\/blank|insert and update)/i.test(cell),
  );
  return instructionHits.length >= Math.max(2, Math.ceil(populated.length * 0.25));
}

function getSourceValue(row, headerMap, header) {
  const index = headerMap.get(normalizeHeader(header));
  return index === undefined ? "" : row[index];
}

function preserveCell(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" && value.trim() === "") return "";
  return value;
}

function normalizeSpaces(value) {
  return toText(value).replace(/\s+/g, " ").trim();
}

function toText(value) {
  if (value === null || value === undefined) return "";
  return String(value);
}

function cleanTags(value) {
  const seen = new Set();
  const tags = [];

  toText(value)
    .split(",")
    .map((part) => normalizeSpaces(part))
    .filter(Boolean)
    .forEach((tag) => {
      const key = tag.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        tags.push(tag);
      }
    });

  return tags.join(", ");
}

function cleanInvoicePaymentDue(value) {
  return isZeroOnly(value) ? "" : preserveCell(value);
}

function cleanAddressFields(row, headerMap) {
  const fragments = [];
  let postcode = cleanAddressValue(getSourceValue(row, headerMap, "Post Code"));

  ["Address1", "Address2", "Address3", "Address4"].forEach((header) => {
    const prepared = cleanAddressValue(getSourceValue(row, headerMap, header));
    if (!prepared) return;

    splitAddressParts(prepared).forEach((part) => {
      const { text, code } = extractPostcode(part);
      if (!postcode && code) {
        postcode = code;
      }
      if (text) {
        fragments.push(text);
      }
    });
  });

  const compact = dedupeFragments(fragments).slice(0, 4);
  return {
    Address1: compact[0] || "",
    Address2: compact[1] || "",
    Address3: compact[2] || "",
    Address4: compact[3] || "",
    "Post Code": normalizePostcode(postcode),
  };
}

function cleanAddressValue(value) {
  const text = cleanAddressText(value);
  return isSingleNumber(text) ? "" : text;
}

function cleanAddressText(value) {
  let text = toText(value)
    .replace(/\u00a0/g, " ")
    .replace(/[\r\n]+/g, ", ")
    .replace(/\s+/g, " ")
    .replace(/\s+,/g, ",")
    .replace(/,\s*/g, ", ")
    .replace(/(?:,\s*){2,}/g, ", ")
    .replace(/^[,\s]+|[,\s]+$/g, "")
    .trim();

  const emailMatch = text.match(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i);
  const bracketedCode = text.match(/\(([^()]*\d[^()]*)\)/);
  if (emailMatch && bracketedCode) {
    text = bracketedCode[1];
  } else if (emailMatch) {
    text = text.replace(emailMatch[0], "");
  }

  return normalizeSpaces(text).replace(/^[,\s]+|[,\s]+$/g, "");
}

function splitAddressParts(value) {
  return toText(value)
    .split(",")
    .map((part) => cleanAddressValue(part))
    .filter(Boolean);
}

function extractPostcode(value) {
  const text = toText(value);
  const ukPostcode =
    /\b(GIR\s?0AA|[A-PR-UWYZ][A-HK-Y]?\d[A-Z\d]?\s*\d[ABD-HJLNP-UW-Z]{2})\b/i;
  const match = text.match(ukPostcode);
  if (!match) return { text: cleanAddressText(text), code: "" };

  const withoutCode = text.replace(match[0], "");
  return {
    text: cleanAddressText(withoutCode),
    code: normalizePostcode(match[0]),
  };
}

function normalizePostcode(value) {
  return cleanAddressValue(value).toUpperCase();
}

function dedupeFragments(fragments) {
  const seen = new Set();
  const clean = [];

  fragments.forEach((fragment) => {
    const value = cleanAddressValue(fragment);
    const key = value.toLowerCase();
    if (value && !seen.has(key)) {
      seen.add(key);
      clean.push(value);
    }
  });

  return clean;
}

function addressChanged(row, headerMap, cleanAddress) {
  return ADDRESS_HEADERS.some((header) => {
    const source = cleanAddressValue(getSourceValue(row, headerMap, header));
    const cleaned = cleanAddressValue(cleanAddress[header]);
    return source !== cleaned;
  });
}

function cleanNotesAndTypes(notesValue, noteTypesValue) {
  const notes = splitNotes(notesValue);
  const noteCount = notes.length;
  const sourceNoteTypes = splitNoteTypes(noteTypesValue);
  let cleanTypes = [];

  if (noteCount > 0) {
    if (sourceNoteTypes.length === 0) {
      cleanTypes = Array(noteCount).fill("2");
    } else {
      cleanTypes = sourceNoteTypes.slice(0, noteCount);
      while (cleanTypes.length < noteCount) {
        cleanTypes.push(cleanTypes[cleanTypes.length - 1] || "2");
      }
    }
  }

  const cleanNotes = notes.join(" <Note_Separator> ");
  const cleanNoteTypes = cleanTypes.join(",");
  return {
    notes: cleanNotes,
    noteTypes: cleanNoteTypes,
    changed:
      toText(notesValue) !== cleanNotes ||
      toText(noteTypesValue) !== cleanNoteTypes ||
      sourceNoteTypes.length !== noteCount,
  };
}

function splitNotes(value) {
  return toText(value)
    .split(/<\s*Note_Separator\s*>/i)
    .map((note) => normalizeSpaces(note))
    .filter(Boolean);
}

function splitNoteTypes(value) {
  return toText(value)
    .split(",")
    .map((type) => normalizeSpaces(type))
    .filter(Boolean);
}

function isSingleNumber(value) {
  return /^\d+$/.test(normalizeSpaces(value));
}

function isZeroOnly(value) {
  return /^0+(?:\.0+)?$/.test(normalizeSpaces(value));
}

function renderResults() {
  elements.rowCount.textContent = state.stats.rows.toLocaleString();
  elements.nameFixes.textContent = state.stats.nameFixes.toLocaleString();
  elements.addressFixes.textContent = state.stats.addressFixes.toLocaleString();
  elements.tagFixes.textContent = state.stats.tagFixes.toLocaleString();
  elements.paymentFixes.textContent = state.stats.paymentFixes.toLocaleString();
  elements.noteFixes.textContent = state.stats.noteFixes.toLocaleString();
  renderTypes();
  renderPreview();
  refreshIcons();
}

function renderTypes() {
  const sorted = [...state.typeCounts.entries()].sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });

  elements.typeCount.textContent = String(sorted.length);
  elements.typeList.classList.toggle("empty-state", sorted.length === 0);
  elements.typeList.innerHTML = "";

  if (!sorted.length) {
    elements.typeList.textContent = "None";
    return;
  }

  sorted.forEach(([name, count]) => {
    const item = document.createElement("div");
    item.className = "type-item";
    item.innerHTML = `<span>${escapeHtml(name)}</span><span>${count.toLocaleString()}</span>`;
    elements.typeList.appendChild(item);
  });
}

function renderPreview() {
  const previewRows = state.cleanedRows.slice(0, 30);
  elements.previewCount.textContent = `${Math.min(state.cleanedRows.length, 30).toLocaleString()} rows`;
  elements.previewBody.innerHTML = "";

  if (!previewRows.length) {
    elements.previewBody.innerHTML =
      '<tr><td colspan="11" class="empty-state">Waiting for workbook</td></tr>';
    return;
  }

  const indices = PREVIEW_HEADERS.map((header) => TEMPLATE_HEADERS.indexOf(header));
  previewRows.forEach((row) => {
    const tr = document.createElement("tr");
    indices.forEach((index) => {
      const td = document.createElement("td");
      td.textContent = toText(row[index]);
      tr.appendChild(td);
    });
    elements.previewBody.appendChild(tr);
  });
}

function resetResults() {
  state.cleanedRows = [];
  state.typeCounts = new Map();
  state.stats = {
    rows: 0,
    nameFixes: 0,
    addressFixes: 0,
    tagFixes: 0,
    paymentFixes: 0,
    noteFixes: 0,
  };
  elements.downloadButton.disabled = true;
  renderResults();
}

function setStatus(text, tone) {
  elements.statusBadge.textContent = text;
  elements.statusBadge.className = `status-badge ${tone || ""}`.trim();
}

function downloadCleanedWorkbook() {
  if (!state.cleanedRows.length) return;

  const data = [TEMPLATE_HEADERS, TEMPLATE_INSTRUCTIONS, ...state.cleanedRows];
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  worksheet["!cols"] = TEMPLATE_HEADERS.map((header) => ({
    wch: Math.max(12, Math.min(34, header.length + 5)),
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Customers");
  XLSX.writeFile(workbook, state.outputFileName, { compression: true });
}

function makeOutputName(fileName) {
  const base = fileName.replace(/\.[^.]+$/, "").trim() || "customers";
  return `${base}_cleaned_customers.xlsx`;
}

function escapeHtml(value) {
  return toText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function refreshIcons() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

window.CustomerCleaner = {
  cleanWorkbookRows,
  cleanAddressFields,
  cleanTags,
  cleanInvoicePaymentDue,
  cleanNotesAndTypes,
  normalizeSpaces,
};
