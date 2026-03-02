const expenseRows = document.getElementById("expenseRows");
const rowTemplate = document.getElementById("rowTemplate");

const addRowBtn = document.getElementById("addRowBtn");
const printBtn = document.getElementById("printBtn");
const saveJsonBtn = document.getElementById("saveJsonBtn");
const loadJsonBtn = document.getElementById("loadJsonBtn");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const jsonFileInput = document.getElementById("jsonFileInput");

const currencySelect = document.getElementById("currency");
const mileageRateInput = document.getElementById("mileageRate");
const baseTotalEl = document.getElementById("baseTotal");
const mileageTotalEl = document.getElementById("mileageTotal");
const grandTotalEl = document.getElementById("grandTotal");
const categoryBreakdownEl = document.getElementById("categoryBreakdown");
const printReport = document.getElementById("printReport");
const rowReceiptData = new WeakMap();
const rowScanState = new WeakMap();
let pdfWorkerConfigured = false;

const headerFieldIds = [
  "employeeName",
  "department",
  "manager",
  "purpose",
  "startDate",
  "endDate",
  "currency",
  "mileageRate"
];

function getCurrencyFormatter() {
  const currency = currencySelect.value || "USD";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency,
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function parseNumber(value) {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatReportDate(value) {
  if (!value) {
    return "—";
  }

  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return date.toLocaleDateString("en-US", {
    year: "numeric",
    month: "short",
    day: "2-digit"
  });
}

function formatReportTimestamp(value = new Date()) {
  return value.toLocaleString("en-US", {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "numeric",
    minute: "2-digit"
  });
}

function getHeaderValues() {
  const headers = {};
  headerFieldIds.forEach((id) => {
    headers[id] = document.getElementById(id).value.trim();
  });
  return headers;
}

function getMeaningfulRows() {
  return readRows().filter(
    (item) => item.date || item.description || item.amount > 0 || item.miles > 0 || item.receipt
  );
}

function rowElements(row) {
  return {
    date: row.querySelector(".row-date"),
    category: row.querySelector(".row-category"),
    description: row.querySelector(".row-description"),
    amount: row.querySelector(".row-amount"),
    miles: row.querySelector(".row-miles"),
    mileageTotal: row.querySelector(".row-mileage-total"),
    total: row.querySelector(".row-total"),
    receiptFile: row.querySelector(".row-receipt-file"),
    receiptUpload: row.querySelector(".row-receipt-upload"),
    receiptView: row.querySelector(".row-receipt-view"),
    receiptScan: row.querySelector(".row-receipt-scan"),
    receiptName: row.querySelector(".row-receipt-name"),
    receiptStatus: row.querySelector(".row-receipt-status"),
    remove: row.querySelector(".row-remove")
  };
}

function setReceiptStatus(row, text, state = "muted") {
  const els = rowElements(row);
  els.receiptStatus.textContent = text;
  els.receiptStatus.dataset.state = state;
}

function setScanBusy(row, isBusy) {
  const els = rowElements(row);
  const hasReceipt = Boolean(rowReceiptData.get(row));

  els.receiptUpload.disabled = isBusy;
  els.receiptView.disabled = isBusy || !hasReceipt;
  els.receiptScan.disabled = isBusy || !hasReceipt;
  els.receiptScan.textContent = isBusy ? "Scanning..." : "Scan";
  rowScanState.set(row, isBusy);
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;
  return `${value.toFixed(value >= 100 || index === 0 ? 0 : 1)} ${units[index]}`;
}

function updateReceiptUi(row) {
  const els = rowElements(row);
  const data = rowReceiptData.get(row);

  if (!data || !data.dataUrl) {
    els.receiptName.textContent = "No file";
    setScanBusy(row, false);
    els.receiptView.disabled = true;
    els.receiptScan.disabled = true;
    setReceiptStatus(row, "Attach a file to scan.", "muted");
    return;
  }

  els.receiptName.textContent = `${data.name} (${formatFileSize(data.size)})`;
  setScanBusy(row, false);
  setReceiptStatus(row, "Ready to scan.", "muted");
}

function setReceiptData(row, receipt) {
  if (receipt && typeof receipt === "object" && receipt.dataUrl) {
    rowReceiptData.set(row, {
      name: receipt.name || "receipt",
      type: receipt.type || "application/octet-stream",
      size: parseNumber(receipt.size),
      dataUrl: receipt.dataUrl
    });
  } else {
    rowReceiptData.delete(row);
  }

  updateReceiptUi(row);
}

function readReceiptFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      resolve({
        name: file.name,
        type: file.type || "application/octet-stream",
        size: file.size,
        dataUrl: String(reader.result || "")
      });
    };
    reader.onerror = () => reject(new Error("Failed to read the selected file."));
    reader.readAsDataURL(file);
  });
}

function dataUrlToBytes(dataUrl) {
  const base64 = String(dataUrl).split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes;
}

function getLibrariesOrThrow(receiptType = "") {
  if (!window.Tesseract || typeof window.Tesseract.recognize !== "function") {
    throw new Error("OCR library not available.");
  }

  const isPdf = receiptType === "application/pdf";
  if (isPdf && !window.pdfjsLib) {
    throw new Error("PDF library not available.");
  }
}

function configurePdfWorker() {
  if (pdfWorkerConfigured || !window.pdfjsLib) {
    return;
  }

  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  pdfWorkerConfigured = true;
}

async function recognizeImageText(source) {
  const result = await window.Tesseract.recognize(source, "eng");
  return result?.data?.text || "";
}

async function extractTextFromPdfDataUrl(dataUrl) {
  configurePdfWorker();
  const bytes = dataUrlToBytes(dataUrl);
  const pdf = await window.pdfjsLib.getDocument({ data: bytes }).promise;
  const pageCount = Math.min(pdf.numPages || 1, 2);
  const pageTexts = [];

  for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", { alpha: false });

    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;

    const text = await recognizeImageText(canvas);
    pageTexts.push(text);

    canvas.width = 0;
    canvas.height = 0;
  }

  return pageTexts.join("\n");
}

function receiptIsPdf(receipt) {
  return (
    receipt?.type === "application/pdf" ||
    String(receipt?.name || "").toLowerCase().endsWith(".pdf") ||
    String(receipt?.dataUrl || "").startsWith("data:application/pdf")
  );
}

function receiptIsImage(receipt) {
  return (
    String(receipt?.type || "").startsWith("image/") ||
    String(receipt?.dataUrl || "").startsWith("data:image/")
  );
}

async function renderPdfReceiptPages(receipt) {
  configurePdfWorker();
  const bytes = dataUrlToBytes(receipt.dataUrl);
  const pdf = await window.pdfjsLib.getDocument({ data: bytes }).promise;
  const pageImages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", { alpha: false });

    if (!context) {
      throw new Error("Could not create canvas context for PDF rendering.");
    }

    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;

    pageImages.push(canvas.toDataURL("image/png"));
    canvas.width = 0;
    canvas.height = 0;
  }

  return pageImages;
}

function extractAmountValues(text) {
  const matches = String(text).match(/(?:[$€£]\s*)?(-?\d{1,3}(?:,\d{3})*(?:\.\d{2})|-?\d+\.\d{2})/g) || [];

  return matches
    .map((match) => Number.parseFloat(match.replace(/[$€£,\s]/g, "")))
    .filter((amount) => Number.isFinite(amount) && amount > 0 && amount < 100000);
}

function detectAmountFromText(text) {
  const lines = String(text)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const keywordRules = [
    { keyword: "grand total", score: 110 },
    { keyword: "total due", score: 100 },
    { keyword: "amount due", score: 95 },
    { keyword: "balance due", score: 95 },
    { keyword: "amount paid", score: 90 },
    { keyword: "total", score: 80 },
    { keyword: "paid", score: 70 }
  ];
  const deductionRules = [
    { keyword: "subtotal", penalty: 40 },
    { keyword: "tax", penalty: 25 },
    { keyword: "tip", penalty: 20 },
    { keyword: "discount", penalty: 25 },
    { keyword: "change", penalty: 30 }
  ];
  const candidates = [];

  lines.forEach((line) => {
    const values = extractAmountValues(line);
    if (!values.length) {
      return;
    }

    const lower = line.toLowerCase();
    let score = 0;
    keywordRules.forEach((rule) => {
      if (lower.includes(rule.keyword)) {
        score = Math.max(score, rule.score);
      }
    });
    deductionRules.forEach((rule) => {
      if (lower.includes(rule.keyword)) {
        score -= rule.penalty;
      }
    });

    if (score > 0) {
      candidates.push({
        amount: Math.max(...values),
        score,
        line
      });
    }
  });

  if (candidates.length) {
    candidates.sort((left, right) => right.score - left.score || right.amount - left.amount);
    const best = candidates[0];
    return {
      amount: best.amount,
      confidence: best.score >= 95 ? "high" : "medium",
      sourceLine: best.line
    };
  }

  const fallbackValues = extractAmountValues(text);
  if (!fallbackValues.length) {
    return null;
  }

  return {
    amount: Math.max(...fallbackValues),
    confidence: "low",
    sourceLine: ""
  };
}

async function extractReceiptText(receipt) {
  const isPdf = receiptIsPdf(receipt);
  getLibrariesOrThrow(isPdf ? "application/pdf" : "image");

  if (isPdf) {
    return extractTextFromPdfDataUrl(receipt.dataUrl);
  }

  return recognizeImageText(receipt.dataUrl);
}

async function scanReceiptForAmount(row, auto = false) {
  if (rowScanState.get(row)) {
    return;
  }

  const receipt = rowReceiptData.get(row);
  if (!receipt) {
    setReceiptStatus(row, "Attach a file to scan.", "warning");
    return;
  }

  setScanBusy(row, true);
  setReceiptStatus(row, auto ? "Auto-scanning receipt..." : "Scanning receipt...", "muted");

  try {
    const text = await extractReceiptText(receipt);
    const detection = detectAmountFromText(text);

    if (!detection) {
      setReceiptStatus(row, "No total found. Enter amount manually.", "warning");
      return;
    }

    const els = rowElements(row);
    els.amount.value = detection.amount.toFixed(2);
    updateSummary();

    const amountLabel = getCurrencyFormatter().format(detection.amount);
    const confidenceLabel =
      detection.confidence === "high" ? "high confidence" : detection.confidence === "medium" ? "medium confidence" : "low confidence";
    setReceiptStatus(row, `Detected ${amountLabel} (${confidenceLabel}).`, detection.confidence === "low" ? "warning" : "success");
  } catch (error) {
    setReceiptStatus(row, "Scan failed. Try another file or enter manually.", "error");
  } finally {
    setScanBusy(row, false);
  }
}

function addExpenseRow(initialData = null) {
  const fragment = rowTemplate.content.cloneNode(true);
  const row = fragment.querySelector("tr");
  const els = rowElements(row);

  if (initialData) {
    els.date.value = initialData.date || "";
    els.category.value = initialData.category || "Other";
    els.description.value = initialData.description || "";
    els.amount.value = parseNumber(initialData.amount).toFixed(2);
    els.miles.value = parseNumber(initialData.miles).toFixed(1);
  }

  row.addEventListener("input", updateSummary);
  row.addEventListener("change", updateSummary);

  els.receiptUpload.addEventListener("click", () => {
    els.receiptFile.click();
  });

  els.receiptFile.addEventListener("change", async (event) => {
    const [file] = event.target.files || [];
    if (!file) {
      return;
    }

    try {
      const receipt = await readReceiptFile(file);
      setReceiptData(row, receipt);
      await scanReceiptForAmount(row, true);
    } catch (error) {
      window.alert("Could not process this receipt file.");
      setReceiptData(row, null);
    } finally {
      event.target.value = "";
    }
  });

  els.receiptView.addEventListener("click", () => {
    const receipt = rowReceiptData.get(row);
    if (!receipt || !receipt.dataUrl) {
      return;
    }

    const tab = window.open(receipt.dataUrl, "_blank", "noopener,noreferrer");
    if (!tab) {
      window.alert("Please allow popups to preview the receipt.");
    }
  });

  els.receiptScan.addEventListener("click", async () => {
    await scanReceiptForAmount(row, false);
  });

  els.remove.addEventListener("click", () => {
    rowReceiptData.delete(row);
    row.remove();
    if (!expenseRows.querySelector("tr")) {
      addExpenseRow();
    }
    updateSummary();
  });

  expenseRows.appendChild(fragment);
  setReceiptData(row, initialData && typeof initialData.receipt === "object" ? initialData.receipt : null);
  updateSummary();
}

function readRows() {
  return Array.from(expenseRows.querySelectorAll("tr")).map((row) => {
    const els = rowElements(row);
    const amount = parseNumber(els.amount.value);
    const miles = parseNumber(els.miles.value);
    const mileageRate = parseNumber(mileageRateInput.value);
    const mileageValue = miles * mileageRate;
    const total = amount + mileageValue;

    return {
      row,
      date: els.date.value,
      category: els.category.value,
      description: els.description.value.trim(),
      amount,
      miles,
      mileageValue,
      total,
      receipt: rowReceiptData.get(row) || null
    };
  });
}

function updateSummary() {
  const rows = readRows();
  const formatter = getCurrencyFormatter();

  let baseTotal = 0;
  let mileageTotal = 0;
  const categoryTotals = {};

  rows.forEach((item) => {
    const { row, category, amount, mileageValue, total } = item;
    const els = rowElements(row);

    baseTotal += amount;
    mileageTotal += mileageValue;
    categoryTotals[category] = (categoryTotals[category] || 0) + total;

    els.mileageTotal.textContent = formatter.format(mileageValue);
    els.total.textContent = formatter.format(total);
  });

  const grandTotal = baseTotal + mileageTotal;

  baseTotalEl.textContent = formatter.format(baseTotal);
  mileageTotalEl.textContent = formatter.format(mileageTotal);
  grandTotalEl.textContent = formatter.format(grandTotal);

  renderCategoryBreakdown(categoryTotals, formatter);
}

function renderCategoryBreakdown(categoryTotals, formatter) {
  categoryBreakdownEl.innerHTML = "";

  const entries = Object.entries(categoryTotals)
    .sort((a, b) => b[1] - a[1])
    .filter(([, value]) => value > 0);

  if (!entries.length) {
    const li = document.createElement("li");
    li.className = "empty-breakdown";
    li.textContent = "No expenses entered yet.";
    categoryBreakdownEl.appendChild(li);
    return;
  }

  entries.forEach(([category, total]) => {
    const li = document.createElement("li");
    const label = document.createElement("span");
    const value = document.createElement("strong");

    label.textContent = category;
    value.textContent = formatter.format(total);

    li.appendChild(label);
    li.appendChild(value);
    categoryBreakdownEl.appendChild(li);
  });
}

async function buildReceiptAppendixHtml(receiptItems, formatter) {
  if (!receiptItems.length) {
    return `
      <h2 class="print-report__section-title">Receipt Appendix (Ordered by Line Item)</h2>
      <p class="print-report__muted">No receipts were attached to line items.</p>
    `;
  }

  let sections = "";

  for (const receiptItem of receiptItems) {
    const { lineNumber, item } = receiptItem;
    const receipt = item.receipt;
    const summary = [
      formatReportDate(item.date),
      item.category || "Uncategorized",
      formatter.format(item.total)
    ].join(" • ");

    let previewHtml = `<p class="print-report__muted">Preview unavailable for this file type.</p>`;

    try {
      if (receiptIsPdf(receipt)) {
        if (!window.pdfjsLib) {
          throw new Error("PDF library unavailable.");
        }

        const pages = await renderPdfReceiptPages(receipt);
        previewHtml = pages.length
          ? pages
              .map(
                (pageImage, pageIndex) => `
              <figure class="print-report__receipt-figure${pageIndex === 0 ? " print-report__receipt-figure--first" : ""}">
                <img
                  class="print-report__receipt-image"
                  src="${pageImage}"
                  alt="Receipt page ${pageIndex + 1} for line item ${lineNumber}"
                />
                <figcaption class="print-report__receipt-caption">Page ${pageIndex + 1}</figcaption>
              </figure>
            `
              )
              .join("")
          : `<p class="print-report__muted">PDF receipt has no renderable pages.</p>`;
      } else if (receiptIsImage(receipt)) {
        previewHtml = `
          <figure class="print-report__receipt-figure print-report__receipt-figure--first">
            <img
              class="print-report__receipt-image"
              src="${receipt.dataUrl}"
              alt="Receipt image for line item ${lineNumber}"
            />
          </figure>
        `;
      }
    } catch (error) {
      previewHtml = `<p class="print-report__muted">Unable to render preview for this receipt.</p>`;
    }

    sections += `
      <article class="print-report__receipt">
        <header class="print-report__receipt-head">
          <h3 class="print-report__receipt-title">Line Item #${lineNumber}</h3>
          <p class="print-report__receipt-meta">${escapeHtml(summary)}</p>
          <p class="print-report__receipt-meta">
            <strong>Receipt File:</strong> ${escapeHtml(receipt?.name || "Unknown file")}
          </p>
          <p class="print-report__receipt-meta">
            <strong>Description:</strong> ${escapeHtml(item.description || "—")}
          </p>
        </header>
        ${previewHtml}
      </article>
    `;
  }

  return `
    <h2 class="print-report__section-title print-report__section-break">Receipt Appendix (Ordered by Line Item)</h2>
    <p class="print-report__appendix-note">
      Receipts are attached below in the same sequence as the line-item numbers above.
    </p>
    ${sections}
  `;
}

async function buildPrintableReport() {
  if (!printReport) {
    return;
  }

  const headers = getHeaderValues();
  const rows = getMeaningfulRows();
  const formatter = getCurrencyFormatter();
  const receiptItems = [];

  let baseTotal = 0;
  let mileageTotal = 0;
  const categoryTotals = {};

  const rowHtml = rows.length
    ? rows
        .map((item, index) => {
          baseTotal += item.amount;
          mileageTotal += item.mileageValue;
          categoryTotals[item.category] = (categoryTotals[item.category] || 0) + item.total;
          if (item.receipt) {
            receiptItems.push({ lineNumber: index + 1, item });
          }

          return `
            <tr>
              <td>${index + 1}</td>
              <td>${escapeHtml(formatReportDate(item.date))}</td>
              <td>${escapeHtml(item.category || "—")}</td>
              <td class="print-report__desc">${escapeHtml(item.description || "—")}</td>
              <td class="numeric">${formatter.format(item.amount)}</td>
              <td class="numeric">${item.miles.toFixed(1)}</td>
              <td class="numeric">${formatter.format(item.mileageValue)}</td>
              <td class="numeric">${formatter.format(item.total)}</td>
              <td>${escapeHtml(item.receipt ? item.receipt.name : "None")}</td>
            </tr>
          `;
        })
        .join("")
    : `
      <tr>
        <td colspan="9" class="print-report__muted">No line items entered for this report.</td>
      </tr>
    `;

  const categoryRows = Object.entries(categoryTotals)
    .sort((left, right) => right[1] - left[1])
    .filter(([, total]) => total > 0);

  const categoryHtml = categoryRows.length
    ? `
      <table class="print-report__table">
        <thead>
          <tr>
            <th>Category</th>
            <th class="numeric">Total</th>
          </tr>
        </thead>
        <tbody>
          ${categoryRows
            .map(
              ([category, total]) => `
            <tr>
              <td>${escapeHtml(category)}</td>
              <td class="numeric">${formatter.format(total)}</td>
            </tr>
          `
            )
            .join("")}
        </tbody>
      </table>
    `
    : `<p class="print-report__muted">No non-zero category totals.</p>`;

  const grandTotal = baseTotal + mileageTotal;
  const periodText =
    headers.startDate || headers.endDate
      ? `${formatReportDate(headers.startDate)} to ${formatReportDate(headers.endDate)}`
      : "—";
  const receiptAppendixHtml = await buildReceiptAppendixHtml(receiptItems, formatter);

  printReport.innerHTML = `
    <header class="print-report__header">
      <div>
        <h1 class="print-report__title">Expense Reimbursement Report</h1>
        <p class="print-report__subtitle">Detailed and printable record for finance approval</p>
      </div>
      <p class="print-report__meta">Prepared: ${escapeHtml(formatReportTimestamp())}</p>
    </header>

    <section class="print-report__grid">
      <div class="print-report__field">
        <span class="print-report__field-label">Employee</span>
        <span>${escapeHtml(headers.employeeName || "—")}</span>
      </div>
      <div class="print-report__field">
        <span class="print-report__field-label">Department</span>
        <span>${escapeHtml(headers.department || "—")}</span>
      </div>
      <div class="print-report__field">
        <span class="print-report__field-label">Manager</span>
        <span>${escapeHtml(headers.manager || "—")}</span>
      </div>
      <div class="print-report__field">
        <span class="print-report__field-label">Report Period</span>
        <span>${escapeHtml(periodText)}</span>
      </div>
      <div class="print-report__field">
        <span class="print-report__field-label">Purpose</span>
        <span>${escapeHtml(headers.purpose || "—")}</span>
      </div>
      <div class="print-report__field">
        <span class="print-report__field-label">Currency / Rate</span>
        <span>${escapeHtml(headers.currency || "USD")} / ${escapeHtml(headers.mileageRate || "0")} per mile</span>
      </div>
    </section>

    <h2 class="print-report__section-title">Line Items</h2>
    <table class="print-report__table">
      <thead>
        <tr>
          <th>#</th>
          <th>Date</th>
          <th>Category</th>
          <th>Description</th>
          <th class="numeric">Amount</th>
          <th class="numeric">Miles</th>
          <th class="numeric">Mileage</th>
          <th class="numeric">Line Total</th>
          <th>Receipt</th>
        </tr>
      </thead>
      <tbody>${rowHtml}</tbody>
    </table>

    <table class="print-report__totals">
      <tbody>
        <tr>
          <td>Base Expenses</td>
          <td>${formatter.format(baseTotal)}</td>
        </tr>
        <tr>
          <td>Mileage Reimbursement</td>
          <td>${formatter.format(mileageTotal)}</td>
        </tr>
        <tr class="grand-total">
          <td>Grand Total</td>
          <td>${formatter.format(grandTotal)}</td>
        </tr>
      </tbody>
    </table>

    <h2 class="print-report__section-title">Category Summary</h2>
    ${categoryHtml}
    ${receiptAppendixHtml}

    <section class="print-report__compliance">
      I certify that these expenses were incurred for approved business purposes, are accurate to
      the best of my knowledge, and have supporting receipts attached where required by policy.
    </section>

    <section class="print-report__signatures">
      <div class="print-report__sig-line">Employee Signature / Date</div>
      <div class="print-report__sig-line">Manager Approval / Date</div>
    </section>
  `;
}

function collectReportData() {
  const headers = getHeaderValues();

  const rows = readRows().map((item) => ({
    date: item.date,
    category: item.category,
    description: item.description,
    amount: item.amount,
    miles: item.miles,
    receipt: item.receipt
  }));

  return {
    meta: {
      createdAt: new Date().toISOString()
    },
    headers,
    rows
  };
}

function applyReportData(data) {
  if (!data || typeof data !== "object") {
    return;
  }

  headerFieldIds.forEach((id) => {
    const input = document.getElementById(id);
    if (input && data.headers && Object.hasOwn(data.headers, id)) {
      input.value = data.headers[id];
    }
  });

  expenseRows.innerHTML = "";

  const rows = Array.isArray(data.rows) ? data.rows : [];
  if (!rows.length) {
    addExpenseRow();
    return;
  }

  rows.forEach((item) => addExpenseRow(item));
  updateSummary();
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function saveAsJson() {
  const data = collectReportData();
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  downloadBlob("expense-report.json", blob);
}

function loadFromJsonFile(file) {
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || "{}"));
      applyReportData(parsed);
    } catch (error) {
      window.alert("Could not read JSON file. Please verify its format.");
    }
  };
  reader.readAsText(file);
}

function escapeCsv(value) {
  const text = String(value ?? "");
  if (text.includes(",") || text.includes("\"") || text.includes("\n")) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function exportCsv() {
  const data = collectReportData();
  const rows = data.rows;

  const headerLines = [
    ["Employee Name", data.headers.employeeName],
    ["Department", data.headers.department],
    ["Manager", data.headers.manager],
    ["Purpose", data.headers.purpose],
    ["Start Date", data.headers.startDate],
    ["End Date", data.headers.endDate],
    ["Currency", data.headers.currency],
    ["Mileage Rate", data.headers.mileageRate],
    []
  ];

  const bodyLines = [["Date", "Category", "Description", "Amount", "Miles", "Receipt File"]];
  rows.forEach((row) => {
    bodyLines.push([
      row.date,
      row.category,
      row.description,
      row.amount.toFixed(2),
      row.miles.toFixed(1),
      row.receipt ? row.receipt.name : ""
    ]);
  });

  const lines = [...headerLines, ...bodyLines]
    .map((parts) => parts.map(escapeCsv).join(","))
    .join("\n");

  const blob = new Blob([lines], { type: "text/csv;charset=utf-8" });
  downloadBlob("expense-report.csv", blob);
}

async function printCompliantReport() {
  const previousLabel = printBtn.textContent;
  printBtn.disabled = true;
  printBtn.textContent = "Preparing...";

  try {
    await buildPrintableReport();
    window.print();
  } finally {
    printBtn.disabled = false;
    printBtn.textContent = previousLabel;
  }
}

addRowBtn.addEventListener("click", () => addExpenseRow());
printBtn.addEventListener("click", printCompliantReport);
saveJsonBtn.addEventListener("click", saveAsJson);
loadJsonBtn.addEventListener("click", () => jsonFileInput.click());
exportCsvBtn.addEventListener("click", exportCsv);
window.addEventListener("beforeprint", () => {
  void buildPrintableReport();
});

jsonFileInput.addEventListener("change", (event) => {
  const [file] = event.target.files || [];
  loadFromJsonFile(file);
  event.target.value = "";
});

currencySelect.addEventListener("change", updateSummary);
mileageRateInput.addEventListener("input", updateSummary);

headerFieldIds.forEach((id) => {
  const input = document.getElementById(id);
  if (input) {
    input.addEventListener("input", () => {
      if (id === "currency" || id === "mileageRate") {
        updateSummary();
      }
    });
  }
});

addExpenseRow();
updateSummary();
