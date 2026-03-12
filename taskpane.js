/**
 * BC Inquiry Checker – Outlook Add-in
 * Tototheo Maritime Ltd
 *
 * Flow:
 *  1. Office.onReady → read email subject + body
 *  2. Extract RFQ reference with regex
 *  3. User clicks Search → query BC OData API
 *  4. Show results with link to open Inquiry List filtered by Customer Ref
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * CONFIGURATION – update these values for your environment
 * ─────────────────────────────────────────────────────────────────────────────
 */
const BC_CONFIG = {
  // Your BC tenant URL (from the URL you provided)
  baseUrl: "https://businesscentral.dynamics.com/Tototheo/",

  // Company name (URL-encoded)
  company: "Tototheo Maritime Ltd",

  // Inquiry List page number
  inquiryListPage: 70355879,

  // OData API endpoint for your Inquiry table
  // Adjust the entity set name to match your BC setup
  // Common patterns: "InquiryHeaders", "SalesInquiries", or your custom API name
  // Example: /api/v2.0/companies(...)/inquiryHeaders
  // If you have a custom API page, replace below:
  odataEntity: "InquiryHeaders", // ← CHANGE to your actual OData entity name

  // The OData field name for Customer Ref (the field users type the RFQ number into)
  customerRefField: "customerReference", // ← CHANGE to match your BC field name

  // Fields to display in results
  displayFields: ["no", "customerReference", "shipName", "customerName", "status", "assignedTo", "documentDate"],
};

/**
 * Regex patterns to find RFQ numbers in email subject/body.
 * Add more patterns if your customers use different formats.
 * Matches: RFQ 53700205 | RFQ: 53700205 | RFQ#53700205 | Ref: 53700205
 */
const RFQ_PATTERNS = [
  /\bRFQ[:\s#-]*(\d{5,12})\b/i,
  /\bRef(?:erence)?[:\s#-]*(\d{5,12})\b/i,
  /\bRequest\s+for\s+Quotation[:\s#-]*(\d{5,12})\b/i,
  /\bOrder\s+Ref[:\s#-]*(\d{5,12})\b/i,
  /\bPO[:\s#-]*(\d{5,12})\b/i,
];

// ─────────────────────────────────────────────────────────────────────────────
// STATE MANAGEMENT
// ─────────────────────────────────────────────────────────────────────────────

let currentRef = "";

function showState(name) {
  const states = ["loading", "detected", "no-detect", "searching", "found", "notfound"];
  states.forEach(s => {
    const el = document.getElementById(`state-${s}`);
    if (el) el.style.display = (s === name) ? "" : "none";
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// OFFICE.JS INIT
// ─────────────────────────────────────────────────────────────────────────────

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initAddin();
  }
});

function initAddin() {
  attachEventListeners();
  extractRfqFromEmail();
}

// ─────────────────────────────────────────────────────────────────────────────
// EXTRACT RFQ FROM EMAIL
// ─────────────────────────────────────────────────────────────────────────────

function extractRfqFromEmail() {
  showState("loading");

  const item = Office.context.mailbox.item;

  // 1. Check subject first (fastest)
  const subject = item.subject || "";
  const fromSubject = findRfqInText(subject);

  if (fromSubject) {
    setDetectedState(fromSubject, "from email subject");
    return;
  }

  // 2. Check body (async)
  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const bodyText = result.value || "";
      const fromBody = findRfqInText(bodyText);
      if (fromBody) {
        setDetectedState(fromBody, "from email body");
      } else {
        showState("no-detect");
      }
    } else {
      // Body read failed – show manual input
      showState("no-detect");
    }
  });
}

function findRfqInText(text) {
  for (const pattern of RFQ_PATTERNS) {
    const match = text.match(pattern);
    if (match) return match[1];
  }
  return null;
}

function setDetectedState(ref, source) {
  currentRef = ref;
  document.getElementById("detected-ref-value").textContent = ref;
  document.getElementById("detected-ref-source").textContent = source;
  showState("detected");
}

// ─────────────────────────────────────────────────────────────────────────────
// EVENT LISTENERS
// ─────────────────────────────────────────────────────────────────────────────

function attachEventListeners() {
  // Detected state: search with detected ref
  document.getElementById("btn-search-detected").addEventListener("click", function () {
    searchInquiry(currentRef);
  });

  // Detected state: manual override
  document.getElementById("btn-search-manual").addEventListener("click", function () {
    const val = document.getElementById("manual-input").value.trim();
    if (val) searchInquiry(val);
  });
  document.getElementById("manual-input").addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
      const val = this.value.trim();
      if (val) searchInquiry(val);
    }
  });

  // No-detect state: manual search
  document.getElementById("btn-search-manual-2").addEventListener("click", function () {
    const val = document.getElementById("manual-input-2").value.trim();
    if (val) searchInquiry(val);
  });
  document.getElementById("manual-input-2").addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
      const val = this.value.trim();
      if (val) searchInquiry(val);
    }
  });

  // No-detect: open list directly
  document.getElementById("btn-open-list-direct").addEventListener("click", function () {
    openInquiryList(null);
  });

  // Not found: open list
  document.getElementById("btn-open-list-notfound").addEventListener("click", function () {
    openInquiryList(currentRef);
  });

  // Search again buttons
  document.getElementById("btn-search-again-found").addEventListener("click", resetToSearch);
  document.getElementById("btn-search-again-notfound").addEventListener("click", resetToSearch);
}

function resetToSearch() {
  if (currentRef) {
    showState("detected");
  } else {
    showState("no-detect");
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SEARCH BUSINESS CENTRAL via OData
// ─────────────────────────────────────────────────────────────────────────────

async function searchInquiry(ref) {
  currentRef = ref;
  document.getElementById("searching-ref-display").textContent = `Customer Ref: ${ref}`;
  showState("searching");

  try {
    const results = await queryBCOData(ref);

    if (results && results.length > 0) {
      renderFoundState(results, ref);
    } else {
      renderNotFoundState(ref);
    }
  } catch (err) {
    console.error("BC OData error:", err);
    // On error, fall back to opening BC directly
    renderNotFoundState(ref, true);
  }
}

/**
 * Query Business Central OData API
 *
 * Uses the standard BC OData v4 endpoint.
 * Authentication is handled automatically via the user's existing
 * Microsoft 365 / Entra ID session (same account as Outlook).
 *
 * OData URL format:
 * /ODataV4/Company('Tototheo Maritime Ltd')/InquiryHeaders?$filter=customerReference eq '53700205'
 *
 * NOTE: If your Inquiry table is a custom page published as API,
 * the URL pattern will be different – update BC_CONFIG.odataEntity accordingly.
 */
async function queryBCOData(ref) {
  const companyEncoded = encodeURIComponent(BC_CONFIG.company);
  const refEncoded = ref.replace(/'/g, "''"); // OData single-quote escaping

  // Try OData v4 endpoint first
  const odataUrl = `https://api.businesscentral.dynamics.com/v2.0/Tototheo/sandbox/ODataV4/Company('${companyEncoded}')/${BC_CONFIG.odataEntity}?$filter=${BC_CONFIG.customerRefField} eq '${refEncoded}'&$top=10`;

  const response = await fetch(odataUrl, {
    method: "GET",
    headers: {
      "Accept": "application/json",
      // Auth handled automatically by the browser session / MSAL
    },
    credentials: "include",
  });

  if (!response.ok) {
    // If OData fails (e.g. 401, 404), throw so we fall back to manual open
    throw new Error(`OData error: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  return data.value || [];
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER STATES
// ─────────────────────────────────────────────────────────────────────────────

function renderFoundState(results, ref) {
  const container = document.getElementById("results-container");
  container.innerHTML = "";

  const countLabel = document.getElementById("found-count-label");
  countLabel.textContent = results.length === 1
    ? "1 Inquiry found in BC"
    : `${results.length} Inquiries found in BC`;

  results.forEach(function (item) {
    const statusClass = getStatusClass(item.status || item.Status || "");
    const card = document.createElement("div");
    card.className = "inquiry-card";
    card.innerHTML = `
      <div class="inquiry-header">
        <div class="inquiry-id">${escHtml(item.no || item.No || "—")}</div>
        <span class="badge ${statusClass}">
          <span class="badge-dot"></span>
          ${escHtml(item.status || item.Status || "—")}
        </span>
      </div>
      <div class="inquiry-field"><strong>Customer Ref:</strong> <span class="ref-value">${escHtml(item.customerReference || item[BC_CONFIG.customerRefField] || ref)}</span></div>
      ${item.shipName || item.ShipName ? `<div class="inquiry-field"><strong>Vessel:</strong> ${escHtml(item.shipName || item.ShipName)}</div>` : ""}
      ${item.customerName || item.CustomerName ? `<div class="inquiry-field"><strong>Customer:</strong> ${escHtml(item.customerName || item.CustomerName)}</div>` : ""}
      ${item.assignedTo || item.AssignedTo ? `<div class="inquiry-field"><strong>Assignee:</strong> ${escHtml(item.assignedTo || item.AssignedTo)}</div>` : ""}
      <button
        class="btn btn-primary"
        style="width:100%;margin-top:10px;font-size:11px;padding:7px"
        onclick="openInquiryListItem('${escHtml(item.no || item.No || "")}', '${escHtml(ref)}')"
      >
        Open in Inquiry List →
      </button>
    `;
    container.appendChild(card);
  });

  showState("found");
}

function renderNotFoundState(ref, isError) {
  document.getElementById("notfound-ref-display").textContent = isError
    ? `Ref: ${ref} (could not reach BC – open manually)`
    : `Ref: ${ref} — not found`;
  showState("notfound");
}

// ─────────────────────────────────────────────────────────────────────────────
// OPEN BUSINESS CENTRAL PAGES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Opens the Inquiry List page in BC, filtered by Customer Ref.
 * Uses the exact URL pattern from the link you provided.
 */
function openInquiryList(ref) {
  const companyParam = encodeURIComponent(BC_CONFIG.company);
  let url = `${BC_CONFIG.baseUrl}?company=${companyParam}&page=${BC_CONFIG.inquiryListPage}`;

  // If we have a ref, add a filter – BC supports URL filter params
  if (ref) {
    // BC filter syntax: &filter=CustomerReference IS '53700205'
    // OR use the bookmark from the BC URL directly
    url += `&filter=${encodeURIComponent(BC_CONFIG.customerRefField + " IS '" + ref + "'")}`;
  }

  Office.context.ui.openBrowserWindow(url);
}

/**
 * Opens the Inquiry List filtered to show a specific inquiry by No.
 */
function openInquiryListItem(no, ref) {
  const companyParam = encodeURIComponent(BC_CONFIG.company);
  let url = `${BC_CONFIG.baseUrl}?company=${companyParam}&page=${BC_CONFIG.inquiryListPage}`;

  if (ref) {
    url += `&filter=${encodeURIComponent(BC_CONFIG.customerRefField + " IS '" + ref + "'")}`;
  }

  Office.context.ui.openBrowserWindow(url);
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function getStatusClass(status) {
  const s = (status || "").toLowerCase();
  if (s.includes("open"))   return "badge-open";
  if (s.includes("quot"))   return "badge-quoted";
  if (s.includes("clos"))   return "badge-closed";
  return "badge-draft";
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
