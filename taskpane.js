/**
 * BC Inquiry Checker – Outlook Add-in
 * Tototheo Maritime Ltd
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * CONFIRMED FIELD NAMES (from Inquiry_List Excel export):
 *   No.                     → No
 *   Customer Name           → Customer_Name
 *   Vessel Name             → Vessel_Name
 *   Status                  → Status
 *   Sales Operator          → Sales_Operator
 *   Category Code           → Category_Code
 *   Creation Date           → Creation_Date
 *   Case Code               → Case_Code
 *
 * ⚠️  "Customer Ref." field NOT in export → field name to confirm below.
 *     Open this URL while logged in to M365 and look for the RFQ number field:
 *     https://api.businesscentral.dynamics.com/v2.0/e5296d4f-699b-43ea-a129-066a3f7010e3/Tototheo/ODataV4/Company('Tototheo%20Maritime%20Ltd')/Inquiry_List?$top=1
 * ─────────────────────────────────────────────────────────────────────────────
 */
const BC_CONFIG = {
  baseUrl:        "https://businesscentral.dynamics.com/Tototheo/",
  company:        "Tototheo Maritime Ltd",
  inquiryListPage: 70355879,

  // ✅ Confirmed OData endpoint
  odataBase:   "https://api.businesscentral.dynamics.com/v2.0/e5296d4f-699b-43ea-a129-066a3f7010e3/Tototheo/ODataV4",

  // ✅ Search uses InquiryCard (has Customer_Ref_No field)
  // ✅ Open button uses Inquiry_List page (70355879) — confirmed
  odataEntity: "InquiryCard",

  // ✅ CONFIRMED field name (exists in InquiryCard & Inquiry_Card_Excel, NOT in Inquiry_List)
  customerRefField: "Customer_Ref_No",

  // BC UI filter label
  bcFilterField: "Customer Ref. No",
};

/**
 * RFQ PATTERNS – tuned for your exact email format:
 * "FW: New RFQ 53700205 from MSC Shipmanagement Ltd – Cyprus for Vessel MSC Richika F"
 */
const RFQ_PATTERNS = [
  /\bNew\s+RFQ\s+(\d{5,12})\b/i,              // ← your exact format
  /\bRFQ[:\s#/-]*(\d{5,12})\b/i,
  /\bRFQ\s*No\.?\s*[:\s]*(\d{5,12})\b/i,
  /\bQuotation\s+(?:No\.?\s*)?(\d{5,12})\b/i,
  /\bRef(?:erence)?\.?\s*[#:\s-]+(\d{5,12})\b/i,
  /\bPO[:\s#-]*(\d{5,12})\b/i,
];

// ─────────────────────────────────────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────────────────────────────────────

let currentRef = "";

// ─────────────────────────────────────────────────────────────────────────────
// OFFICE.JS INIT
// ─────────────────────────────────────────────────────────────────────────────

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    attachEventListeners();
    extractRfqFromEmail();
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// EXTRACT RFQ FROM EMAIL
// ─────────────────────────────────────────────────────────────────────────────

function extractRfqFromEmail() {
  showState("loading");
  const item    = Office.context.mailbox.item;
  const subject = item.subject || "";

  // 1. Try subject first (instant, no async)
  const fromSubject = findRfqInText(subject);
  if (fromSubject) {
    setDetectedState(fromSubject, "detected from subject");
    return;
  }

  // 2. Try body
  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const fromBody = findRfqInText(result.value || "");
      if (fromBody) {
        setDetectedState(fromBody, "detected from email body");
      } else {
        showState("no-detect");
      }
    } else {
      showState("no-detect");
    }
  });
}

function findRfqInText(text) {
  for (const p of RFQ_PATTERNS) {
    const m = text.match(p);
    if (m && m[1]) return m[1];
  }
  return null;
}

function setDetectedState(ref, source) {
  currentRef = ref;
  document.getElementById("detected-ref-value").textContent  = ref;
  document.getElementById("detected-ref-source").textContent = source;
  showState("detected");
}

// ─────────────────────────────────────────────────────────────────────────────
// UI STATE
// ─────────────────────────────────────────────────────────────────────────────

function showState(name) {
  ["loading","detected","no-detect","searching","found","notfound"].forEach(function(s) {
    const el = document.getElementById("state-" + s);
    if (el) el.style.display = (s === name) ? "" : "none";
  });
}

function resetToSearch() {
  currentRef ? showState("detected") : showState("no-detect");
}

// ─────────────────────────────────────────────────────────────────────────────
// EVENT LISTENERS
// ─────────────────────────────────────────────────────────────────────────────

function attachEventListeners() {
  on("btn-search-detected",      "click", function() { searchInquiry(currentRef); });
  on("btn-search-manual",        "click", function() { const v=val("manual-input");       if(v) searchInquiry(v); });
  on("btn-search-manual-2",      "click", function() { const v=val("manual-input-2");     if(v) searchInquiry(v); });
  on("btn-open-list-direct",     "click", function() { openInquiryList(null); });
  on("btn-open-list-notfound",   "click", function() { openInquiryList(currentRef); });
  on("btn-search-again-found",   "click", resetToSearch);
  on("btn-search-again-notfound","click", resetToSearch);
  onEnter("manual-input",   function(v) { searchInquiry(v); });
  onEnter("manual-input-2", function(v) { searchInquiry(v); });
}

function on(id, ev, fn)  { const el=document.getElementById(id); if(el) el.addEventListener(ev,fn); }
function val(id)         { const el=document.getElementById(id); return el ? el.value.trim() : ""; }
function onEnter(id, fn) { on(id,"keydown",function(e){ if(e.key==="Enter"){ const v=val(id); if(v) fn(v); }}); }

// ─────────────────────────────────────────────────────────────────────────────
// SEARCH BC OData
// ─────────────────────────────────────────────────────────────────────────────

async function searchInquiry(ref) {
  currentRef = ref;
  document.getElementById("searching-ref-display").textContent = "Customer Ref: " + ref;
  showState("searching");
  try {
    const results = await queryBCOData(ref);
    results && results.length > 0
      ? renderFoundState(results, ref)
      : renderNotFoundState(ref, false);
  } catch(err) {
    console.error("BC OData error:", err);
    renderNotFoundState(ref, true);
  }
}

/**
 * Query:
 * GET /ODataV4/Company('Tototheo Maritime Ltd')/Inquiry_List
 *     ?$filter=Customer_Ref_No eq '53700205'
 *     &$top=10
 *     &$select=No,Customer_Ref_No,Customer_Name,Vessel_Name,Status,Sales_Operator,Category_Code,Creation_Date
 *
 * Auth: reuses the user's existing M365 browser session — no extra login.
 */
async function queryBCOData(ref) {
  const co   = encodeURIComponent(BC_CONFIG.company);
  const ref2 = ref.replace(/'/g, "''");   // OData escape
  const f    = BC_CONFIG.customerRefField;

  // Fields from InquiryCard entity
  const select = [
    "No", f,
    "Customer_Name",
    "Vessel_Name",
    "Status",
    "Sales_Operator",
    "Category_Code",
    "Creation_Date",
  ].join(",");

  const url = BC_CONFIG.odataBase
    + "/Company('" + co + "')/" + BC_CONFIG.odataEntity
    + "?$filter=" + encodeURIComponent(f + " eq '" + ref2 + "'")
    + "&$top=10"
    + "&$select=" + select;

  const resp = await fetch(url, {
    method:      "GET",
    headers:     { "Accept": "application/json" },
    credentials: "include",   // uses existing M365 session
  });

  if (!resp.ok) throw new Error("HTTP " + resp.status);
  const data = await resp.json();
  return data.value || [];
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER
// ─────────────────────────────────────────────────────────────────────────────

function renderFoundState(results, ref) {
  const container = document.getElementById("results-container");
  container.innerHTML = "";
  document.getElementById("found-count-label").textContent =
    results.length === 1 ? "1 Inquiry found in BC" : results.length + " Inquiries found in BC";

  results.forEach(function(item) {
    const no       = item.No || "—";
    const custRef  = item[BC_CONFIG.customerRefField] || ref;
    const customer = item.Customer_Name  || "";
    const vessel   = item.Vessel_Name    || "";
    const status   = item.Status         || "";
    const operator = item.Sales_Operator || "";
    const category = item.Category_Code  || "";
    const date     = item.Creation_Date  ? item.Creation_Date.substring(0,10) : "";

    const card = document.createElement("div");
    card.className = "inquiry-card";
    card.innerHTML =
      '<div class="inquiry-header">'
        + '<div class="inquiry-id">' + esc(no) + '</div>'
        + '<span class="badge ' + getStatusClass(status) + '">'
          + '<span class="badge-dot"></span>' + esc(status || "—")
        + '</span>'
      + '</div>'
      + row("Ref",      custRef,  true)
      + row("Customer", customer)
      + row("Vessel",   vessel)
      + row("Operator", operator)
      + row("Category", category)
      + row("Date",     date)
      + '<button class="btn btn-primary"'
        + ' style="width:100%;margin-top:10px;font-size:11px;padding:7px"'
        + ' onclick="openInquiryList(\'' + esc(ref) + '\')">'
        + 'Open in Inquiry List →'
      + '</button>';

    container.appendChild(card);
  });

  showState("found");
}

function row(label, value, mono) {
  if (!value) return "";
  return '<div class="inquiry-field"><strong>' + label + ':</strong> '
    + (mono ? '<span class="ref-value">' : '')
    + esc(value)
    + (mono ? '</span>' : '')
    + '</div>';
}

function renderNotFoundState(ref, isError) {
  document.getElementById("notfound-ref-display").textContent = isError
    ? ref + " — could not reach BC. Open manually."
    : "No Inquiry found for Ref: " + ref;
  showState("notfound");
}

// ─────────────────────────────────────────────────────────────────────────────
// OPEN BC
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Opens BC Inquiry List filtered by Customer Ref.
 * URL format confirmed from your BC link:
 * https://businesscentral.dynamics.com/Tototheo/?company=Tototheo%20Maritime%20Ltd&page=70355879
 */
function openInquiryList(ref) {
  let url = BC_CONFIG.baseUrl
    + "?company=" + encodeURIComponent(BC_CONFIG.company)
    + "&page="    + BC_CONFIG.inquiryListPage;

  if (ref) {
    url += "&filter=" + encodeURIComponent("'" + BC_CONFIG.bcFilterField + "' IS '" + ref + "'");
  }
  Office.context.ui.openBrowserWindow(url);
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function getStatusClass(status) {
  const s = (status || "").toLowerCase();
  if (s.includes("receiv") || s.includes("open"))  return "badge-open";
  if (s.includes("quot"))                          return "badge-quoted";
  if (s.includes("clos") || s.includes("cancel"))  return "badge-closed";
  return "badge-draft";
}

function esc(str) {
  return String(str)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");
}
