const cfg = window.ADDIN_CONFIG;

const RFQ_PATTERNS = [
  /\bNew\s+RFQ\s+(\d{5,12})\b/i,
  /\bRFQ[:\s#/-]*(\d{5,12})\b/i,
  /\bRFQ\s*No\.?\s*[:\s]*(\d{5,12})\b/i,
  /\bQuotation\s+(?:No\.?\s*)?(\d{5,12})\b/i,
  /\bRef(?:erence)?\.?\s*[#:\s-]+(\d{5,12})\b/i,
  /\bPO[:\s#-]*(\d{5,12})\b/i,
];

let currentRef = "";

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    attachEventListeners();
    extractRfqFromEmail();
  }
});

// -----------------------------
// RFQ detection
// -----------------------------
function extractRfqFromEmail() {
  showState("loading");

  const item = Office.context.mailbox.item;
  const subject = item.subject || "";

  const fromSubject = findRfqInText(subject);
  if (fromSubject) {
    setDetectedState(fromSubject, "detected from subject");
    return;
  }

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
  document.getElementById("detected-ref-value").textContent = ref;
  document.getElementById("detected-ref-source").textContent = source;
  showState("detected");
}

// -----------------------------
// UI state
// -----------------------------
function showState(name) {
  ["loading", "detected", "no-detect"].forEach(function (s) {
    const el = document.getElementById("state-" + s);
    if (el) el.style.display = (s === name) ? "" : "none";
  });
  
  ["searching", "found", "notfound"].forEach(function (s) {
    const el = document.getElementById("state-" + s);
    if (el) el.style.display = "none";
  });
}

function resetToSearch() {
  currentRef ? showState("detected") : showState("no-detect");
}

// -----------------------------
// Events
// -----------------------------
function attachEventListeners() {
  on("btn-search-detected", "click", function () {
    openInquiryList(currentRef);
  });

  on("btn-search-manual", "click", function () {
    const v = val("manual-input");
    if (v) openInquiryList(v);
  });

  on("btn-search-manual-2", "click", function () {
    const v = val("manual-input-2");
    if (v) openInquiryList(v);
  });

  on("btn-open-list-direct", "click", function () {
    openInquiryList(null);
  });

  on("btn-open-list-notfound", "click", function () {
    openInquiryList(currentRef);
  });

  on("btn-search-again-found", "click", resetToSearch);
  on("btn-search-again-notfound", "click", resetToSearch);

  onEnter("manual-input", function (v) {
    openInquiryList(v);
  });

  onEnter("manual-input-2", function (v) {
    openInquiryList(v);
  });
}

function on(id, ev, fn) {
  const el = document.getElementById(id);
  if (el) el.addEventListener(ev, fn);
}

function val(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}

function onEnter(id, fn) {
  on(id, "keydown", function (e) {
    if (e.key === "Enter") {
      const v = val(id);
      if (v) fn(v);
    }
  });
}

// -----------------------------
// Token handling
// -----------------------------
function getStoredToken() {
  try {
    const raw = localStorage.getItem("bcApiToken");
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    if (!parsed.accessToken || !parsed.expiresOn) return null;

    const expires = new Date(parsed.expiresOn).getTime();
    const now = Date.now();

    // refresh 2 minutes before expiry
    if (expires - now < 120000) {
      localStorage.removeItem("bcApiToken");
      return null;
    }

    return parsed.accessToken;
  } catch {
    return null;
  }
}

function storeToken(accessToken, expiresOn) {
  localStorage.setItem("bcApiToken", JSON.stringify({
    accessToken,
    expiresOn
  }));
}

function openAuthDialog() {
  return new Promise(function (resolve, reject) {
    const dialogUrl = window.location.origin + window.location.pathname.replace(/\/[^/]*$/, "/") + "auth.html";

    Office.context.ui.displayDialogAsync(dialogUrl, {
      height: 60,
      width: 35,
      displayInIframe: false
    }, function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error("Could not open authentication dialog."));
        return;
      }

      const dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        try {
          const msg = JSON.parse(arg.message);

          if (msg.type === "auth-success") {
            if (msg.accessToken) {
              storeToken(msg.accessToken, msg.expiresOn);
              dialog.close();
              resolve(msg.accessToken);
              return;
            }
          }

          if (msg.type === "auth-error") {
            dialog.close();
            reject(new Error(msg.message || "Authentication failed."));
            return;
          }

          dialog.close();
          reject(new Error("Unknown authentication response."));
        } catch (e) {
          dialog.close();
          reject(new Error("Invalid dialog response."));
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        reject(new Error("Authentication dialog was closed or blocked."));
      });
    });
  });
}

async function ensureBcToken() {
  const cached = getStoredToken();
  if (cached) return cached;
  return await openAuthDialog();
}

// -----------------------------
// BC search
// -----------------------------
async function searchInquiry(ref) {
  currentRef = ref;
  document.getElementById("searching-ref-display").textContent = "Customer Ref: " + ref;
  showState("searching");

  try {
    const accessToken = await ensureBcToken();
    const results = await queryBCOData(ref, accessToken);

    results && results.length > 0
      ? renderFoundState(results, ref)
      : renderNotFoundState(ref, false);
  } catch (err) {
    console.error("BC OData error:", err);
    renderNotFoundState(ref, true, err);
  }
}

async function queryBCOData(ref, accessToken) {
  const co = encodeURIComponent(cfg.company);
  const ref2 = ref.replace(/'/g, "''");
  const f = cfg.customerRefField;

  const select = [
    "No", f,
    "Customer_Name",
    "Vessel_Name",
    "Status",
    "Sales_Operator",
    "Category_Code",
    "Creation_Date"
  ].join(",");

  const url = cfg.odataBase
    + "/Company('" + co + "')/" + cfg.odataEntity
    + "?$filter=" + encodeURIComponent(f + " eq '" + ref2 + "'")
    + "&$top=10"
    + "&$select=" + select;

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      "Accept": "application/json",
      "Authorization": "Bearer " + accessToken
    }
  });

  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error("BC API HTTP " + resp.status + (text ? " - " + text : ""));
  }

  const data = await resp.json();
  return data.value || [];
}

// -----------------------------
// Render
// -----------------------------
function renderFoundState(results, ref) {
  const container = document.getElementById("results-container");
  container.innerHTML = "";

  document.getElementById("found-count-label").textContent =
    results.length === 1 ? "1 Inquiry found in BC" : results.length + " Inquiries found in BC";

  results.forEach(function (item) {
    const no = item.No || "—";
    const custRef = item[cfg.customerRefField] || ref;
    const customer = item.Customer_Name || "";
    const vessel = item.Vessel_Name || "";
    const status = item.Status || "";
    const operator = item.Sales_Operator || "";
    const category = item.Category_Code || "";
    const date = item.Creation_Date ? item.Creation_Date.substring(0, 10) : "";

    const card = document.createElement("div");
    card.className = "inquiry-card";
    card.innerHTML =
      '<div class="inquiry-header">'
      + '<div class="inquiry-id">' + esc(no) + '</div>'
      + '<span class="badge ' + getStatusClass(status) + '">'
      + '<span class="badge-dot"></span>' + esc(status || "—")
      + '</span>'
      + '</div>'
      + row("Ref", custRef, true)
      + row("Customer", customer)
      + row("Vessel", vessel)
      + row("Operator", operator)
      + row("Category", category)
      + row("Date", date)
      + '<button class="btn btn-primary"'
      + ' style="width:100%;margin-top:10px;font-size:11px;padding:7px"'
      + ' onclick="openInquiryList(\'' + esc(custRef) + '\')">'
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

function renderNotFoundState(ref, isError, err) {
  document.getElementById("notfound-ref-display").textContent = isError
    ? ref + " — authentication/API search failed. Use Open Inquiry List."
    : "No Inquiry found for Ref: " + ref;

  console.warn("Inquiry not found or API issue:", err || "");
  showState("notfound");
}

// -----------------------------
// Open BC UI
// -----------------------------
function openInquiryList(ref) {
  let url = cfg.baseUrl
    + "?company=" + encodeURIComponent(cfg.company)
    + "&page=" + cfg.inquiryListPage;

  if (ref) {
    url += "&filter=" + encodeURIComponent("'" + cfg.bcFilterField + "' IS '" + ref + "'");
  }

  Office.context.ui.openBrowserWindow(url);
}

// -----------------------------
// Helpers
// -----------------------------
function getStatusClass(status) {
  const s = (status || "").toLowerCase();
  if (s.includes("receiv") || s.includes("open")) return "badge-open";
  if (s.includes("quot")) return "badge-quoted";
  if (s.includes("clos") || s.includes("cancel")) return "badge-closed";
  return "badge-draft";
}

function esc(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
