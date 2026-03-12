# BC Inquiry Checker – Outlook Add-in
## Tototheo Maritime Ltd

Αυτό το Outlook add-in εμφανίζεται κάθε φορά που ανοίγεις ένα email.
Ανιχνεύει αυτόματα το RFQ reference, ψάχνει στο Business Central Inquiry List,
και δείχνει αν υπάρχει ήδη Inquiry ή αν πρέπει να δημιουργηθεί νέο.

---

## 📁 Αρχεία

| Αρχείο | Περιγραφή |
|--------|-----------|
| `manifest.xml` | Ορισμός add-in για Outlook (upload στο Admin Center) |
| `taskpane.html` | Το UI panel που βλέπουν οι χρήστες |
| `taskpane.js` | Logic: διαβάζει email, ψάχνει BC OData, εμφανίζει αποτελέσματα |

---

## 🚀 Deployment – Βήματα

### Βήμα 1: Host τα αρχεία (HTTPS απαιτείται)

Τα αρχεία `taskpane.html` και `taskpane.js` πρέπει να φιλοξενηθούν σε HTTPS URL.

**Επιλογές:**
- **Azure Static Web Apps** (δωρεάν, recommended)
- **Azure Blob Storage** με static website enabled
- **Οποιοδήποτε web server** με SSL certificate

Μόλις ανεβάσεις τα αρχεία, σημείωσε το URL, π.χ.:
```
https://tototheo-addin.azurestaticapps.net
```

### Βήμα 2: Ενημέρωσε το manifest.xml

Άνοιξε το `manifest.xml` και αντικατέστησε **όλες** τις εμφανίσεις του:
```
https://YOUR-HOSTED-URL
```
με το πραγματικό URL σου, π.χ.:
```
https://tototheo-addin.azurestaticapps.net
```

### Βήμα 3: Ενημέρωσε το taskpane.js

Άνοιξε το `taskpane.js` και έλεγξε/ενημέρωσε το `BC_CONFIG`:

```javascript
const BC_CONFIG = {
  baseUrl: "https://businesscentral.dynamics.com/Tototheo/",
  company: "Tototheo Maritime Ltd",         // ← ακριβώς όπως στο BC
  inquiryListPage: 70355879,                // ✓ σωστό
  odataEntity: "InquiryHeaders",            // ← αλλαγή στο σωστό entity name
  customerRefField: "customerReference",    // ← αλλαγή στο σωστό field name
};
```

**Για να βρεις το σωστό OData entity name:**
1. Άνοιξε BC → Inquiry List → Settings → OData
2. Ή ρώτα τον BC administrator για το API endpoint

### Βήμα 4: Deploy το manifest στο Microsoft 365 Admin Center

1. Πήγαινε στο: https://admin.microsoft.com
2. **Settings** → **Integrated Apps** → **Upload custom apps**
3. Επίλεξε "Office Add-in" και upload το `manifest.xml`
4. Assign σε users ή groups (π.χ. η ομάδα των inquiry handlers)
5. Το add-in εμφανίζεται στο Outlook εντός **24 ωρών**

---

## 🔧 Πώς λειτουργεί

```
Χρήστης ανοίγει email (RFQ 53700205)
        ↓
Add-in διαβάζει subject + body
        ↓
Βρίσκει "53700205" με regex
        ↓
Εμφανίζει detected ref → χρήστης κλικάρει "Search"
        ↓
Query στο BC OData: /InquiryHeaders?$filter=customerReference eq '53700205'
        ↓
    ┌──────────────────┬──────────────────┐
    │ ΒΡΕΘΗΚΕ          │ ΔΕΝ ΒΡΕΘΗΚΕ      │
    │ Εμφανίζει card   │ Εμφανίζει ⚠️     │
    │ με status,       │ Κουμπί: Open     │
    │ vessel, assignee │ Inquiry List     │
    │ + Open in BC     │ (για νέο)        │
    └──────────────────┴──────────────────┘
```

---

## 🎯 RFQ Patterns (αυτόματη ανίχνευση)

Το add-in ανιχνεύει αυτόματα references με τα παρακάτω formats:
- `RFQ 53700205`
- `RFQ: 53700205`
- `RFQ#53700205`
- `Ref: 53700205`
- `Reference: 53700205`
- `PO: 53700205`

Αν χρειαστείς περισσότερα patterns, πρόσθεσέ τα στο `RFQ_PATTERNS` array στο `taskpane.js`.

---

## ❓ Συχνές Ερωτήσεις

**Πότε εμφανίζεται το add-in;**
Κάθε φορά που ο χρήστης ανοίγει ένα email. Εμφανίζεται ως panel στα δεξιά.
Υπάρχει και κουμπί στο ribbon "BC Inquiry Checker".

**Τι γίνεται αν το OData API δεν απαντά;**
Εμφανίζεται το κουμπί "Open Inquiry List" που ανοίγει απευθείας τη σελίδα στο BC.

**Χρειάζεται ξεχωριστό login;**
Όχι. Χρησιμοποιεί το ίδιο Microsoft 365 account του χρήστη.

**Τι κάνει το κουμπί "Open in Inquiry List";**
Ανοίγει τη σελίδα 70355879 φιλτραρισμένη με το Customer Ref. Ο χρήστης βλέπει
το Inquiry και μπορεί να το ανοίξει ή να δημιουργήσει νέο από εκεί.
