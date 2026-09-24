# Courier Details Input Portal (a.k.a. "Dispatch & Dimensions Portal")

A single-page, no-build-step web app that replaces manual invoice-to-Excel typing for
courier manifests. An operator types an Order ID, the app looks the order up in
Supabase and auto-fills shipment data, the operator adds physical package details,
and successful entries accumulate in an on-screen "Active Manifest" that can be
exported to a courier-ready Excel workbook at any time.

This is a **separate, standalone app** — not part of, and not to be confused with,
`arovehic-app-order`.

---

## 1. Tech stack

- **Frontend:** Plain HTML/CSS/JS. No bundler, no framework, no build step. UI markup
  uses Tailwind CSS (via CDN) and Phosphor Icons (via CDN).
- **Backend:** None — this app talks directly to Supabase from the browser using the
  Supabase JS client (via CDN), with the anon/publishable key. There is no server
  component.
- **Excel generation:** [ExcelJS](https://github.com/exceljs/exceljs) (via CDN), used
  to load the source template, write rows into it in-browser, and produce a
  downloadable `.xlsx` file that preserves the template's headers/formatting.
- **Persistence:** `localStorage`, single key (`courierManifest`), holding the
  in-progress Active Manifest as JSON. This is a single-user, single-computer tool —
  there is no server-side or multi-device sync.

## 2. Project structure

```
courier-manifest-app/
├── index.html          # UI markup (Tailwind-based), loads the 3 CDN scripts + app.js
├── style.css            # (legacy — the redesigned index.html is mostly self-styled
│                         via Tailwind + one inline <style> block; this file may be
│                         largely unused post-redesign, verify before relying on it)
├── app.js                # All application logic — see section 4
└── assets/
    └── bulkdomesticone_28042026_.xlsx   # THE source Excel template — see section 6
```

`app.js` targets a fixed set of element IDs and is intentionally decoupled from the
visual design — the UI was redesigned once already (Tailwind rewrite of `index.html`)
without touching `app.js` at all, by keeping the same IDs. See section 7 for the
full ID contract if redesigning again.

## 3. Setup

1. Open `app.js` and set your real Supabase credentials:
   ```js
   const SUPABASE_URL = 'https://<your-project-ref>.supabase.co';
   const SUPABASE_ANON_KEY = '<your anon/publishable key>';
   ```
   This is the **publishable/anon** key, safe for client-side use — but only because
   Row Level Security (RLS) on the `orders` table restricts the anon role to
   `SELECT` only. Never put a `service_role` key here.
2. Serve the folder with any static file server (e.g. VS Code's "Live Server"
   extension). It cannot be opened as a bare `file://` page reliably, since `fetch()`
   is used to load the Excel template asset.
3. Confirm `assets/bulkdomesticone_28042026_.xlsx` exists and is the correct,
   current template (see section 6 — this file gets edited directly, not just read).

## 4. Application flow (`app.js`)

1. **Lookup** (`performLookup`, triggered on **Enter** in the Order ID field) — see
   section 5.1 for which Supabase column this actually queries (it is **not**
   `shopify_order_id`, despite that being the table's primary key).
2. **Data cleaning** (`cleanValue`, `cleanPhoneNumber`, `buildAddressLines`) — see
   section 5.2–5.4.
3. **Validation + row build** (`validateManualFields`, `buildManifestRow`) — combines
   the looked-up Supabase row, the 5 manually-entered package fields, and a fixed set
   of constant values (section 5.5) into one row object with 48 keys matching the
   Excel template's columns.
4. **Duplicate check** (`isDuplicate`) — checked only against the *current session's*
   Active Manifest entries (`manifestEntries` in memory), not against Supabase or any
   historical data.
5. **Render + persist** (`renderManifestRow`, `saveManifestToStorage`) — appends the
   row to the on-screen table and writes the whole Active Manifest state to
   `localStorage`.
6. **Reset** (`resetManifest`, wired to the "Reset Data (Keep Template)" dropdown) —
   clears `manifestEntries`, resets the serial counter to 1, clears `localStorage`.
   Does **not** touch the Excel template asset or Supabase.
7. **Download** (`downloadExcel`) — fetches a fresh copy of the template asset,
   writes every current Active Manifest entry starting at row 2, and triggers a
   browser download named `bulkdomesticone_<DDMMYYYY>_.xlsx` (today's date, computed
   fresh on every download — not the literal date in the template's own filename).
   The Active Manifest is **not** cleared after download.
8. **Clipboard** (`copyPincodeBtn` handler) — copies the Pincode field's value via
   `navigator.clipboard.writeText`, with a caught-failure fallback message.

## 5. Business logic reference

### 5.1 Order ID lookup — queries `order_name`, NOT `shopify_order_id`

This is the single most important non-obvious fact about this app. `shopify_order_id`
is the table's primary key, but it is an internal ID that never appears on a printed
invoice. The Order ID an operator actually reads off a physical invoice (e.g. `#7720`,
or manual entries like `KAR 19`) is stored in the **`order_name`** column.

The lookup normalizes the operator's input to accept it with or without a leading
`#` (queries both `order_name.eq.<value>` and `order_name.eq.#<value>` via `.or()`).
It does **not** handle the `KAR 19`-style manual format specially — those must be
typed exactly as stored.

`shopify_order_id` is still used internally as the unique key for the
in-session duplicate check, since it's guaranteed unique (unlike `order_name`, which
has no unique constraint in the schema).

### 5.2 Payment type / COD logic — derived from `invoice_number`, NOT `payment_type`

The `orders` table has a `payment_type` column, but in real production data it is
**always null**. Both of the following are instead derived from the `invoice_number`
column, normalized (trimmed, uppercased):

- **On-screen Payment Type display:** `invoice_number` starts with `COD` → shows
  `COD`; starts with `SHP` → shows `SHP`; otherwise blank.
- **Excel `CODR/COD` column:** `COD` if `invoice_number` starts with `COD`
  (case-insensitive), else blank.
- **Excel `VALUE FOR CODR/COD` column:** only populated for COD orders (see 5.3 for
  which field it reads from), else blank. Guarded against `null` — a COD order with
  no amount value will submit successfully with this field left blank rather than
  crashing.

### 5.3 Order Price / COD value — reads `amount_to_receive`, NOT `balance_due`

`balance_due` looked like the right field (it's literally named for a pending
amount) but turned out to represent something else in this dataset — confirmed
against real orders where `balance_due` didn't match either the printed invoice
total or the "amount to receive" shown in another internal dashboard for the same
order. Rather than deriving a formula, a dedicated `amount_to_receive` column was
added to the `orders` table upstream, and this app now reads that directly for:

- The read-only **Order Price** display field.
- The Excel **`VALUE FOR CODR/COD`** column (COD orders only).

Both are rounded to 2 decimal places (`Number(x.toFixed(2))`) to avoid JavaScript
floating-point artifacts (e.g. `296.45000000000005`).

**Do not silently revert to `balance_due`** if refactoring this — it is not
equivalent, and using it will produce financially wrong Excel output.

### 5.4 Data cleaning pipeline (`cleanValue`)

Applied to `customer_name`, `shipping_address1`, `shipping_address2`,
`shipping_city`, `shipping_state` before use:

1. `null`/`undefined` → `''`
2. Trimmed value case-insensitively equal to the literal string `"null"` → `''`
3. Standalone `null` tokens embedded mid-string removed (word-boundary,
   case-insensitive), collapsing any resulting double spaces/dangling commas
4. Trim leading/trailing whitespace and stray leading/trailing commas
5. Collapse repeated internal spaces to one

### 5.5 Receiver address line 1/2 rule (`buildAddressLines`)

```
line1 = clean(shipping_address1)
line2Base = clean(shipping_address2) OR clean(shipping_city)   // fallback
overflow = []
while length(line1) >= 50:
    words = split(line1, " ")
    overflow.unshift(words.pop())     // last word out, kept in original order
    line1 = join(words, " ")
line2 = overflow.length ? join(overflow, " ") + " " + line2Base : line2Base
while length(line2) >= 50:
    words = split(line2, " ")
    words.pop()
    line2 = join(words, " ")
```

In words: if Line 1 is ≥50 characters, whole words are moved off its end (never
splitting a word) until it's under 50. Those removed words are then placed at the
**front** of Line 2 (ahead of whatever was already there — real `shipping_address2`
or the city fallback). If the resulting merged Line 2 is itself ≥50 characters,
words are trimmed from its end the same way. This merge behavior applies
unconditionally, regardless of whether Line 2's base content came from real
`shipping_address2` or the city fallback.

### 5.6 Receiver phone cleaning (`cleanPhoneNumber`)

`customer_phone` is stripped of all non-digit characters, then reduced to the last
10 digits if longer than 10 — this removes `+91`/`91` country-code prefixes and any
stray spaces/dashes, regardless of exact input format.

### 5.7 Fixed values (identical on every row, never derived)

| Field | Value |
|---|---|
| SHAPE OF ARTICLE | `NROL` |
| PRIORITY FLAG | `TRUE` |
| DELIVERY INSTRUCTION | `ND` |
| INSTRUCTION RTS | `RTS` |
| SENDER NAME | `AROVEHIC` |
| SENDER COMPANY | `ARIHANTH COMPLEX,3RD FLOOR` |
| SENDER ADD LINE 1 | `J C ROAD,A M ROAD,` |
| SENDER ADD LINE 2 | `BENGALURU-560002` |
| SENDER CITY | `BENGALURU` |
| SENDER STATE | `KARNATAKA` |
| SENDER PINCODE | `560002` |
| SENDER EMAILID | `arovehic@gmail.com` |
| SENDER MOBILE NO | `6360818919` |
| ALT ADDRESS FLAG | `False` |
| PICKUP ADDRESS FLAG | `False` |
| DROP OFF PINCODE | `560001` |
| DROPOFF/PICKUP OFFICE ID | `21250001` |
| ACK | `FALSE` |
| REGISTRATION | `FALSE` |
| OTP BASED DELIVERY | `FALSE` |

Defined once in the `FIXED_VALUES` object in `app.js`.

### 5.8 Physical dimensions — whole numbers only

Weight is entered/stored in **grams**, and Length/Breadth-Diameter/Height in
**centimeters** — all as whole numbers, no decimals. This matches what the real
Excel template expects (its own "Information" sheet states decimals aren't
permitted for these field types).

## 6. Excel template — `assets/bulkdomesticone_28042026_.xlsx`

**This file is treated as a live, hand-maintained source template, not a static
asset.** It has been edited directly in Excel more than once during development —
most recently to permanently delete 17 pre-existing historical data rows (so new
entries now start writing at row 2), and to clear a misapplied data-validation
dropdown restriction. Anyone maintaining this app needs to know:

- New Active Manifest entries are written into the `ArticleDetails` sheet starting
  at **row 2** (row 1 is headers). This assumes rows 2+ are empty in the template —
  if the template ever accumulates leftover data again, either clear it manually
  before use, or add explicit row-clearing logic to `downloadExcel` (not currently
  implemented).
- **Known template quirk (fixed once, could reappear if the template is edited
  again):** the original template had a data-validation dropdown (`ROLL`/`NROL`/`DOC`,
  intended for the "Shape of Article" column) misapplied to a row range in **column
  C (Physical Weight)** instead of column D. This was cleared via Excel's Data →
  Data Validation → Clear All on both columns C and D. If Excel ever throws "This
  value doesn't match the data validation restrictions defined for this cell" again
  on a numeric column, this is almost certainly the cause — check both columns C
  and D for stray validation rules, and always verify you're editing the actual
  `assets/` file (not a downloaded output copy — this has bitten us once already).
- The workbook has 4 sheets total: `ArticleDetails` (the only one this app writes
  to), `PickupAddress`, `AltAddress`, `Information` (an instructions/legend sheet).
  None of these three are read or written by the app; `ALT ADDRESS FLAG` and
  `PICKUP ADDRESS FLAG` are always `False` on every row, so the `PickupAddress`/
  `AltAddress` sheets are never actually needed by this workflow.
- No formulas exist anywhere in the workbook.

## 7. UI element ID contract (for redesigning `index.html` again)

`app.js` is written against fixed element IDs and does not care about markup
structure, classes, or styling framework — this is by design, so the UI can be
redesigned without touching logic. If rebuilding `index.html` again, these IDs
must all exist and behave as described:

| ID | Type | Purpose |
|---|---|---|
| `orderId` | text input | Order ID entry; Enter key triggers lookup |
| `paymentType` | readonly input | Auto-filled Payment Type display |
| `orderPrice` | readonly input | Auto-filled Order Price display |
| `pincode` | readonly input | Auto-filled Pincode display |
| `copyPincodeBtn` | button | Copies Pincode value to clipboard on click |
| `copiedNotice` | any element | Toggled `.show` class briefly on successful copy |
| `barcode`, `weight`, `length`, `breadth`, `height` | inputs | Manual package data |
| `formError` | any element | Error/status message text is written here |
| `submitBtn` | button | Triggers validation + row build + append |
| `resetSelect` | `<select>` | Must have an `<option value="reset">` |
| `downloadBtn` | button | Triggers Excel generation/download |
| `manifestBody` | `<tbody>` | Rows are appended here by `renderManifestRow` |
| (table itself) | `<table>` | No ID required by `app.js`, but has `id="manifestTable"` in the current markup for CSS gridline targeting |

`app.js` also expects 3 `<script>` tags present, in this order, before its own tag:
ExcelJS CDN, Supabase JS CDN, then `app.js` itself.

## 8. Known open items / not yet implemented

- **`KAR 19`-style manual order names** are not normalized the way `#7720`-style
  ones are — the operator must type the exact stored format for these.
- **Embedded mid-string `null` cleaning** (e.g. `"House 12, null Road"` →
  `"House 12, Road"`) is implemented in `cleanValue` but has only been verified for
  the whole-field-null case in live testing, not the embedded case specifically.
- **No `KAR`/manual-format normalization, no partial-ID search, no fuzzy matching**
  — lookup is exact-match only (with the `#`-optional exception).
- The `style.css` file's relationship to the current Tailwind-based `index.html` is
  unclear post-redesign — verify whether it's still linked/needed before removing or
  relying on it.

## 9. Testing notes

Every core workflow (valid/invalid/duplicate lookups, null-field cleaning, long
address overflow across both address lines, empty-address2 city fallback, COD vs
SHP branching, refresh/reopen persistence, download-doesn't-clear, reset-clears,
serial numbering, clipboard copy, Excel column integrity) has been manually tested
against both synthetic test data and real production Supabase rows. There is no
automated test suite — all verification has been manual, screenshot-confirmed
step-by-step during development.
