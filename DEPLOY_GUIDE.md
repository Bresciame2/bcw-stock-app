# BCW Magazzino — Deployment & Integration Guide (enhanced build)

This package turns the BCW stock engine into a hosted, multi-user app over the
**real** `MAGAZZINO BCW fixed.xlsx`, with per-product value, a print-ready
REGISTRO export, and an export/import licence archive.

## Files

| File | Purpose |
|------|---------|
| `stock_agent.py` | Core engine. Carico / scarico / cerca / stato + **stato_per_prodotto**, **export_registro**, **validate_workbook**, atomic save, validation. Writes the real workbook, preserves formulas (cols A & O) and the REGISTRO VLOOKUPs. |
| `storage.py` | Shared-storage backend (local / S3-compatible / Dropbox) with a **lock** so two colleagues can't corrupt the register. |
| `license_archive.py` | Export/Import licence + EUC archive with searchable index and expiry alerts. |
| `doc_numbering.py` | Central, gapless, per-company/per-year document numbering held in the **same shared storage + lock** so concurrent users never mint duplicate invoice numbers (BME-2026-001 / BCW-2026-001 / BME-PL-2026-001). |
| `doc_templates.py` | Faithful Python port of the HTML proforma / packing-list / EUR1+EUC layouts. Renders HTML → PDF (WeasyPrint) and Word (.doc). Logos in `assets/`. |
| `app.py` | Streamlit UI: Dashboard, Carico, Scarico, Cerca, **Documenti**, Registro, Licenze, Impostazioni. |
| `assets/` | `bme_logo.png`, `bcw_logo.png` embedded into generated documents. |
| `requirements.txt` | Python dependencies. |
| `packages.txt` | **System** libs Streamlit Cloud must apt-install for WeasyPrint (pango/cairo). Without it, PDF generation fails (Word still works). |
| `make_synthetic_workbook.py` | Builds a fake test workbook (do not use in production). |

## Why shared storage is mandatory on Streamlit Cloud

Streamlit Cloud's local disk is **ephemeral** — wiped on every restart, sleep,
or redeploy. A register written to local disk there will silently lose data and
is not shared between colleagues. So the authoritative workbook must live in
durable shared storage that every session reads and writes. `storage.py` does
this and adds a lock so concurrent saves can't corrupt the file.

`BCW_STORAGE=local` is for development on one machine only.

## Choose a backend

Pick **one**. S3-compatible is the most robust for a shared register; Dropbox is
easiest if the shop already uses it.

### Option A — Cloudflare R2  ← CHOSEN
R2 is S3-compatible, so it uses the `s3` backend with an endpoint. A full
fill-in-the-blanks template is in **`secrets.toml.example`**. Set in Streamlit secrets:
```
BCW_STORAGE = "s3"
BCW_S3_BUCKET = "bcw-magazzino"
BCW_S3_PREFIX = ""
BCW_S3_ENDPOINT = "https://<ACCOUNT_ID>.r2.cloudflarestorage.com"   # from R2 dashboard
AWS_ACCESS_KEY_ID = "<R2 access key id>"
AWS_SECRET_ACCESS_KEY = "<R2 secret>"
AWS_REGION = "auto"                # R2 requires the literal value: auto
BCW_WORKBOOK_NAME = "MAGAZZINO BCW fixed.xlsx"
ANTHROPIC_API_KEY = "sk-ant-…"
```
The same keys work for AWS S3 (omit `BCW_S3_ENDPOINT`, use a real region) and
Backblaze B2 (set its S3 endpoint).

### Option B — Dropbox
```
BCW_STORAGE = "dropbox"
BCW_DROPBOX_DIR = "/bcw"
BCW_DROPBOX_TOKEN = "…"            # or the refresh-token trio below
# BCW_DROPBOX_APP_KEY / BCW_DROPBOX_APP_SECRET / BCW_DROPBOX_REFRESH_TOKEN
BCW_WORKBOOK_NAME = "MAGAZZINO BCW fixed.xlsx"
ANTHROPIC_API_KEY = "sk-ant-…"
```
`requirements.txt` currently keeps only `boto3` (R2); uncomment `dropbox` if you switch.

### Verify the backend before go-live
After setting the secrets (or the equivalent env vars locally), run the
connection self-test — it proves read/write/lock work and that the register
object is present, without touching your data:
```bash
python test_storage.py
```

## Deploy steps (Streamlit Cloud)

1. Put `app.py`, `stock_agent.py`, `storage.py`, `license_archive.py`,
   `doc_numbering.py`, `doc_templates.py`, the `assets/` folder,
   `requirements.txt` **and `packages.txt`** in a Git repo.
2. Create the app on https://streamlit.io/cloud pointing at `app.py`.
3. Paste the secrets above into the app's **Settings → Secrets**.
4. Open the app → **Impostazioni** → upload your real `MAGAZZINO BCW fixed.xlsx`
   to initialise the shared storage (one-time).
5. Each colleague enters their name in the sidebar (used for the lock message).

## Run locally
```bash
pip install -r requirements.txt
export BCW_STORAGE=local
export BCW_LOCAL_DIR="$PWD/data"
export ANTHROPIC_API_KEY=sk-ant-...
streamlit run app.py
```

## ⚠️ Verify against your REAL workbook before relying on it

This build was tested end-to-end against a **synthetic** workbook matching the
documented column contract (carico, scarico, double-sale guard, validation,
per-product totals, REGISTRO VLOOKUP add + scarico update, registro export,
formula preservation, shared-storage lock — all pass). Your real
`MAGAZZINO BCW fixed.xlsx` was not included, so confirm these against it:

- The MAGAZZINO data starts on **row 3** and the column order matches `COL` in
  `stock_agent.py` (A=GIACENZA … X=CLIENTE).
- `find_last_data_row` lands on the right next-free row (it keys on col C date +
  col L fornitore; check your reserved-N.OP slots and summary rows are skipped).
- The REGISTRO header offset (data from **row 4**) and VLOOKUP range match.
- Run **Impostazioni → Verifica struttura workbook** after uploading — it checks
  the formula columns are still formulas.

Make a copy of the real file and run a test carico/scarico on the copy first.

## Safety invariants (unchanged, do not break)
- Cols **A** and **O** in MAGAZZINO are formulas — the engine rewrites them per
  row; never replace with static values.
- Workbook opened with `data_only=False` so formulas survive saves.
- Every save makes a `*_backup.xlsx` and writes atomically (temp + replace).
- Always confirm `data_carico` (arrival at Gardone) with a human — it often
  differs from the DDT/invoice date. The Carico tab shows this reminder.

## What's new vs. the original package
- Per-product (TIPOLOGIA) value breakdown — `stato_per_prodotto()` + Dashboard.
- One-click print-ready REGISTRO export (Excel + PDF) — `export_registro()`.
- Export/Import licence + EUC archive with search and expiry alerts.
- Carico validation (tipologia/date/cost) so a bad entry can't reach the file.
- Atomic save and `validate_workbook()` hardening.
- Shared-storage backend + lock for safe hosted multi-user use.
- **Documenti tab** — generate proforma invoices (BME $/BCW €), packing lists
  (parcel-based, dual serials, shipment summary) and EUR1+EUC declarations, each
  as PDF + Word. Numbers are allocated centrally and gaplessly via
  `doc_numbering` (same storage + lock as the workbook), so concurrent colleagues
  never get duplicate invoice numbers; declarations consume no number.
- **Stock → line items** — for BCW, a "Aggiungi da magazzino" picker pulls
  in-stock guns (`stock_agent.list_in_stock`) straight into invoice/packing lines
  by matricola/marca/modello/calibro/costo. Non-destructive: it does not deduct
  stock (sale is recorded via the Scarico tab).
