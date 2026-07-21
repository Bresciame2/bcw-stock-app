"""
BCW Stock Management — Web App (enhanced, shared-storage)
========================================================
Streamlit front-end over the real MAGAZZINO workbook held in durable shared
storage. Capabilities:
  • Dashboard with per-product (TIPOLOGIA) value breakdown
  • Carico/Scarico from scanned DDT / invoices (Claude extraction) with validation
  • Search the stock
  • One-click print-ready REGISTRO export (Excel + PDF)
  • Export/Import licence + EUC archive
Persistence: see storage.py. Set BCW_STORAGE (local | s3 | dropbox) and the
related secrets. On Streamlit Cloud you MUST use s3 or dropbox — local disk is
wiped on restart.

Bilingual UI: a sidebar switch toggles the whole interface between English
(default) and Italian. All user-facing strings go through T("it", "en").
"""
import io
import os
import json
import re
import time
import base64
from datetime import datetime

import streamlit as st
import pandas as pd

import stock_agent
import storage
import doc_numbering
import doc_templates
import bulk_io
from license_archive import LicenseArchive, DOC_TYPES
from doc_archive import DocArchive, DOC_TYPE_LABELS

# Optional barcode support (QR + Code128 labels / scanning). If the libraries
# are missing the app still runs — the Barcode UI just shows a notice.
try:
    import barcode_agent
    _BARCODE_OK = True
except Exception:
    barcode_agent = None
    _BARCODE_OK = False

st.set_page_config(page_title="BCW Magazzino", page_icon="🔫", layout="wide")


# ── i18n ──────────────────────────────────────────────────────────────────────
# Default the interface to English. The sidebar switch (built in main()) writes
# st.session_state["lang"] = "en" | "it". T() is defined before the password gate
# so every screen — including the login page — renders in the chosen language.
if "lang" not in st.session_state:
    st.session_state["lang"] = "en"


def T(it, en=None):
    """Return the English string when the UI language is English, else Italian.
    Usage: T("Salva", "Save"). With one argument, returns it unchanged (useful
    for strings that are identical in both languages or are pure data)."""
    if en is None:
        return it
    return en if st.session_state.get("lang", "en") == "en" else it


# ── access gate ───────────────────────────────────────────────────────────────
# The app is deployed as a "public" Streamlit app (the one private slot is taken),
# so a password gate keeps it restricted to colleagues. Set APP_PASSWORD in the
# Streamlit secrets. Everything below the gate only runs once authenticated.
def _check_password():
    expected = None
    try:
        expected = st.secrets["APP_PASSWORD"]
    except Exception:
        expected = os.environ.get("APP_PASSWORD")
    if not expected:
        st.error(T("APP_PASSWORD non configurata. Aggiungila nei secrets dell'app "
                   "(Streamlit Cloud → Settings → Secrets) per abilitare l'accesso.",
                   "APP_PASSWORD is not configured. Add it to the app secrets "
                   "(Streamlit Cloud → Settings → Secrets) to enable access."))
        st.stop()
    if st.session_state.get("_authed"):
        return
    st.markdown("## 🔒 BCW Magazzino")
    st.caption(T("Inserisci la password per accedere.",
                 "Enter the password to sign in."))
    pw = st.text_input(T("Password", "Password"), type="password", key="_pw_input")
    if st.button(T("Entra", "Enter")):
        if pw == expected:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error(T("Password errata.", "Wrong password."))
    if not st.session_state.get("_authed"):
        st.stop()


_check_password()

# Live typology list, read from the workbook's INVENTARIO dropdown source so the
# Carico selectbox always mirrors the real menu (falls back to the constant).
try:
    TIPOLOGIE = sorted(stock_agent.load_valid_tipologie())
except Exception:
    TIPOLOGIE = sorted(stock_agent.VALID_TIPOLOGIE)


# ── config / helpers ──────────────────────────────────────────────────────────
def get_api_key():
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return os.environ.get("ANTHROPIC_API_KEY", st.session_state.get("api_key", ""))


def current_user():
    return st.session_state.get("user_name") or "app"


@st.cache_resource
def get_backend():
    return storage.get_backend()


def backend_ready(backend):
    try:
        return backend.exists(storage.workbook_name())
    except Exception as e:
        st.error(T(f"Errore accesso storage: {e}", f"Storage access error: {e}"))
        return False


# ── Claude extraction (lifted from the reference app) ─────────────────────────
def pdf_to_images_b64(pdf_bytes):
    try:
        import fitz
        from PIL import Image as PILImage
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        images = []
        for page in doc:
            pix = page.get_pixmap(dpi=120)
            img = PILImage.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            if max(img.size) > 1400:
                img.thumbnail((1400, 1400), PILImage.LANCZOS)
            buf = io.BytesIO(); img.save(buf, format="JPEG", quality=85)
            images.append(base64.standard_b64encode(buf.getvalue()).decode("utf-8"))
        return images, "image/jpeg"
    except ImportError:
        st.error(T("PyMuPDF o Pillow non installato.",
                   "PyMuPDF or Pillow is not installed."))
        return [], "image/jpeg"


def extract_with_claude(doc_bytes, mime_type, operation_type, api_key):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    if operation_type == "CARICO":
        fields_prompt = """Documento: DDT o fattura acquisto armeria. Estrai OGNI articolo.
Campi per oggetto (usa null se assente, stringhe BREVI):
tipologia,calibro,marca,modello,matr_arma,matr_canna,matr_agg,fornitore,costo,costo_imballo,data_ddt,data_fattura,n_fattura,quantita
- tipologia: categoria arma (es. "PISTOLA SEMIAUTOMATICA")
- matr_arma: matricola o "N/D"
- costo: prezzo unitario (numero)
- date: formato DD/MM/YYYY
- quantita: intero (default 1)
- IMPORTANTE — UN PEZZO PER RIGA: se quantità > 1, genera UN oggetto per ogni
  pezzo fisico (ripeti l'oggetto tante volte quanta è la quantità), ognuno con
  quantita=1. Il numero totale di oggetti deve essere uguale alla somma delle
  quantità di tutte le righe.
- Matricole: se il documento indica un intervallo (es. "Mat: Da: 371123 A:
  371129") assegna a ciascun pezzo la sua matricola consecutiva (371123, 371124,
  … 371129). Se è indicata una sola matricola, usala per quel pezzo. Accessori
  senza matricola: "N/D".
Output: array JSON compatto, NESSUN testo fuori dal JSON. [{...},{...}]"""
    elif operation_type == "PERMIT":
        fields_prompt = """Documento: PERMESSO/AUTORIZZAZIONE DI ESPORTAZIONE armi
(licenza di esportazione, autorizzazione UAMA, nulla osta, carnet ATA, EUC).
Estrai UN SOLO oggetto JSON con questi campi (null se assente):
numero,data_emissione,data_scadenza,paese_destinazione,cliente,tipo,articoli
- numero: numero del permesso / ATA / autorizzazione (questo è il N. ATA)
- date: formato DD/MM/YYYY
- paese_destinazione: paese di destinazione dell'export
- cliente: destinatario / end user
- tipo: uno tra "Licenza Esportazione","Autorizzazione UAMA","EUC (End User Certificate)","Transito / Brokering","Altro"
- articoli: array delle armi elencate nel permesso, UN oggetto per arma con
  {matricola,marca,modello,calibro} (array vuoto se non elencate). Riporta la
  matricola ESATTA come scritta.
Output: UN SOLO oggetto JSON {...}, nessun testo fuori dal JSON."""
    else:
        fields_prompt = """Documento: fattura vendita armeria. Estrai OGNI arma venduta.
Campi per oggetto (usa null se assente):
matr_arma,marca,modello,calibro,tipologia,cliente,prezzo_vendita,data_scarico,n_fattura,data_fattura_vendita,n_ata,titolo_acquisto
- marca: produttore (es. "BERETTA"); modello; calibro (es. "9X21")
- date: formato DD/MM/YYYY
- prezzo_vendita: numero
- cliente: nome breve
Output: array JSON compatto, NESSUN testo fuori dal JSON. [{...},{...}]"""
    system_msg = ("Sei un assistente magazzino armeria italiana. Rispondi SOLO con JSON valido, "
                  "array compatto, nessun testo prima o dopo, nessun campo extra.")
    if mime_type in ("image/jpg", "image/jpeg"):
        mime_type = "image/jpeg"
    if mime_type == "application/pdf":
        images_b64, pdf_mime = pdf_to_images_b64(doc_bytes)
        if not images_b64:
            return None, T("Impossibile convertire il PDF.",
                           "Could not convert the PDF.")
        content = [{"type": "image", "source": {"type": "base64", "media_type": pdf_mime, "data": b}}
                   for b in images_b64[:10]]
        content.append({"type": "text", "text": fields_prompt})
    elif mime_type.startswith("image/"):
        b64 = base64.standard_b64encode(doc_bytes).decode("utf-8")
        content = [{"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": b64}},
                   {"type": "text", "text": fields_prompt}]
    else:
        text = doc_bytes.decode("utf-8", errors="replace")
        content = [{"type": "text", "text": f"Documento:\n{text}\n\n{fields_prompt}"}]
    last_exc = None
    response = None
    for attempt in range(3):
        try:
            # Stream so large max_tokens is allowed (non-streaming has a
            # timeout-based cap that rejects big outputs).
            with client.messages.stream(
                    model="claude-sonnet-4-6", max_tokens=32000,
                    system=system_msg,
                    messages=[{"role": "user", "content": content}]) as stream:
                response = stream.get_final_message()
            break
        except anthropic.AuthenticationError:
            return None, T("Chiave API Anthropic non valida o scaduta. Aggiorna "
                           "ANTHROPIC_API_KEY nei secrets dell'app (o inseriscine una "
                           "valida nella barra laterale). Puoi comunque inserire i dati "
                           "manualmente qui sotto.",
                           "Invalid or expired Anthropic API key. Update "
                           "ANTHROPIC_API_KEY in the app secrets (or enter a valid one "
                           "in the sidebar). You can still enter the data manually below.")
        except anthropic.APIStatusError as e:
            if e.status_code == 529:
                last_exc = e; time.sleep(10 * (attempt + 1))
            else:
                return None, T(f"Errore API Claude ({e.status_code}): {e}",
                               f"Claude API error ({e.status_code}): {e}")
        except anthropic.APIConnectionError as e:
            return None, T(f"Connessione a Claude non riuscita: {e}",
                           f"Could not connect to Claude: {e}")
    else:
        return None, T(f"Claude API sovraccarica: {last_exc}",
                       f"Claude API overloaded: {last_exc}")
    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()
    raw = re.sub(r',"note":"[^}]*"}', '}', raw)
    raw = re.sub(r'"note":"[^}]*",', '', raw)
    raw = re.sub(r',?\s*"note"\s*:\s*null', '', raw)
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        recovered, depth, start = [], 0, None
        for i, ch in enumerate(raw):
            if ch == '{':
                if depth == 0:
                    start = i
                depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0 and start is not None:
                    try:
                        recovered.append(json.loads(raw[start:i + 1]))
                    except Exception:
                        pass
                    start = None
        if recovered:
            return recovered, T(f"⚠️ Risposta troncata: recuperati {len(recovered)} articoli.",
                                f"⚠️ Truncated response: recovered {len(recovered)} items.")
        return None, T(f"Parsing JSON fallito. Risposta: {raw[:300]}",
                       f"JSON parsing failed. Response: {raw[:300]}")
    if isinstance(parsed, dict):
        parsed = [parsed]
    return parsed, None


# ── shared-storage operation wrappers ─────────────────────────────────────────
def with_workbook(fn, read_only=False):
    """Run fn(local_path) inside a locked workbook session."""
    backend = get_backend()
    with storage.workbook_session(backend, owner=current_user(), read_only=read_only) as path:
        stock_agent.EXCEL_FILE = path
        return fn(path)


def flash(message, level="success"):
    """Show a confirmation that stays visible regardless of the active tab."""
    icon = {"success": "✅", "error": "❌", "warning": "⚠️"}.get(level, "ℹ️")
    try:
        st.toast(message, icon=icon)
    except Exception:
        pass
    st.session_state["_flash"] = (level, f"{icon} {message}")


def render_flash():
    """Render and clear any pending flash message at the top of the page."""
    f = st.session_state.pop("_flash", None)
    if not f:
        return
    level, msg = f
    {"success": st.success, "error": st.error, "warning": st.warning}.get(
        level, st.info)(msg)


# ── UI sections ───────────────────────────────────────────────────────────────
def section_dashboard():
    st.subheader(T("📊 Stato magazzino", "📊 Stock status"))
    try:
        s = with_workbook(lambda _p: stock_agent.stato(), read_only=True)
        pp = with_workbook(lambda _p: stock_agent.stato_per_prodotto(), read_only=True)
    except Exception as e:
        st.error(T(f"Impossibile leggere il magazzino: {e}",
                   f"Could not read the stock: {e}")); return
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(T("Totale articoli", "Total items"), s["totale_articoli"])
    c2.metric(T("In giacenza", "In stock"), s["in_giacenza"])
    c3.metric(T("Venduti", "Sold"), s["venduti"])
    c4.metric(T("Max N. operazione", "Max operation no."), s["max_n_operazione"])
    tot = pp["totale"]
    v1, v2, v3 = st.columns(3)
    v1.metric(T("Valore costo in giacenza", "Cost value in stock"),
              f"€ {tot['cost_in_stock']:,.2f}")
    v2.metric(T("Valore vendite (venduti)", "Sales value (sold)"),
              f"€ {tot['sale_value_sold']:,.2f}")
    v3.metric(T("Margine sui venduti", "Margin on sold"),
              f"€ {tot['margin_sold']:,.2f}")
    st.markdown(T("##### Per tipologia di prodotto", "##### By product type"))
    rows = pp["per_prodotto"]
    if not rows:
        st.info(T("Nessun dato.", "No data.")); return
    col_tip = T("Tipologia", "Type")
    col_instock = T("In giacenza", "In stock")
    col_cost_stock = T("Costo giacenza €", "Stock cost €")
    df = pd.DataFrame(rows)[
        ["tipologia", "in_stock", "sold", "cost_in_stock", "cost_sold",
         "sale_value_sold", "margin_sold"]
    ].rename(columns={
        "tipologia": col_tip, "in_stock": col_instock, "sold": T("Venduti", "Sold"),
        "cost_in_stock": col_cost_stock, "cost_sold": T("Costo venduti €", "Sold cost €"),
        "sale_value_sold": T("Ricavo venduti €", "Sold revenue €"),
        "margin_sold": T("Margine €", "Margin €")})
    st.dataframe(df, use_container_width=True, hide_index=True)
    chart = df[df[col_instock] > 0].set_index(col_tip)[col_cost_stock]
    if not chart.empty:
        st.bar_chart(chart)


def section_search():
    st.subheader(T("🔎 Cerca articolo", "🔎 Search item"))
    # --- free-text search: marca / modello / matricola / calibro / tipologia ---
    st.markdown(T("**Ricerca per marca, modello, matricola, calibro…**",
                  "**Search by brand, model, serial, caliber…**"))
    term = st.text_input(T("Cerca", "Search"), key="search_term",
                         placeholder=T("es. Beretta, 686, 12, FUCILE…",
                                       "e.g. Beretta, 686, 12, SHOTGUN…"))
    if st.button(T("Cerca", "Search"), type="primary", key="search_go") and term.strip():
        res = with_workbook(lambda _p: stock_agent.search_items(term), read_only=True)
        results = res.get("results", [])
        if not results:
            st.warning(T("Nessun articolo corrispondente.", "No matching items."))
        else:
            st.success(T(f"{res['count']} articoli trovati", f"{res['count']} items found")
                       + (T(" (mostrati i primi 300)", " (showing first 300)")
                          if res["count"] >= 300 else "") + ".")
            df = pd.DataFrame([{
                T("Stato", "Status"): r["stato"], T("Tipologia", "Type"): r["tipologia"],
                T("Marca", "Brand"): r["marca"], T("Modello", "Model"): r["modello"],
                T("Calibro", "Caliber"): r["calibro"],
                T("Matr. arma", "Weapon s/n"): r["matr_arma"],
                T("Matr. canna", "Barrel s/n"): r["matr_canna"],
                T("N.op", "Op. no."): r["n_operazione"],
                T("Data carico", "Load date"): r["data_carico"],
                T("Fornitore", "Supplier"): r["fornitore"], T("Riga", "Row"): r["row"],
            } for r in results])
            st.dataframe(df, hide_index=True, use_container_width=True)
    # --- exact lookup by matricola / n. operazione (returns the full record) ---
    with st.expander(T("Ricerca esatta per matricola / n. operazione",
                       "Exact lookup by serial / operation no.")):
        col1, col2 = st.columns(2)
        matr = col1.text_input(T("Matricola arma", "Weapon serial"), key="exact_matr")
        n_op = col2.text_input(T("N. operazione", "Operation no."), key="exact_nop")
        if st.button(T("Cerca esatto", "Exact search"), key="exact_go"):
            q = {}
            if matr.strip():
                q["matr_arma"] = matr.strip()
            if n_op.strip():
                q["n_operazione"] = n_op.strip()
            if not q:
                st.warning(T("Inserisci una matricola o un n. operazione.",
                             "Enter a serial or an operation no."))
            else:
                res = with_workbook(lambda _p: stock_agent.cerca_item(q), read_only=True)
                if res.get("status") == "found":
                    st.success(T(f"Trovato alla riga {res['row']}.",
                                 f"Found at row {res['row']}."))
                    st.json({k: v for k, v in res.items() if k not in ("status",)})
                else:
                    st.warning(T("Articolo non trovato.", "Item not found."))
    # --- cancel / delete (mistake correction) ---
    with st.expander(T("🗑️ Annulla scarico / Elimina carico (correzione errori)",
                       "🗑️ Reverse sale / Delete load (error correction)")):
        st.caption(T("Usa questi strumenti solo per correggere un errore di "
                     "inserimento. Le operazioni modificano il registro legale.",
                     "Use these tools only to correct a data-entry mistake. "
                     "These operations modify the legal register."))
        d1, d2 = st.columns(2)
        del_matr = d1.text_input(T("Matricola arma", "Weapon serial"), key="del_matr")
        del_nop = d2.text_input(T("N. operazione", "Operation no."), key="del_nop")

        def _del_query():
            q = {}
            if del_matr.strip():
                q["matr_arma"] = del_matr.strip()
            if del_nop.strip():
                q["n_operazione"] = del_nop.strip()
            return q

        st.markdown(T("**↩️ Annulla uno scarico** — rimette l'articolo in giacenza.",
                      "**↩️ Reverse a sale** — puts the item back in stock."))
        if st.button(T("↩️ Annulla scarico", "↩️ Reverse sale"), key="rev_scarico_go"):
            q = _del_query()
            if not q:
                st.warning(T("Inserisci una matricola o un n. operazione.",
                             "Enter a serial or an operation no."))
            else:
                r = with_workbook(lambda _p: stock_agent.reverse_scarico(q))
                if r.get("status") == "success":
                    st.success(r["message"])
                else:
                    st.error(r.get("message", T("Operazione non riuscita.",
                                                "Operation failed.")))
        st.divider()
        st.markdown(T("**🗑️ Elimina un carico** — cancella l'articolo dal "
                      "magazzino (solo se non ancora venduto).",
                      "**🗑️ Delete a load** — removes the item from the "
                      "stock (only if not yet sold)."))
        if st.button(T("🗑️ Elimina carico", "🗑️ Delete load"), key="del_carico_go"):
            q = _del_query()
            if not q:
                st.warning(T("Inserisci una matricola o un n. operazione.",
                             "Enter a serial or an operation no."))
            else:
                r = with_workbook(lambda _p: stock_agent.delete_carico(q))
                if r.get("status") == "confirm":
                    st.session_state["del_confirm"] = q
                    st.warning(r["message"])
                elif r.get("status") == "success":
                    st.session_state.pop("del_confirm", None)
                    st.success(r["message"])
                else:
                    st.error(r.get("message", T("Operazione non riuscita.",
                                                "Operation failed.")))
        if st.session_state.get("del_confirm"):
            st.error(T("⚠️ Questa eliminazione lascerà un buco nella sequenza "
                       "del registro legale.",
                       "⚠️ This deletion will leave a gap in the legal "
                       "register sequence."))
            cc1, cc2 = st.columns(2)
            if cc1.button(T("✅ Conferma eliminazione", "✅ Confirm deletion"),
                          key="del_confirm_yes"):
                q = st.session_state["del_confirm"]
                r = with_workbook(
                    lambda _p: stock_agent.delete_carico(q, allow_registro_gap=True))
                st.session_state.pop("del_confirm", None)
                if r.get("status") == "success":
                    st.success(r["message"])
                else:
                    st.error(r.get("message", T("Operazione non riuscita.",
                                                "Operation failed.")))
            if cc2.button(T("❌ Annulla", "❌ Cancel"), key="del_confirm_no"):
                st.session_state.pop("del_confirm", None)
                st.info(T("Eliminazione annullata.", "Deletion cancelled."))


def _extract_block(op_key, api_key):
    docs = st.file_uploader(T("DDT / fatture / foto (PDF, JPG, PNG) — multipli",
                              "DDTs / invoices / photos (PDF, JPG, PNG) — multiple"),
                            type=["pdf", "jpg", "jpeg", "png", "webp"],
                            accept_multiple_files=True, key=f"docs_{op_key}")
    if docs and st.button(T("📄 Estrai dati con AI", "📄 Extract data with AI"),
                          key=f"ext_{op_key}", type="primary"):
        if not api_key:
            st.error(T("Nessuna API key configurata. Aggiungi ANTHROPIC_API_KEY nei "
                       "secrets dell'app (o inseriscine una nella barra laterale) "
                       "prima di estrarre i dati.",
                       "No API key configured. Add ANTHROPIC_API_KEY to the app "
                       "secrets (or enter one in the sidebar) before extracting data."))
            return
        all_items, errors = [], []
        bar = st.progress(0.0, text=T("Elaborazione…", "Processing…"))
        for i, f in enumerate(docs):
            bar.progress(i / len(docs), text=T(f"Elaboro {f.name}…", f"Processing {f.name}…"))
            items, err = extract_with_claude(f.read(), f.type or "application/octet-stream",
                                             op_key, api_key)
            if err:
                errors.append(f"{f.name}: {err}")
            if items:
                all_items.extend(items)
        bar.empty()
        for e in errors:
            st.warning(e)
        if all_items:
            st.session_state[f"items_{op_key}"] = all_items
            st.success(T(f"{len(all_items)} articoli estratti.",
                         f"{len(all_items)} items extracted."))
    return st.session_state.get(f"items_{op_key}")


def _serial_range(raw, q):
    """Return q consecutive serials parsed from a range string, else None.

    Accepts forms like '371123-371129', 'Da 371123 A 371129', '371123 / 371129'.
    Only returns serials when exactly two numbers are present AND they span
    exactly q units — otherwise returns None so the caller falls back to N/D.
    Preserves zero-padding width of the starting number.
    """
    import re
    if not raw:
        return None
    nums = re.findall(r"\d+", str(raw))
    if len(nums) != 2:
        return None
    start_s, end_s = nums
    try:
        start, end = int(start_s), int(end_s)
    except ValueError:
        return None
    if end >= start and (end - start + 1) == q:
        width = len(start_s)
        return [str(n).zfill(width) for n in range(start, end + 1)]
    return None


def _expand_quantities(items):
    """Split any line with quantità > 1 into one row per physical unit.

    Firearms compliance needs one row per serial (each row carries its own
    GIACENZA). If the serial is given as a range, assign consecutive serials;
    otherwise emit N/D serials for the operator to fill in before saving.
    Idempotent on already-split data (quantita<=1 passes through untouched).
    """
    out = []
    for it in items or []:
        it = dict(it)
        try:
            q = int(float(it.get("quantita") or 1))
        except (ValueError, TypeError):
            q = 1
        it.pop("quantita", None)
        if q <= 1:
            out.append(it)
            continue
        serials = _serial_range(it.get("matr_arma"), q)
        for k in range(q):
            copy = dict(it)
            copy["matr_arma"] = serials[k] if serials else "N/D"
            out.append(copy)
    return out


def section_carico(api_key):
    st.subheader(T("➕ Carico (acquisto)", "➕ Load (purchase)"))
    cbulk = _bulk_excel("carico", key="carico")
    if cbulk:
        st.session_state["items_CARICO"] = cbulk
        st.success(T(f"{len(cbulk)} righe importate dall'Excel.",
                     f"{len(cbulk)} rows imported from Excel."))
        st.rerun()
    items = _expand_quantities(_extract_block("CARICO", api_key) or [{}])
    rows = []
    for it in items:
        rows.append({
            "data_carico": it.get("data_carico") or it.get("data_ddt") or it.get("data_fattura") or "",
            "n_operazione": 0,
            "tipologia": (it.get("tipologia") or "").upper(),
            "calibro": it.get("calibro") or "", "marca": it.get("marca") or "",
            "modello": it.get("modello") or "", "matr_arma": it.get("matr_arma") or "N/D",
            "matr_canna": it.get("matr_canna") or "N/D", "matr_agg": it.get("matr_agg") or "N/D",
            "fornitore": it.get("fornitore") or "",
            "costo": float(it.get("costo") or 0), "costo_imballo": float(it.get("costo_imballo") or 0),
            "data_ddt": it.get("data_ddt") or "", "data_fattura": it.get("data_fattura") or "",
        })
    st.caption(T("⚠️ Conferma la DATA DI CARICO (arrivo a Gardone) — spesso diversa dalla data del DDT.",
                 "⚠️ Confirm the LOAD DATE (arrival at Gardone) — often different from the DDT date."))
    with st.form("carico_form"):
        edited = st.data_editor(
            pd.DataFrame(rows), num_rows="dynamic", use_container_width=True,
            column_config={
                "n_operazione": st.column_config.NumberColumn(
                    T("N. Op. (0=no registro)", "Op. no. (0=no register)"), step=1),
                "tipologia": st.column_config.SelectboxColumn(
                    T("Tipologia *", "Type *"), options=TIPOLOGIE),
                "costo": st.column_config.NumberColumn(T("Costo €", "Cost €"), format="%.2f"),
                "costo_imballo": st.column_config.NumberColumn(
                    T("Imballo €", "Packaging €"), format="%.2f"),
            })
        submit = st.form_submit_button(
            T("✅ Aggiungi al MAGAZZINO", "✅ Add to STOCK"), type="primary")
    if submit:
        def _do(_p):
            ok, errs = [], []
            for _, r in edited.iterrows():
                d = r.to_dict()
                # data_editor returns NaN (float) for blank cells; turn those
                # into None so downstream `x or ""` / .strip() guards work.
                d = {k: (None if (isinstance(v, float) and pd.isna(v)) else v)
                     for k, v in d.items()}
                # skip fully-empty rows the user may have left in the editor
                if not (str(d.get("tipologia") or "").strip()
                        or str(d.get("matr_arma") or "").strip()):
                    continue
                try:
                    n = int(d.get("n_operazione") or 0)
                except (ValueError, TypeError):
                    n = 0
                d["n_operazione"] = n if n > 0 else None
                res = stock_agent.add_carico(d)
                (ok if res["status"] == "success" else errs).append(res)
            return ok, errs
        try:
            ok, errs = with_workbook(_do)
        except Exception as e:
            flash(T(f"Errore durante il salvataggio del carico: {e}",
                    f"Error while saving the load: {e}"), "error")
            st.rerun()
        if ok:
            st.session_state.pop("items_CARICO", None)
            if errs:
                flash(T(f"{len(ok)} carichi salvati; {len(errs)} non riusciti: ",
                        f"{len(ok)} loads saved; {len(errs)} failed: ")
                      + " | ".join(e["message"] for e in errs), "warning")
            else:
                flash(T(f"{len(ok)} carichi salvati nel registro condiviso.",
                        f"{len(ok)} loads saved to the shared register."))
        elif errs:
            flash(T("Carico non riuscito: ", "Load failed: ")
                  + " | ".join(e["message"] for e in errs), "error")
        else:
            flash(T("Nessun carico salvato — controlla i dati inseriti.",
                    "No load saved — check the entered data."), "warning")
        st.rerun()


def _extract_permit_block(api_key, invoice_no=""):
    """Upload the export permit, AI-extract the ATA number + metadata, archive
    the file in the licence section, and return the ATA number to pre-fill the
    scarico rows. The extracted permit is kept in session_state['scarico_permit']."""
    with st.expander(T("📑 Permesso di esportazione (estrae il N. ATA e archivia la licenza)",
                       "📑 Export permit (extracts the ATA no. and archives the licence)"),
                     expanded=False):
        permit = st.session_state.get("scarico_permit")
        if permit:
            st.success(T(f"Permesso applicato — N. ATA **{permit.get('numero','—')}** "
                         f"· {permit.get('tipo','—')} · {permit.get('paese_destinazione','—')} "
                         f"· archiviato in 📁 Licenze.",
                         f"Permit applied — ATA no. **{permit.get('numero','—')}** "
                         f"· {permit.get('tipo','—')} · {permit.get('paese_destinazione','—')} "
                         f"· archived in 📁 Licences."))
            if st.button(T("🗑️ Rimuovi permesso", "🗑️ Remove permit"), key="permit_clear"):
                st.session_state.pop("scarico_permit", None)
                st.session_state.pop("scarico_ata", None)
                st.rerun()
            return permit.get("numero", "")
        up = st.file_uploader(T("Permesso / licenza export (PDF, JPG, PNG)",
                                "Export permit / licence (PDF, JPG, PNG)"),
                              type=["pdf", "jpg", "jpeg", "png", "webp"],
                              key="permit_up")
        if up and st.button(T("📑 Estrai N. ATA e archivia", "📑 Extract ATA no. and archive"),
                            key="permit_go", type="primary"):
            if not api_key:
                flash(T("Serve la API key (barra laterale) per leggere il permesso.",
                        "An API key (sidebar) is required to read the permit."), "error")
                st.rerun()
            raw = up.read()
            with st.spinner(T("Lettura permesso…", "Reading permit…")):
                got, err = extract_with_claude(raw, up.type or "application/pdf",
                                               "PERMIT", api_key)
            if err and not got:
                flash(T(f"Estrazione permesso non riuscita: {err}",
                        f"Permit extraction failed: {err}"), "error")
                st.rerun()
            p = (got[0] if got else {}) or {}
            # archive the file in the licence section
            try:
                arch = LicenseArchive(get_backend())
                arch.add(raw, up.name, {
                    "doc_type": p.get("tipo") or "Licenza Esportazione",
                    "number": p.get("numero", ""),
                    "country": p.get("paese_destinazione", ""),
                    "counterparty": p.get("cliente", ""),
                    "issue_date": p.get("data_emissione", ""),
                    "expiry_date": p.get("data_scadenza", ""),
                    "invoice_no": invoice_no or "",
                    "tags": ["export", "scarico"],
                    "notes": ("Matricole: " + ", ".join(
                        a.get("matricola", "") for a in p.get("articoli", [])
                        if a.get("matricola"))) if p.get("articoli") else "",
                })
            except Exception as e:
                flash(T(f"N. ATA estratto ma archiviazione licenza non riuscita: {e}",
                        f"ATA no. extracted but licence archiving failed: {e}"),
                      "warning")
            st.session_state["scarico_permit"] = p
            st.session_state["scarico_ata"] = p.get("numero", "")
            flash(T(f"N. ATA {p.get('numero','(non letto)')} applicato e licenza archiviata.",
                    f"ATA no. {p.get('numero','(not read)')} applied and licence archived."))
            st.rerun()
    return st.session_state.get("scarico_ata", "")


def _scan_to_sell():
    """Scanner-driven quick sale: scan a code, confirm the weapon, register."""
    with st.expander(T("📷 Vendita rapida con scanner", "📷 Quick sale with scanner")):
        st.caption(T("Scansiona il codice dell'arma, poi inserisci prezzo e "
                     "cliente e registra la vendita.",
                     "Scan the weapon's code, then enter price and customer and "
                     "register the sale."))
        code = st.text_input(T("Codice scansionato", "Scanned code"),
                             key="sell_scan_code")
        if st.button(T("🔍 Trova arma", "🔍 Find weapon"),
                     key="sell_scan_find") and code.strip():
            def _find(_p):
                barcode_agent.EXCEL_FILE = _p
                return barcode_agent.cmd_scan({"code": code.strip()})
            try:
                st.session_state["sell_scan_rec"] = with_workbook(_find, read_only=True)
            except Exception as e:
                st.session_state["sell_scan_rec"] = {"status": "error", "message": str(e)}
        rec = st.session_state.get("sell_scan_rec")
        if not rec:
            return
        if rec.get("status") != "found":
            st.warning(rec.get("message", T("Non trovato.", "Not found.")))
            return
        if rec.get("venduto"):
            st.warning(T(f"⚠️ Già venduto a {rec.get('CLIENTE')} — niente da fare.",
                         f"⚠️ Already sold to {rec.get('CLIENTE')} — nothing to do."))
            return
        st.success(T(f"{rec.get('MARCA','')} {rec.get('MODELLO','')} · "
                     f"SN {rec.get('MATR_ARMA','')} · riga {rec.get('row')}",
                     f"{rec.get('MARCA','')} {rec.get('MODELLO','')} · "
                     f"SN {rec.get('MATR_ARMA','')} · row {rec.get('row')}"))
        with st.form("scan_sell_form"):
            c1, c2 = st.columns(2)
            price = c1.number_input(T("Prezzo €", "Price €"), min_value=0.0,
                                    step=1.0, format="%.2f")
            client = c2.text_input(T("Cliente", "Customer"))
            c3, c4 = st.columns(2)
            sdate = c3.text_input(T("Data vendita GG/MM/AAAA", "Sale date DD/MM/YYYY"))
            invno = c4.text_input(T("N. fattura", "Invoice no."))
            go = st.form_submit_button(T("✅ Registra vendita", "✅ Register sale"),
                                       type="primary")
        if go:
            d = {"code": code.strip(), "prezzo_vendita": price, "cliente": client,
                 "data_scarico": sdate, "data_fattura_vendita": sdate, "n_fattura": invno}
            try:
                res = with_workbook(lambda _p: stock_agent.add_scarico(d))
            except Exception as e:
                res = {"status": "error", "message": str(e)}
            if res.get("status") == "success":
                st.session_state.pop("sell_scan_rec", None)
                st.session_state.pop("sell_scan_code", None)
                flash(T(f"Vendita registrata (riga {res['row']}).",
                        f"Sale registered (row {res['row']})."))
                st.rerun()
            else:
                st.error(res.get("message"))


def section_scarico(api_key):
    st.subheader(T("➖ Scarico (vendita)", "➖ Unload (sale)"))
    if _BARCODE_OK:
        _scan_to_sell()
    sbulk = _bulk_excel("scarico", key="scarico")
    if sbulk:
        st.session_state["items_SCARICO"] = sbulk
        st.success(T(f"{len(sbulk)} righe importate dall'Excel.",
                     f"{len(sbulk)} rows imported from Excel."))
        st.rerun()
    items = _extract_block("SCARICO", api_key) or [{}]
    invoice_no = next((it.get("n_fattura") for it in items if it.get("n_fattura")), "")
    ata = _extract_permit_block(api_key, invoice_no)
    _render_export_check(items)
    rows = [{
        "matr_arma": it.get("matr_arma") or "", "cliente": it.get("cliente") or "",
        "prezzo_vendita": float(it.get("prezzo_vendita") or 0),
        "data_scarico": it.get("data_fattura_vendita") or it.get("data_scarico") or "",
        "n_fattura": it.get("n_fattura") or "",
        "data_fattura_vendita": it.get("data_fattura_vendita") or "",
        "n_ata": it.get("n_ata") or ata or "", "titolo_acquisto": it.get("titolo_acquisto") or "",
    } for it in items]
    with st.form("scarico_form"):
        edited = st.data_editor(
            pd.DataFrame(rows), num_rows="dynamic", use_container_width=True,
            column_config={"prezzo_vendita": st.column_config.NumberColumn(
                T("Prezzo €", "Price €"), format="%.2f")})
        submit = st.form_submit_button(
            T("✅ Registra scarico", "✅ Register sale"), type="primary")
    if submit:
        def _do(_p):
            ok, errs, fixes = [], [], []
            for _, r in edited.iterrows():
                d = r.to_dict()
                d = {k: (None if (isinstance(v, float) and pd.isna(v)) else v)
                     for k, v in d.items()}
                if not str(d.get("matr_arma") or "").strip():
                    continue
                res = stock_agent.add_scarico(d)
                if res["status"] == "success":
                    ok.append(res)
                else:
                    errs.append(res)
                    # keep only IN-STOCK close matches as one-click fixes
                    instock = [s for s in res.get("suggestions", [])
                               if not s.get("venduto")]
                    if instock:
                        fixes.append({"data": d, "best": instock[0],
                                      "message": res["message"]})
            return ok, errs, fixes
        try:
            ok, errs, fixes = with_workbook(_do)
        except Exception as e:
            flash(T(f"Errore durante il salvataggio dello scarico: {e}",
                    f"Error while saving the sale: {e}"), "error")
            st.rerun()
        st.session_state["scarico_fixes"] = fixes
        if ok:
            st.session_state.pop("items_SCARICO", None)
            if not errs:
                # all done — reset the permit so the next sale starts clean
                st.session_state.pop("scarico_permit", None)
                st.session_state.pop("scarico_ata", None)
            if errs:
                flash(T(f"{len(ok)} scarichi registrati; {len(errs)} non riusciti.",
                        f"{len(ok)} sales registered; {len(errs)} failed."), "warning")
            else:
                flash(T(f"{len(ok)} scarichi registrati nel registro condiviso.",
                        f"{len(ok)} sales registered in the shared register."))
        elif errs:
            flash(T(f"{len(errs)} scarichi non riusciti — vedi sotto.",
                    f"{len(errs)} sales failed — see below."), "error")
        else:
            flash(T("Nessuno scarico registrato — controlla i dati inseriti.",
                    "No sale registered — check the entered data."), "warning")
        st.rerun()
    _render_scarico_fixes()


def _render_export_check(invoice_items):
    """Compare the sales invoice against the export permit (serials, brands,
    calibers) so discrepancies are caught before customs. Anchored on the
    MAGAZZINO record for the authoritative brand/caliber of each serial."""
    permit = st.session_state.get("scarico_permit")
    if not permit:
        return
    permit_items = permit.get("articoli") or []
    st.markdown("---")
    st.markdown(T("#### 🛃 Controllo doganale — fattura vs permesso",
                  "#### 🛃 Customs check — invoice vs permit"))
    if not permit_items:
        st.info(T("Il permesso non elenca le matricole, quindi non posso confrontare "
                  "marche/calibri/matricole. Verifica manualmente che corrispondano.",
                  "The permit does not list serials, so brands/calibers/serials "
                  "cannot be compared. Verify manually that they match."))
        return
    inv_with_serial = [it for it in invoice_items
                       if (it.get("matr_arma") or "").strip()]
    if not inv_with_serial:
        st.info(T("Nessuna matricola sulla fattura da confrontare.",
                  "No serials on the invoice to compare."))
        return
    try:
        rep = with_workbook(
            lambda _p: stock_agent.reconcile_export(
                inv_with_serial, permit_items, stock_agent.load_wb()["MAGAZZINO"]),
            read_only=True)
    except Exception as e:
        st.warning(T(f"Controllo non disponibile: {e}", f"Check unavailable: {e}"))
        return
    total = rep["matched"] + len(rep["only_invoice"]) + len(rep["only_permit"])
    problems = (len(rep["only_invoice"]) + len(rep["only_permit"])
                + len(rep["mismatches"]))
    if rep["ok"]:
        st.success(T(f"✅ Conciliato — tutte le {rep['matched']} armi coincidono "
                     "(matricole, marche e calibri tra fattura e permesso). "
                     "Pronto per l'export.",
                     f"✅ Reconciled — all {rep['matched']} weapons match "
                     "(serials, brands and calibers between invoice and permit). "
                     "Ready for export."))
        return
    st.error(T(f"⚠️ {problems} discrepanze da risolvere prima dell'export "
               f"· {rep['matched']} di {total} armi già conciliate.",
               f"⚠️ {problems} discrepancies to resolve before export "
               f"· {rep['matched']} of {total} weapons already reconciled."))
    st.progress(rep["matched"] / total if total else 0.0,
                text=T(f"{rep['matched']}/{total} conciliate — {problems} da sistemare",
                       f"{rep['matched']}/{total} reconciled — {problems} to fix"))
    # ── allineamento rapido (probabili errori di lettura OCR) ──────────────
    import difflib
    only_inv, only_per = rep["only_invoice"], rep["only_permit"]
    if only_inv and only_per:
        st.markdown(T("**Allineamento rapido matricole** — correggo il permesso "
                      "sulla matricola della fattura/magazzino:",
                      "**Quick serial alignment** — corrects the permit to the "
                      "invoice/stock serial:"))
        inv_serials = [x["matr_arma"] for x in only_inv]
        for pi, pitem in enumerate(only_per):
            ps = pitem["matr_arma"]
            best = max(inv_serials, default=None,
                       key=lambda s: difflib.SequenceMatcher(
                           None, stock_agent._norm_serial(ps),
                           stock_agent._norm_serial(s)).ratio())
            if not best:
                continue
            score = difflib.SequenceMatcher(
                None, stock_agent._norm_serial(ps),
                stock_agent._norm_serial(best)).ratio()
            c1, c2, c3 = st.columns([3, 3, 2])
            c1.markdown(T(f"📜 permesso: **{ps}**", f"📜 permit: **{ps}**"))
            c2.markdown(T(f"📄 fattura: **{best}** ", f"📄 invoice: **{best}** ")
                        + f"<span style='color:gray'>({int(score*100)}%)</span>",
                        unsafe_allow_html=True)
            if score >= 0.55 and c3.button(T("Allinea", "Align"), key=f"algn_{pi}",
                                           type="primary"):
                for art in st.session_state["scarico_permit"]["articoli"]:
                    cur = str(art.get("matricola")
                              or art.get("matr_arma") or "").strip()
                    if cur == ps:
                        if "matricola" in art:
                            art["matricola"] = best
                        else:
                            art["matr_arma"] = best
                        break
                st.rerun()
    # ── modifica manuale del permesso ─────────────────────────────────────
    with st.expander(T("✏️ Modifica manuale del permesso (per far coincidere tutto)",
                       "✏️ Manually edit the permit (to make everything match)")):
        df = pd.DataFrame([
            {T("Matricola", "Serial"): a.get("matricola") or a.get("matr_arma") or "",
             T("Marca", "Brand"): a.get("marca") or a.get("brand") or "",
             T("Modello", "Model"): a.get("modello") or a.get("model") or "",
             T("Calibro", "Caliber"): a.get("calibro") or a.get("cal") or ""}
            for a in permit_items])
        edited = st.data_editor(df, use_container_width=True, hide_index=True,
                                num_rows="dynamic", key="permit_editor")
        if st.button(T("💾 Salva e ricontrolla", "💾 Save and recheck"),
                     key="permit_save", type="primary"):
            new_items = []
            col_s = T("Matricola", "Serial")
            col_b = T("Marca", "Brand")
            col_m = T("Modello", "Model")
            col_c = T("Calibro", "Caliber")
            for _, r in edited.iterrows():
                m = str(r[col_s]).strip()
                if not m:
                    continue
                new_items.append({"matricola": m,
                                  "marca": str(r[col_b]).strip(),
                                  "modello": str(r[col_m]).strip(),
                                  "calibro": str(r[col_c]).strip()})
            st.session_state["scarico_permit"]["articoli"] = new_items
            st.rerun()
    # ── dettaglio discrepanze (riferimento) ───────────────────────────────
    if rep["only_invoice"]:
        st.markdown(T("**Sulla fattura ma NON sul permesso** (non esportabili così):",
                      "**On the invoice but NOT on the permit** (not exportable as-is):"))
        st.dataframe(pd.DataFrame([
            {T("Matricola", "Serial"): x["matr_arma"], T("Marca", "Brand"): x["marca"],
             T("Modello", "Model"): x["modello"], T("Calibro", "Caliber"): x["calibro"]}
            for x in rep["only_invoice"]]), use_container_width=True, hide_index=True)
    if rep["only_permit"]:
        st.markdown(T("**Sul permesso ma NON in fattura** (manca dalla vendita?):",
                      "**On the permit but NOT on the invoice** (missing from the sale?):"))
        st.dataframe(pd.DataFrame([
            {T("Matricola", "Serial"): x["matr_arma"], T("Marca", "Brand"): x["marca"],
             T("Modello", "Model"): x["modello"], T("Calibro", "Caliber"): x["calibro"]}
            for x in rep["only_permit"]]), use_container_width=True, hide_index=True)
    if rep["mismatches"]:
        st.markdown(T("**Dati che non corrispondono** (fattura/magazzino ≠ permesso):",
                      "**Data that does not match** (invoice/stock ≠ permit):"))
        st.dataframe(pd.DataFrame([
            {T("Matricola", "Serial"): m["matricola"], T("Campo", "Field"): m["campo"],
             T("Fattura/Magazzino", "Invoice/Stock"): m["fattura"],
             T("Permesso", "Permit"): m["permesso"],
             T("Nota", "Note"): m.get("nota", "")}
            for m in rep["mismatches"]]), use_container_width=True, hide_index=True)


def _render_scarico_fixes():
    """Show not-found scarichi that have a close in-stock serial, with a button
    to apply the correction and register the sale automatically."""
    fixes = st.session_state.get("scarico_fixes") or []
    if not fixes:
        return
    st.markdown("---")
    st.markdown(T("#### 🔧 Matricole non trovate — correzioni suggerite",
                  "#### 🔧 Serials not found — suggested corrections"))
    st.caption(T("Probabili errori di lettura dal PDF. Controlla e premi **Applica** "
                 "per registrare la vendita con la matricola corretta.",
                 "Likely OCR misreads from the PDF. Check and press **Apply** "
                 "to register the sale with the correct serial."))
    for i, fx in enumerate(fixes):
        d, best = fx["data"], fx["best"]
        orig = str(d.get("matr_arma") or "").strip()
        sim = int(round(best["score"] * 100))
        c1, c2, c3 = st.columns([3, 3, 2])
        c1.markdown(T(f"📄 letta: **{orig or '(vuota)'}**", f"📄 read: **{orig or '(empty)'}**")
                    + (T(f" · cliente {d.get('cliente')}", f" · customer {d.get('cliente')}")
                       if d.get("cliente") else ""))
        c2.markdown(T(f"✅ suggerita: **{best['matricola']}** ",
                      f"✅ suggested: **{best['matricola']}** ")
                    + T(f"<span style='color:gray'>({sim}% · riga {best['riga']})</span>",
                        f"<span style='color:gray'>({sim}% · row {best['riga']})</span>"),
                    unsafe_allow_html=True)
        if c3.button(T("Applica", "Apply"), key=f"fix_scar_{i}", type="primary"):
            d2 = dict(d); d2["matr_arma"] = best["matricola"]
            try:
                res = with_workbook(lambda _p: stock_agent.add_scarico(d2))
            except Exception as e:
                flash(T(f"Errore: {e}", f"Error: {e}"), "error"); st.rerun()
            if res["status"] == "success":
                flash(T(f"Scarico registrato con matricola {best['matricola']} "
                        f"(riga {res['row']}).",
                        f"Sale registered with serial {best['matricola']} "
                        f"(row {res['row']})."))
                lst = st.session_state.get("scarico_fixes") or []
                st.session_state["scarico_fixes"] = [
                    x for j, x in enumerate(lst) if j != i]
            else:
                flash(res["message"], "error")
            st.rerun()
    if st.button(T("Ignora tutti", "Ignore all"), key="fix_scar_clear"):
        st.session_state["scarico_fixes"] = []
        st.rerun()


def section_registro():
    st.subheader(T("📑 Esporta REGISTRO (per la P.S.)",
                   "📑 Export REGISTER (for the police)"))
    st.caption(T("Genera il registro aggiornato pronto da stampare, ricostruito dai dati attuali.",
                 "Generates the up-to-date register ready to print, rebuilt from the current data."))
    if st.button(T("Genera registro", "Generate register"), type="primary"):
        import tempfile
        xlsx_path = os.path.join(tempfile.gettempdir(), "REGISTRO_export.xlsx")
        pdf_path = os.path.join(tempfile.gettempdir(), "REGISTRO_export.pdf")
        res = with_workbook(
            lambda _p: stock_agent.export_registro(xlsx_path, pdf_path), read_only=True)
        st.success(T(f"Registro generato: {res['righe']} operazioni.",
                     f"Register generated: {res['righe']} operations."))
        with open(xlsx_path, "rb") as f:
            st.download_button(T("⬇️ Scarica REGISTRO (Excel)", "⬇️ Download REGISTER (Excel)"),
                               f.read(),
                               file_name=f"REGISTRO {datetime.now():%Y-%m-%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if res.get("pdf") and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                st.download_button(T("⬇️ Scarica REGISTRO (PDF stampa)",
                                     "⬇️ Download REGISTER (print PDF)"), f.read(),
                                   file_name=f"REGISTRO {datetime.now():%Y-%m-%d}.pdf",
                                   mime="application/pdf")
        else:
            st.info(T("PDF non generato (reportlab non installato) — usa l'Excel.",
                      "PDF not generated (reportlab not installed) — use the Excel."))
    st.divider()
    st.markdown(T("**📥 Scarica il file MAGAZZINO completo (tutti i fogli, aggiornato)**",
                  "**📥 Download the full STOCK file (all sheets, up to date)**"))
    st.caption(T("Scarica l'intero workbook così com'è nello storage condiviso — "
                 "ISTRUZIONI, INVENTARIO, MAGAZZINO, REGISTRO, BASE, PRINT.",
                 "Download the entire workbook as it is in shared storage — "
                 "ISTRUZIONI, INVENTARIO, MAGAZZINO, REGISTRO, BASE, PRINT."))
    if st.button(T("Prepara download MAGAZZINO completo", "Prepare full STOCK download"),
                 key="dl_full_wb"):
        try:
            data = get_backend().read_bytes(storage.workbook_name())
            st.download_button(
                T("⬇️ Scarica MAGAZZINO completo (.xlsx)", "⬇️ Download full STOCK (.xlsx)"),
                data,
                file_name=f"{os.path.splitext(storage.workbook_name())[0]} "
                          f"{datetime.now():%Y-%m-%d %H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_full_wb_btn")
        except Exception as e:
            st.error(T(f"Impossibile scaricare il workbook: {e}",
                       f"Could not download the workbook: {e}"))


def section_licenses(api_key=""):
    st.subheader(T("📁 Archivio licenze (Export / Import / EUC)",
                   "📁 Licence archive (Export / Import / EUC)"))
    arch = LicenseArchive(get_backend())
    exp = arch.expiring_soon(days=60)
    if exp:
        st.warning(T("In scadenza / scaduti: ", "Expiring / expired: ") +
                   ", ".join(f"{e['doc_type']} {e.get('number','')} ({e['days_left']}g)" for e in exp[:8]))
    with st.expander(T("➕ Carica nuova licenza / EUC", "➕ Upload new licence / EUC"),
                     expanded=True):
        up = st.file_uploader(T("File (PDF/immagine/qualsiasi)", "File (PDF/image/any)"),
                              key="lic_up")
        # ── AI auto-extraction: read the uploaded licence and pre-fill the form ──
        ex = st.session_state.get("lic_extract", {})
        if up is not None:
            if st.button(T("🔍 Leggi e compila automaticamente",
                           "🔍 Read & auto-fill"), key="lic_extract_go",
                         type="primary"):
                if not api_key:
                    st.error(T("Serve la API key (barra laterale) per leggere il documento.",
                               "An API key (sidebar) is required to read the document."))
                else:
                    raw = up.getvalue()
                    with st.spinner(T("Lettura documento…", "Reading document…")):
                        got, err = extract_with_claude(raw, up.type or "application/pdf",
                                                       "PERMIT", api_key)
                    if err and not got:
                        st.error(T(f"Estrazione non riuscita: {err}",
                                   f"Extraction failed: {err}"))
                    else:
                        p = (got[0] if got else {}) or {}
                        matricole = ", ".join(a.get("matricola", "")
                                              for a in p.get("articoli", [])
                                              if a.get("matricola"))
                        st.session_state["lic_extract"] = {
                            "doc_type": p.get("tipo") or "",
                            "number": p.get("numero") or "",
                            "country": p.get("paese_destinazione") or "",
                            "counterparty": p.get("cliente") or "",
                            "issue_date": p.get("data_emissione") or "",
                            "expiry_date": p.get("data_scadenza") or "",
                            "notes": (T("Matricole: ", "Serials: ") + matricole) if matricole else "",
                        }
                        st.rerun()
            if ex:
                st.success(T("Dati estratti dal documento — controlla e salva.",
                             "Data extracted from the document — review and save."))
        # ── form (pre-filled from AI extraction when available) ──
        c1, c2 = st.columns(2)
        _dt = ex.get("doc_type", "")
        doc_type = c1.selectbox(T("Tipo documento", "Document type"), DOC_TYPES,
                                index=DOC_TYPES.index(_dt) if _dt in DOC_TYPES else 0)
        number = c2.text_input(T("Numero", "Number"), value=ex.get("number", ""))
        c3, c4 = st.columns(2)
        country = c3.text_input(T("Paese", "Country"), value=ex.get("country", ""))
        counterparty = c4.text_input(T("Controparte / Cliente", "Counterparty / Customer"),
                                     value=ex.get("counterparty", ""))
        c5, c6 = st.columns(2)
        issue = c5.text_input(T("Data emissione (GG/MM/AAAA)", "Issue date (DD/MM/YYYY)"),
                              value=ex.get("issue_date", ""))
        expiry = c6.text_input(T("Data scadenza (GG/MM/AAAA)", "Expiry date (DD/MM/YYYY)"),
                               value=ex.get("expiry_date", ""))
        c7, c8 = st.columns(2)
        n_op = c7.text_input(T("N. operazione collegata", "Linked operation no."))
        inv = c8.text_input(T("N. fattura collegata", "Linked invoice no."))
        tags = st.text_input(T("Tag (separati da virgola)", "Tags (comma-separated)"))
        notes = st.text_area(T("Note", "Notes"), value=ex.get("notes", ""))
        if st.button(T("Salva in archivio", "Save to archive"), type="primary",
                     key="lic_save"):
            if not up:
                st.error(T("Seleziona un file.", "Select a file."))
            else:
                arch.add(up.getvalue(), up.name, {
                    "doc_type": doc_type, "number": number, "country": country,
                    "counterparty": counterparty, "issue_date": issue, "expiry_date": expiry,
                    "n_operazione": n_op, "invoice_no": inv,
                    "tags": [t.strip() for t in tags.split(",") if t.strip()], "notes": notes})
                st.session_state.pop("lic_extract", None)
                st.success(T("Licenza archiviata.", "Licence archived."))
                st.rerun()
    st.markdown(T("##### Archivio", "##### Archive"))
    fc1, fc2 = st.columns([2, 1])
    query = fc1.text_input(T("Cerca (numero, paese, controparte, n.op, tag…)",
                             "Search (number, country, counterparty, op.no, tag…)"))
    all_label = T("(tutti)", "(all)")
    ftype = fc2.selectbox(T("Filtra tipo", "Filter type"), [all_label] + DOC_TYPES)
    entries = arch.list(query=query or None, doc_type=None if ftype == all_label else ftype)
    if not entries:
        st.info(T("Nessuna licenza in archivio.", "No licences in the archive."))
        return
    for e in entries:
        with st.container(border=True):
            cols = st.columns([3, 2, 2, 2, 1])
            cols[0].markdown(f"**{e['doc_type']}** · {e.get('number','—')}\n\n{e['filename']}")
            cols[1].write(T(f"🌍 {e.get('country','—')}\n\n{e.get('counterparty','—')}",
                            f"🌍 {e.get('country','—')}\n\n{e.get('counterparty','—')}"))
            cols[2].write(T(f"Emessa: {e.get('issue_date','—')}\n\nScad.: {e.get('expiry_date','—')}",
                            f"Issued: {e.get('issue_date','—')}\n\nExp.: {e.get('expiry_date','—')}"))
            cols[3].write(T(f"N.Op: {e.get('n_operazione','—')}\n\nFatt: {e.get('invoice_no','—')}",
                            f"Op.no: {e.get('n_operazione','—')}\n\nInv: {e.get('invoice_no','—')}"))
            data, fname = arch.get_file(e["id"])
            cols[4].download_button("⬇️", data, file_name=fname, key=f"dl_{e['id']}")
            if e.get("tags"):
                st.caption("🏷️ " + ", ".join(e["tags"]))
            if st.button(T("Elimina", "Delete"), key=f"del_{e['id']}"):
                arch.delete(e["id"]); st.rerun()


def _storage_error_message(e):
    """Turn a raw storage/botocore exception into an actionable message.
    Streamlit redacts uncaught errors, so we read the code ourselves and show it."""
    try:
        from botocore.exceptions import ClientError, EndpointConnectionError, \
            NoCredentialsError
    except Exception:
        ClientError = EndpointConnectionError = NoCredentialsError = ()
    if NoCredentialsError and isinstance(e, NoCredentialsError):
        return T("Credenziali storage mancanti. Controlla AWS_ACCESS_KEY_ID e "
                 "AWS_SECRET_ACCESS_KEY nei secret dell'app.",
                 "Storage credentials missing. Check AWS_ACCESS_KEY_ID and "
                 "AWS_SECRET_ACCESS_KEY in the app secrets.")
    if EndpointConnectionError and isinstance(e, EndpointConnectionError):
        return T(f"Impossibile raggiungere lo storage. Controlla BCW_S3_ENDPOINT "
                 f"(endpoint R2/S3). Dettaglio: {e}",
                 f"Cannot reach the storage. Check BCW_S3_ENDPOINT "
                 f"(R2/S3 endpoint). Detail: {e}")
    if ClientError and isinstance(e, ClientError):
        err = (getattr(e, "response", {}) or {}).get("Error", {})
        code = err.get("Code", "?")
        msg = err.get("Message", "")
        hints = {
            "InvalidAccessKeyId": T("La access key non è valida — ricontrolla AWS_ACCESS_KEY_ID.",
                                    "The access key is invalid — recheck AWS_ACCESS_KEY_ID."),
            "SignatureDoesNotMatch": T("La secret key non corrisponde — ricontrolla AWS_SECRET_ACCESS_KEY.",
                                       "The secret key does not match — recheck AWS_SECRET_ACCESS_KEY."),
            "AccessDenied": T("Accesso negato — il token R2 non ha permessi di scrittura sul bucket.",
                              "Access denied — the R2 token lacks write permission on the bucket."),
            "NoSuchBucket": T("Il bucket non esiste — controlla BCW_S3_BUCKET.",
                              "The bucket does not exist — check BCW_S3_BUCKET."),
        }
        hint = hints.get(code, "")
        return (T(f"Errore storage [{code}]: {msg} {hint}",
                  f"Storage error [{code}]: {msg} {hint}")).strip()
    return T(f"Errore storage: {e}", f"Storage error: {e}")


def section_settings(backend):
    st.subheader(T("⚙️ Impostazioni", "⚙️ Settings"))
    st.write(T(f"**Backend storage:** `{backend.name}`", f"**Storage backend:** `{backend.name}`"))
    st.write(T(f"**Workbook:** `{storage.workbook_name()}`",
               f"**Workbook:** `{storage.workbook_name()}`"))
    st.write(T(f"**Utente corrente:** `{current_user()}` "
               "_(modificabile nella barra laterale)_",
               f"**Current user:** `{current_user()}` "
               "_(editable in the sidebar)_"))
    if not backend_ready(backend):
        st.warning(T("Il workbook non è ancora presente nello storage condiviso.",
                     "The workbook is not yet present in shared storage."))
        seed = st.file_uploader(T("Carica il MAGAZZINO iniziale per inizializzare",
                                  "Upload the initial STOCK file to initialize"),
                                type=["xlsx"])
        if seed and st.button(T("Inizializza storage", "Initialize storage")):
            try:
                backend.write_bytes(storage.workbook_name(), seed.read())
            except Exception as e:
                st.error(_storage_error_message(e))
            else:
                st.success(T("Workbook caricato nello storage condiviso.",
                             "Workbook uploaded to shared storage.")); st.rerun()
    else:
        st.success(T("Workbook presente nello storage condiviso.",
                     "Workbook present in shared storage."))
        if st.button(T("Verifica struttura workbook", "Verify workbook structure")):
            res = with_workbook(lambda _p: stock_agent.validate_workbook(), read_only=True)
            st.json(res)

        st.divider()
        with st.expander(T("⚠️ Sostituisci il workbook (zona pericolosa)",
                           "⚠️ Replace the workbook (danger zone)")):
            st.caption(T("Carica un nuovo file MAGAZZINO per sostituire quello attuale nello "
                         "storage condiviso. L'operazione SOVRASCRIVE il file live e non è "
                         "reversibile — scarica prima un backup qui sotto.",
                         "Upload a new STOCK file to replace the current one in shared storage. "
                         "This OVERWRITES the live file and cannot be undone — download a "
                         "backup first below."))
            # 1) backup current live workbook
            if st.button(T("⬇️ Prepara backup del workbook attuale",
                           "⬇️ Prepare backup of current workbook"), key="repl_backup"):
                try:
                    cur = get_backend().read_bytes(storage.workbook_name())
                except Exception as e:
                    st.error(_storage_error_message(e))
                else:
                    st.download_button(
                        T("⬇️ Scarica backup (.xlsx)", "⬇️ Download backup (.xlsx)"),
                        cur,
                        file_name=f"{os.path.splitext(storage.workbook_name())[0]} "
                                  f"BACKUP {datetime.now():%Y-%m-%d %H%M}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument."
                             "spreadsheetml.sheet",
                        key="repl_backup_btn")
            # 2) upload replacement — validated before it can go live
            newwb = st.file_uploader(T("Carica il nuovo workbook (.xlsx)",
                                       "Upload the new workbook (.xlsx)"),
                                     type=["xlsx"], key="repl_upload")
            if newwb is not None:
                import tempfile
                data = newwb.getvalue()
                _prev = getattr(stock_agent, "EXCEL_FILE", None)
                tmp = os.path.join(tempfile.gettempdir(), "_replace_check.xlsx")
                try:
                    with open(tmp, "wb") as fh:
                        fh.write(data)
                    stock_agent.EXCEL_FILE = tmp
                    check = stock_agent.validate_workbook()
                finally:
                    stock_agent.EXCEL_FILE = _prev
                ok = isinstance(check, dict) and check.get("status") == "ok"
                if ok:
                    st.success(T("✅ Struttura valida — pronto a sostituire.",
                                 "✅ Structure valid — ready to replace."))
                else:
                    st.error(T("❌ Struttura non valida — sostituzione bloccata.",
                               "❌ Invalid structure — replacement blocked."))
                st.json(check)
                confirm = st.checkbox(
                    T("Ho capito: questo sovrascrive il file live nello storage.",
                      "I understand: this overwrites the live file in storage."),
                    key="repl_confirm")
                if st.button(T("🔁 Sostituisci ora", "🔁 Replace now"), type="primary",
                             disabled=not (ok and confirm), key="repl_go"):
                    try:
                        get_backend().write_bytes(storage.workbook_name(), data)
                    except Exception as e:
                        st.error(_storage_error_message(e))
                    else:
                        st.success(T("Workbook sostituito nello storage condiviso.",
                                     "Workbook replaced in shared storage.")); st.rerun()


# ── Documenti (proforma / packing / declarations) ─────────────────────────────
def _bulk_excel(kind, company=None, key=None):
    """Template-download + Excel-upload widget for a bulk flow."""
    key = key or kind
    suffix = f"_{company}" if company else ""
    with st.expander(T("⬆️ Carica in blocco da Excel", "⬆️ Bulk upload from Excel")):
        st.caption(T("Scarica il modello, compilalo, poi ricaricalo per inserire molte righe insieme.",
                     "Download the template, fill it in, then re-upload it to add many rows at once."))
        st.download_button(
            T("⬇️ Scarica modello Excel", "⬇️ Download Excel template"),
            bulk_io.template_bytes(kind, company),
            file_name=f"modello_{kind}{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"tpl_{key}")
        up = st.file_uploader(T("Carica il file compilato", "Upload the filled-in file"),
                              type=["xlsx"], key=f"up_{key}")
        if up is not None and st.button(T("📥 Importa righe", "📥 Import rows"), key=f"imp_{key}"):
            try:
                return bulk_io.parse(kind, up.read(), company)
            except Exception as e:
                st.error(T(f"Errore nella lettura dell'Excel: {e}",
                           f"Error reading the Excel file: {e}"))
    return None


def _archive_doc(base_name, html, meta):
    """Save a generated document (PDF if available + HTML) to the shared archive
    and record it in the ledger. Never blocks document generation on failure."""
    try:
        pdf_bytes = None
        try:
            pdf_bytes = doc_templates.to_pdf(html)
        except Exception:
            pdf_bytes = None
        DocArchive(get_backend()).add(base_name, html, pdf_bytes, meta)
    except Exception as e:
        st.warning(T(f"Documento generato, ma non archiviato ({e}).",
                     f"Document generated, but not archived ({e})."))


def _doc_downloads(full_html, base_name):
    """Render HTML preview + PDF/Word download buttons for a generated document."""
    import streamlit.components.v1 as components
    st.markdown(T("##### Anteprima", "##### Preview"))
    components.html(full_html, height=560, scrolling=True)
    c1, c2, c3 = st.columns(3)
    try:
        pdf_bytes = doc_templates.to_pdf(full_html)
        c1.download_button("⬇️ PDF", pdf_bytes, file_name=f"{base_name}.pdf",
                           mime="application/pdf", type="primary")
    except Exception as e:
        c1.warning(T(f"PDF non disponibile ({e}). Usa Word o stampa l'anteprima.",
                     f"PDF unavailable ({e}). Use Word or print the preview."))
    doc_bytes = doc_templates.to_doc(full_html)
    c2.download_button("⬇️ Word", doc_bytes, file_name=f"{base_name}.doc",
                       mime="application/msword")
    c3.download_button("⬇️ HTML", full_html.encode("utf-8"),
                       file_name=f"{base_name}.html", mime="text/html")


def _buyer_inputs(prefix):
    st.markdown(T("**Destinatario / Buyer**", "**Recipient / Buyer**"))
    b1, b2 = st.columns(2)
    name = b1.text_input(T("Nome / Name", "Name"), key=f"{prefix}_b_name")
    trade = b2.text_input(T("Ragione sociale / Trade", "Company name / Trade"),
                          key=f"{prefix}_b_trade")
    addr = st.text_input(T("Indirizzo / Address", "Address"), key=f"{prefix}_b_addr")
    b3, b4 = st.columns(2)
    country = b3.text_input(T("Paese / Country", "Country"), key=f"{prefix}_b_country")
    phone = b4.text_input(T("Telefono / Phone", "Phone"), key=f"{prefix}_b_phone")
    return {"name": name, "trade": trade, "addr": addr, "country": country, "phone": phone}


def _stock_picker(target_key, map_fn, label=None):
    """BCW-only: pick in-stock guns and append mapped rows to st.session_state[target_key]."""
    if label is None:
        label = T("➕ Aggiungi da magazzino (in giacenza)",
                  "➕ Add from stock (in stock)")
    with st.expander(label):
        q = st.text_input(T("Filtra (matricola, marca, modello, calibro, tipologia)",
                            "Filter (serial, brand, model, caliber, type)"),
                          key=f"{target_key}_pick_q")
        try:
            stock = with_workbook(lambda _p: stock_agent.list_in_stock(q or None),
                                  read_only=True)
        except Exception as e:
            st.error(T(f"Impossibile leggere il magazzino: {e}",
                       f"Could not read the stock: {e}"))
            return
        if not stock:
            st.info(T("Nessun articolo in giacenza corrispondente.",
                      "No matching items in stock."))
            return
        st.caption(T(f"{len(stock)} articoli in giacenza.", f"{len(stock)} items in stock."))
        col_sel = T("Sel", "Sel")
        df = pd.DataFrame([{
            col_sel: False, T("Tipologia", "Type"): s["tipologia"],
            T("Marca", "Brand"): s["marca"], T("Modello", "Model"): s["modello"],
            T("Calibro", "Caliber"): s["calibro"],
            T("Matr. arma", "Weapon s/n"): s["matr_arma"],
            T("Matr. canna", "Barrel s/n"): s["matr_canna"],
            T("Costo €", "Cost €"): s["costo_compl"],
        } for s in stock])
        edited = st.data_editor(df, hide_index=True, use_container_width=True,
                                key=f"{target_key}_pick_tbl",
                                column_config={col_sel: st.column_config.CheckboxColumn(col_sel)})
        if st.button(T("Aggiungi selezionati", "Add selected"), key=f"{target_key}_pick_add"):
            chosen = [stock[i] for i in edited.index[edited[col_sel]].tolist()]
            if not chosen:
                st.warning(T("Nessuna riga selezionata.", "No rows selected."))
            else:
                cur = list(st.session_state.get(target_key, []))
                cur.extend(map_fn(s) for s in chosen)
                st.session_state[target_key] = cur
                st.success(T(f"{len(chosen)} articoli aggiunti.", f"{len(chosen)} items added."))
                st.rerun()


def _doc_proforma():
    company = st.radio(T("Azienda", "Company"), ["BME", "BCW"], horizontal=True, key="pf_company")
    co = doc_templates.COMPANIES[company]
    backend = get_backend()
    st.caption(T(f"Prossimo numero: **{doc_numbering.preview_next(backend, company, 'proforma')}** "
                 f"· Valuta: {co['currency']}",
                 f"Next number: **{doc_numbering.preview_next(backend, company, 'proforma')}** "
                 f"· Currency: {co['currency']}"))
    date = st.date_input(T("Data", "Date"), key="pf_date")
    buyer = _buyer_inputs("pf")
    items_key = f"pf_items_{company}"
    if items_key not in st.session_state:
        st.session_state[items_key] = []
    bulk = _bulk_excel("proforma", company, key=f"pf_{company}")
    if bulk:
        st.session_state[items_key] = bulk
        st.success(T(f"{len(bulk)} righe importate dall'Excel.",
                     f"{len(bulk)} rows imported from Excel."))
        st.rerun()
    if company == "BCW":
        _stock_picker(items_key, lambda s: {
            "tipo": s["tipologia"], "cal": s["calibro"], "marca": s["marca"],
            "desc": s["modello"], "serial": s["matr_arma"],
            "qty": s["quantita"], "price": s["costo_compl"]})
        cols = ["tipo", "cal", "marca", "desc", "serial", "qty", "price"]
    else:
        cols = ["type", "gauge", "brand", "serial", "qty", "price"]
    base = st.session_state[items_key] if st.session_state[items_key] else [
        {c: ("" if c not in ("qty", "price") else (1 if c == "qty" else 0.0)) for c in cols}]
    df = pd.DataFrame(base, columns=cols)
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="pf_editor",
                            column_config={
                                "qty": st.column_config.NumberColumn("qty", step=1),
                                "price": st.column_config.NumberColumn("price", format="%.2f")})
    notes = st.text_area(T("Note", "Notes"), key="pf_notes")
    if st.button(T("📄 Genera proforma", "📄 Generate proforma"), type="primary", key="pf_gen"):
        items = [r for r in edited.to_dict("records")
                 if any(str(v).strip() for v in r.values())]
        number = doc_numbering.next_number(backend, company, "proforma", owner=current_user())
        f = {"company": company, "number": number, "date": date.isoformat(),
             "currency": co["currency"], "buyer": buyer, "items": items,
             "notes": notes or ""}
        html = doc_templates.document_html(doc_templates.render_proforma(f), number)
        _archive_doc(number, html, {
            "number": number, "company": company, "doc_type": "proforma",
            "buyer": buyer.get("name") or buyer.get("trade") or "",
            "total": doc_templates._doc_total(f), "currency": co["currency"],
            "doc_date": date.isoformat(), "user": current_user()})
        st.session_state["pf_result"] = (html, number)
        st.session_state[items_key] = []
        st.success(T(f"Proforma {number} generata e archiviata.",
                     f"Proforma {number} generated and archived."))
    if st.session_state.get("pf_result"):
        html, number = st.session_state["pf_result"]
        _doc_downloads(html, number)


def _doc_packing():
    company = st.radio(T("Azienda", "Company"), ["BME", "BCW"], horizontal=True, key="pk_company")
    backend = get_backend()
    st.caption(T(f"Prossimo numero: **{doc_numbering.preview_next(backend, company, 'packing')}**",
                 f"Next number: **{doc_numbering.preview_next(backend, company, 'packing')}**"))
    date = st.date_input(T("Data", "Date"), key="pk_date")
    buyer = _buyer_inputs("pk")
    pbulk = _bulk_excel("packing", company, key=f"pk_{company}")
    if pbulk:
        st.session_state["pk_nparcels"] = max(1, pbulk["n_parcels"])
        for i, parcel in enumerate(pbulk["parcels"]):
            st.session_state[f"pk_items_{company}_{i}"] = parcel["items"] or [{}]
            st.session_state[f"pk_dims_{i}"] = parcel["dims"]
            st.session_state[f"pk_weight_{i}"] = parcel["weight"]
        st.success(T(f"{pbulk['n_parcels']} colli importati dall'Excel.",
                     f"{pbulk['n_parcels']} parcels imported from Excel."))
        st.rerun()
    if "pk_nparcels" not in st.session_state:
        st.session_state["pk_nparcels"] = 1
    n_parcels = st.number_input(T("Numero di colli / parcels", "Number of parcels"),
                                min_value=1, max_value=50, step=1, key="pk_nparcels")
    pcols = ["qty", "type", "brand", "model", "caliber", "serial1", "serial2"]
    parcels_input = []
    for i in range(int(n_parcels)):
        st.markdown(T(f"**Collo / Parcel {i + 1}**", f"**Parcel {i + 1}**"))
        pk_key = f"pk_items_{company}_{i}"
        if company == "BCW":
            _stock_picker(pk_key, lambda s: {
                "qty": s["quantita"], "type": s["tipologia"], "brand": s["marca"],
                "model": s["modello"], "caliber": s["calibro"],
                "serial1": s["matr_arma"], "serial2": s["matr_canna"]},
                label=T(f"➕ Aggiungi da magazzino → collo {i + 1}",
                        f"➕ Add from stock → parcel {i + 1}"))
        base = st.session_state.get(pk_key) or [{c: ("" if c != "qty" else 1) for c in pcols}]
        df = pd.DataFrame(base, columns=pcols)
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True,
                                key=f"pk_editor_{i}",
                                column_config={"qty": st.column_config.NumberColumn("qty", step=1)})
        d1, d2 = st.columns(2)
        dims = d1.text_input(T("Dimensioni pacco (es. 60x40x30 CM)",
                               "Parcel dimensions (e.g. 60x40x30 CM)"), key=f"pk_dims_{i}")
        weight = d2.text_input(T("Peso pacco (es. 20,5 KG)", "Parcel weight (e.g. 20.5 KG)"),
                               key=f"pk_weight_{i}")
        items = [r for r in edited.to_dict("records")
                 if any(str(v).strip() for v in r.values())]
        parcels_input.append({"items": items, "dims": dims, "weight": weight})
    notes = st.text_area(T("Note", "Notes"), key="pk_notes")
    if st.button(T("📦 Genera packing list", "📦 Generate packing list"),
                 type="primary", key="pk_gen"):
        number = doc_numbering.next_number(backend, company, "packing", owner=current_user())
        f = {"company": company, "number": number, "date": date.isoformat(),
             "buyer": buyer, "parcels": parcels_input, "notes": notes or ""}
        html = doc_templates.document_html(doc_templates.render_packing(f), number)
        _archive_doc(number, html, {
            "number": number, "company": company, "doc_type": "packing",
            "buyer": buyer.get("name") or buyer.get("trade") or "",
            "total": "", "currency": "",
            "doc_date": date.isoformat(), "user": current_user()})
        st.session_state["pk_result"] = (html, number)
        for i in range(int(n_parcels)):
            st.session_state.pop(f"pk_items_{company}_{i}", None)
        st.success(T(f"Packing list {number} generata e archiviata.",
                     f"Packing list {number} generated and archived."))
    if st.session_state.get("pk_result"):
        html, number = st.session_state["pk_result"]
        _doc_downloads(html, number)


def _doc_declarations():
    st.caption(T("EUR1 + End User Certificate (BCW). Emessi sulla fattura finale — "
                 "non consumano un numero di proforma.",
                 "EUR1 + End User Certificate (BCW). Issued on the final invoice — "
                 "they do not consume a proforma number."))
    buyer = _buyer_inputs("ef")
    st.markdown(T("**Dettagli dichiarazione** — i dati BCW sono già precompilati (modificabili).",
                  "**Declaration details** — BCW data is pre-filled (editable)."))
    BCW_SEAT = "Via Matteotti 311 – Gardone Val Trompia, 25063 (BS) – Italia"
    c1, c2 = st.columns(2)
    rep = c1.text_input(T("Legale rappresentante (firma)", "Legal representative (signature)"),
                        value="Antoine Abi Saab", key="ef_rep")
    sender = c2.text_input(T("Ditta mittente", "Sending company"),
                           value=doc_templates.COMPANIES["BCW"]["name"], key="ef_sender")
    seat = st.text_input(T("Sede ditta", "Company seat"), value=BCW_SEAT, key="ef_seat")
    c3, c4 = st.columns(2)
    inv_no = c3.text_input(T("N. fattura esportazione", "Export invoice no."), key="ef_invno")
    inv_date = c4.date_input(T("Data fattura", "Invoice date"), key="ef_invdate")
    c5, c6 = st.columns(2)
    dest = c5.text_input(T("Paese di destinazione", "Destination country"), key="ef_dest")
    contract = c6.text_input(T("N. contratto (opz.)", "Contract no. (opt.)"), key="ef_contract")
    commodity = st.text_area(T("Descrizione dettagliata merce", "Detailed goods description"),
                             key="ef_commodity")
    st.markdown(T("**Spedizioniere incaricato** (default BS Cargo — modificabile)",
                  "**Assigned freight forwarder** (default BS Cargo — editable)"))
    cf1, cf2 = st.columns(2)
    forwarder = cf1.text_input(T("Società spedizioniere", "Freight forwarder company"),
                               value="BS CARGO SCS SRL", key="ef_forwarder")
    doganalista = cf2.text_input(T("Doganalista", "Customs broker"),
                                 value="MASSIMO TURINELLI", key="ef_doganalista")
    c7, c8 = st.columns(2)
    place = c7.text_input(T("Luogo firma", "Signature place"),
                          value="Gardone Val Trompia", key="ef_place")
    sign_date = c8.date_input(T("Data firma", "Signature date"), key="ef_signdate")
    if st.button(T("📜 Genera dichiarazioni (EUR1 + EUC)",
                   "📜 Generate declarations (EUR1 + EUC)"), type="primary", key="ef_gen"):
        f = {"company": "BCW", "buyer": buyer, "ef": {
            "rep": rep, "sender": sender, "seat": seat, "invNo": inv_no,
            "invDate": inv_date.isoformat(), "dest": dest, "contract": contract,
            "commodity": commodity, "place": place, "signDate": sign_date.isoformat(),
            "forwarder": forwarder, "doganalista": doganalista}}
        base = f"EUR1-EUC {inv_no or ''}".strip()
        html = doc_templates.document_html(doc_templates.render_forms(f), base)
        _archive_doc(base, html, {
            "number": inv_no or "", "company": "BCW", "doc_type": "declaration",
            "buyer": buyer.get("name") or buyer.get("trade") or "",
            "total": "", "currency": "",
            "doc_date": inv_date.isoformat(), "user": current_user()})
        st.session_state["ef_result"] = html
        st.success(T("Dichiarazioni generate e archiviate.",
                     "Declarations generated and archived."))
    if st.session_state.get("ef_result"):
        _doc_downloads(st.session_state["ef_result"],
                       f"EUR1-EUC {st.session_state.get('ef_invno','')}".strip())


def _doc_archive_view():
    st.markdown(T("**Registro documenti emessi** — ogni proforma, packing list e "
                  "dichiarazione generata viene salvata qui e può essere riscaricata.",
                  "**Issued documents register** — every generated proforma, packing "
                  "list and declaration is saved here and can be re-downloaded."))
    arch = DocArchive(get_backend())
    f1, f2, f3 = st.columns([2, 1, 1])
    query = f1.text_input(T("Cerca (numero, cliente, utente…)",
                            "Search (number, customer, user…)"), key="docarch_q")
    all_co = T("(tutte)", "(all)")
    company = f2.selectbox(T("Azienda", "Company"), [all_co, "BME", "BCW"], key="docarch_co")
    all_ty = T("(tutti)", "(all)")
    type_labels = {all_ty: None, **{v: k for k, v in DOC_TYPE_LABELS.items()}}
    ftype = f3.selectbox(T("Tipo", "Type"), list(type_labels.keys()), key="docarch_ty")
    entries = arch.list(
        query=query or None,
        company=None if company == all_co else company,
        doc_type=type_labels[ftype])
    st.caption(T(f"{len(entries)} documenti", f"{len(entries)} documents"))
    if not entries:
        st.info(T("Nessun documento archiviato ancora.", "No documents archived yet."))
        return
    hdr = st.columns([2, 1, 1, 2, 2, 1, 1, 1])
    labels = [T("Numero", "Number"), T("Azienda", "Company"), T("Tipo", "Type"),
              T("Cliente", "Customer"), T("Data", "Date"), T("Utente", "User"),
              "PDF", "HTML"]
    for col, label in zip(hdr, labels):
        col.markdown(f"**{label}**")
    for e in entries:
        c = st.columns([2, 1, 1, 2, 2, 1, 1, 1])
        c[0].write(e.get("number") or "—")
        c[1].write(e.get("company") or "")
        c[2].write(DOC_TYPE_LABELS.get(e.get("doc_type"), e.get("doc_type", "")))
        c[3].write(e.get("buyer") or "")
        c[4].write((e.get("doc_date") or e.get("issued_at", ""))[:10])
        c[5].write(e.get("user") or "")
        if e.get("pdf_key"):
            data, fname = arch.get_file(e["id"], "pdf")
            c[6].download_button("⬇️", data, file_name=fname, mime="application/pdf",
                                 key=f"da_pdf_{e['id']}")
        else:
            c[6].write("—")
        hdata, hname = arch.get_file(e["id"], "html")
        c[7].download_button("⬇️", hdata, file_name=hname, mime="text/html",
                             key=f"da_html_{e['id']}")


def section_documents():
    st.subheader(T("📄 Documenti", "📄 Documents"))
    sub = st.tabs([T("🧾 Proforma", "🧾 Proforma"),
                   T("📦 Packing list", "📦 Packing list"),
                   T("📜 Dichiarazioni (EUR1+EUC)", "📜 Declarations (EUR1+EUC)"),
                   T("📚 Documenti emessi", "📚 Issued documents")])
    with sub[0]:
        _doc_proforma()
    with sub[1]:
        _doc_packing()
    with sub[2]:
        _doc_declarations()
    with sub[3]:
        _doc_archive_view()


# ── main ──────────────────────────────────────────────────────────────────────
def section_barcode():
    st.subheader(T("🏷️ Codici a barre / Etichette", "🏷️ Barcodes / Labels"))
    if not _BARCODE_OK:
        st.info(T("Librerie codici a barre non disponibili sul server "
                  "(qrcode, python-barcode, reportlab).",
                  "Barcode libraries are not available on the server "
                  "(qrcode, python-barcode, reportlab)."))
        return

    # ── Generate labels ───────────────────────────────────────────────────────
    st.markdown(T("#### Genera etichette (QR + Code128)",
                  "#### Generate labels (QR + Code128)"))
    all_mode = st.checkbox(
        T("Tutte le armi in giacenza con matricola",
          "All in-stock weapons with a serial"), key="bc_all")
    if not all_mode:
        g1, g2 = st.columns(2)
        gen_matr = g1.text_input(T("Matricola arma", "Weapon serial"), key="bc_gen_matr")
        gen_nop = g2.text_input(T("N. operazione", "Operation no."), key="bc_gen_nop")
    else:
        gen_matr = gen_nop = ""
    if st.button(T("🏷️ Genera etichette", "🏷️ Generate labels"),
                 type="primary", key="bc_gen_go"):
        payload = {}
        ok_to_go = True
        if not all_mode:
            if gen_matr.strip():
                payload["matr_arma"] = gen_matr.strip()
            elif gen_nop.strip():
                payload["n_operazione"] = gen_nop.strip()
            else:
                st.warning(T("Inserisci una matricola o un n. operazione, "
                             "oppure spunta «tutte».",
                             "Enter a serial or operation no., or tick “all”."))
                ok_to_go = False
        if ok_to_go:
            def _do(_p):
                barcode_agent.EXCEL_FILE = _p
                return barcode_agent.cmd_genera(payload)
            try:
                res = with_workbook(_do)
            except Exception as e:
                res = {"status": "error", "message": str(e)}
            if res.get("status") == "success":
                pdf_path = res.get("pdf")
                data = None
                if pdf_path and os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as fh:
                        data = fh.read()
                st.session_state["bc_pdf"] = (
                    os.path.basename(pdf_path) if pdf_path else "labels.pdf", data)
                flash(T(f"{res.get('labels', 0)} etichetta/e generate.",
                        f"{res.get('labels', 0)} label(s) generated."))
                st.rerun()
            else:
                st.warning(res.get("message", T("Errore.", "Error.")))
    bc_pdf = st.session_state.get("bc_pdf")
    if bc_pdf and bc_pdf[1]:
        st.download_button(
            T("⬇️ Scarica PDF etichette", "⬇️ Download label PDF"),
            bc_pdf[1], file_name=bc_pdf[0], mime="application/pdf", key="bc_pdf_dl")

    st.divider()

    # ── Scan lookup ───────────────────────────────────────────────────────────
    st.markdown(T("#### Scansiona un codice", "#### Scan a code"))
    sc = st.text_input(T("Codice (scanner o manuale)", "Code (scanner or manual)"),
                       key="bc_scan_code")
    if st.button(T("🔍 Cerca codice", "🔍 Look up code"),
                 key="bc_scan_go") and sc.strip():
        def _scan(_p):
            barcode_agent.EXCEL_FILE = _p
            return barcode_agent.cmd_scan({"code": sc.strip()})
        try:
            rec = with_workbook(_scan, read_only=True)
        except Exception as e:
            rec = {"status": "error", "message": str(e)}
        if rec.get("status") == "found":
            st.success(T(f"Trovato alla riga {rec.get('row')} "
                         f"(corrisp.: {rec.get('matched_by')}).",
                         f"Found at row {rec.get('row')} "
                         f"(matched by {rec.get('matched_by')})."))
            if rec.get("venduto"):
                st.warning(T(f"⚠️ Articolo già VENDUTO a {rec.get('CLIENTE')}.",
                             f"⚠️ Item already SOLD to {rec.get('CLIENTE')}."))
            st.json({k: v for k, v in rec.items()
                     if k not in ("status", "matched_by", "venduto")})
        else:
            st.warning(rec.get("message", T("Non trovato.", "Not found.")))

    st.divider()

    # ── Link a maker's barcode to an existing item ────────────────────────────
    with st.expander(T("🔗 Collega un codice fornitore a un articolo",
                       "🔗 Link a maker's barcode to an item")):
        st.caption(T("Per armi in entrata che hanno già un codice a barre del "
                     "produttore: scansionalo qui per collegarlo alla riga, così "
                     "le scansioni future risalgono all'articolo.",
                     "For incoming weapons that already carry a maker's barcode: "
                     "scan it here to attach it to the row, so future scans map "
                     "back to the item."))
        l1, l2 = st.columns(2)
        link_matr = l1.text_input(T("Matricola arma", "Weapon serial"), key="bc_link_matr")
        link_nop = l2.text_input(T("N. operazione", "Operation no."), key="bc_link_nop")
        link_code = st.text_input(T("Codice fornitore scansionato",
                                    "Scanned maker barcode"), key="bc_link_code")
        if st.button(T("🔗 Collega", "🔗 Link"), key="bc_link_go"):
            if not link_code.strip() or not (link_matr.strip() or link_nop.strip()):
                st.warning(T("Servono il codice e (matricola o n. operazione).",
                             "Need the code and (serial or operation no.)."))
            else:
                d = {"code": link_code.strip()}
                if link_matr.strip():
                    d["matr_arma"] = link_matr.strip()
                if link_nop.strip():
                    d["n_operazione"] = link_nop.strip()

                def _link(_p):
                    barcode_agent.EXCEL_FILE = _p
                    return barcode_agent.cmd_link(d)
                try:
                    res = with_workbook(_link)
                except Exception as e:
                    res = {"status": "error", "message": str(e)}
                if res.get("status") == "success":
                    flash(T(f"Codice collegato alla riga {res['row']}.",
                            f"Code linked to row {res['row']}."))
                    st.rerun()
                else:
                    st.error(res.get("message"))


def main():
    st.title(T("🔫 BCW – Gestione Magazzino & Registro",
               "🔫 BCW – Stock & Register Management"))
    render_flash()
    api_key = get_api_key()
    backend = get_backend()
    with st.sidebar:
        # Language switch — default English, re-renders the whole UI on change.
        lang_choice = st.radio(
            "🌐 Language / Lingua", ["English", "Italiano"],
            index=0 if st.session_state.get("lang", "en") == "en" else 1,
            horizontal=True, key="_lang_radio")
        st.session_state["lang"] = "en" if lang_choice == "English" else "it"
        st.divider()
        st.text_input(T("Il tuo nome", "Your name"), key="user_name",
                      placeholder=T("es. Marco", "e.g. Marco"))
        if not api_key:
            k = st.text_input("Anthropic API Key", type="password")
            if st.button(T("Salva key", "Save key")):
                st.session_state["api_key"] = k; st.rerun()
        st.caption(T(f"Storage: {backend.name}", f"Storage: {backend.name}"))
    if not backend_ready(backend):
        st.warning(T("Workbook non inizializzato — vai su **Impostazioni** per caricarlo.",
                     "Workbook not initialized — go to **Settings** to upload it."))
        section_settings(backend); return
    tabs = st.tabs([T("📊 Dashboard", "📊 Dashboard"), T("➕ Carico", "➕ Load"),
                    T("➖ Scarico", "➖ Unload"), T("🔎 Cerca", "🔎 Search"),
                    T("📄 Documenti", "📄 Documents"), T("📑 Registro", "📑 Register"),
                    T("📁 Licenze", "📁 Licences"), T("🏷️ Barcode", "🏷️ Barcode"),
                    T("⚙️ Impostazioni", "⚙️ Settings")])
    with tabs[0]:
        section_dashboard()
    with tabs[1]:
        if api_key:
            section_carico(api_key)
        else:
            st.info(T("Inserisci l'API key nella barra laterale per l'estrazione automatica.",
                      "Enter the API key in the sidebar for automatic extraction."))
            section_carico("")
    with tabs[2]:
        if api_key:
            section_scarico(api_key)
        else:
            section_scarico("")
    with tabs[3]:
        section_search()
    with tabs[4]:
        section_documents()
    with tabs[5]:
        section_registro()
    with tabs[6]:
        section_licenses(api_key)
    with tabs[7]:
        section_barcode()
    with tabs[8]:
        section_settings(backend)


if __name__ == "__main__":
    main()
