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
from license_archive import LicenseArchive, DOC_TYPES

st.set_page_config(page_title="BCW Magazzino", page_icon="🔫", layout="wide")


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
        st.error("APP_PASSWORD non configurata. Aggiungila nei secrets dell'app "
                 "(Streamlit Cloud → Settings → Secrets) per abilitare l'accesso.")
        st.stop()

    if st.session_state.get("_authed"):
        return

    st.markdown("## 🔒 BCW Magazzino")
    st.caption("Inserisci la password per accedere.")
    pw = st.text_input("Password", type="password", key="_pw_input")
    if st.button("Entra"):
        if pw == expected:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("Password errata.")
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
        st.error(f"Errore accesso storage: {e}")
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
        st.error("PyMuPDF o Pillow non installato."); return [], "image/jpeg"


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
Output: array JSON compatto, NESSUN testo fuori dal JSON. [{...},{...}]"""
    else:
        fields_prompt = """Documento: fattura vendita armeria. Estrai OGNI arma venduta.
Campi per oggetto (usa null se assente):
matr_arma,cliente,prezzo_vendita,data_scarico,n_fattura,data_fattura_vendita,n_ata,titolo_acquisto
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
            return None, "Impossibile convertire il PDF."
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
    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-6", max_tokens=8192,
                system=system_msg, messages=[{"role": "user", "content": content}])
            break
        except anthropic.APIStatusError as e:
            if e.status_code == 529:
                last_exc = e; time.sleep(10 * (attempt + 1))
            else:
                raise
    else:
        return None, f"Claude API sovraccarica: {last_exc}"

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
            return recovered, f"⚠️ Risposta troncata: recuperati {len(recovered)} articoli."
        return None, f"Parsing JSON fallito. Risposta: {raw[:300]}"
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


# ── UI sections ───────────────────────────────────────────────────────────────

def section_dashboard():
    st.subheader("📊 Stato magazzino")
    try:
        s = with_workbook(lambda _p: stock_agent.stato(), read_only=True)
        pp = with_workbook(lambda _p: stock_agent.stato_per_prodotto(), read_only=True)
    except Exception as e:
        st.error(f"Impossibile leggere il magazzino: {e}"); return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Totale articoli", s["totale_articoli"])
    c2.metric("In giacenza", s["in_giacenza"])
    c3.metric("Venduti", s["venduti"])
    c4.metric("Max N. operazione", s["max_n_operazione"])

    tot = pp["totale"]
    v1, v2, v3 = st.columns(3)
    v1.metric("Valore costo in giacenza", f"€ {tot['cost_in_stock']:,.2f}")
    v2.metric("Valore vendite (venduti)", f"€ {tot['sale_value_sold']:,.2f}")
    v3.metric("Margine sui venduti", f"€ {tot['margin_sold']:,.2f}")

    st.markdown("##### Per tipologia di prodotto")
    rows = pp["per_prodotto"]
    if not rows:
        st.info("Nessun dato."); return
    df = pd.DataFrame(rows)[
        ["tipologia", "in_stock", "sold", "cost_in_stock", "cost_sold",
         "sale_value_sold", "margin_sold"]
    ].rename(columns={
        "tipologia": "Tipologia", "in_stock": "In giacenza", "sold": "Venduti",
        "cost_in_stock": "Costo giacenza €", "cost_sold": "Costo venduti €",
        "sale_value_sold": "Ricavo venduti €", "margin_sold": "Margine €"})
    st.dataframe(df, use_container_width=True, hide_index=True)
    chart = df[df["In giacenza"] > 0].set_index("Tipologia")["Costo giacenza €"]
    if not chart.empty:
        st.bar_chart(chart)


def section_search():
    st.subheader("🔎 Cerca articolo")

    # --- free-text search: marca / modello / matricola / calibro / tipologia ---
    st.markdown("**Ricerca per marca, modello, matricola, calibro…**")
    term = st.text_input("Cerca", key="search_term",
                         placeholder="es. Beretta, 686, 12, FUCILE…")
    if st.button("Cerca", type="primary", key="search_go") and term.strip():
        res = with_workbook(lambda _p: stock_agent.search_items(term), read_only=True)
        results = res.get("results", [])
        if not results:
            st.warning("Nessun articolo corrispondente.")
        else:
            st.success(f"{res['count']} articoli trovati"
                       + (" (mostrati i primi 300)" if res["count"] >= 300 else "") + ".")
            df = pd.DataFrame([{
                "Stato": r["stato"], "Tipologia": r["tipologia"], "Marca": r["marca"],
                "Modello": r["modello"], "Calibro": r["calibro"],
                "Matr. arma": r["matr_arma"], "Matr. canna": r["matr_canna"],
                "N.op": r["n_operazione"], "Data carico": r["data_carico"],
                "Fornitore": r["fornitore"], "Riga": r["row"],
            } for r in results])
            st.dataframe(df, hide_index=True, use_container_width=True)

    # --- exact lookup by matricola / n. operazione (returns the full record) ---
    with st.expander("Ricerca esatta per matricola / n. operazione"):
        col1, col2 = st.columns(2)
        matr = col1.text_input("Matricola arma", key="exact_matr")
        n_op = col2.text_input("N. operazione", key="exact_nop")
        if st.button("Cerca esatto", key="exact_go"):
            q = {}
            if matr.strip():
                q["matr_arma"] = matr.strip()
            if n_op.strip():
                q["n_operazione"] = n_op.strip()
            if not q:
                st.warning("Inserisci una matricola o un n. operazione.")
            else:
                res = with_workbook(lambda _p: stock_agent.cerca_item(q), read_only=True)
                if res.get("status") == "found":
                    st.success(f"Trovato alla riga {res['row']}.")
                    st.json({k: v for k, v in res.items() if k not in ("status",)})
                else:
                    st.warning("Articolo non trovato.")


def _extract_block(op_key, api_key):
    docs = st.file_uploader("DDT / fatture / foto (PDF, JPG, PNG) — multipli",
                            type=["pdf", "jpg", "jpeg", "png", "webp"],
                            accept_multiple_files=True, key=f"docs_{op_key}")
    if docs and st.button("📄 Estrai dati con AI", key=f"ext_{op_key}", type="primary"):
        all_items, errors = [], []
        bar = st.progress(0.0, text="Elaborazione…")
        for i, f in enumerate(docs):
            bar.progress(i / len(docs), text=f"Elaboro {f.name}…")
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
            st.success(f"{len(all_items)} articoli estratti.")
    return st.session_state.get(f"items_{op_key}")


def section_carico(api_key):
    st.subheader("➕ Carico (acquisto)")
    items = _extract_block("CARICO", api_key) or [{}]
    rows = []
    for it in items:
        rows.append({
            "data_carico": it.get("data_ddt") or it.get("data_fattura") or "",
            "n_operazione": 0,
            "tipologia": (it.get("tipologia") or "").upper(),
            "calibro": it.get("calibro") or "", "marca": it.get("marca") or "",
            "modello": it.get("modello") or "", "matr_arma": it.get("matr_arma") or "N/D",
            "matr_canna": it.get("matr_canna") or "N/D", "matr_agg": it.get("matr_agg") or "N/D",
            "fornitore": it.get("fornitore") or "",
            "costo": float(it.get("costo") or 0), "costo_imballo": float(it.get("costo_imballo") or 0),
            "data_ddt": it.get("data_ddt") or "", "data_fattura": it.get("data_fattura") or "",
        })
    st.caption("⚠️ Conferma la DATA DI CARICO (arrivo a Gardone) — spesso diversa dalla data del DDT.")
    with st.form("carico_form"):
        edited = st.data_editor(
            pd.DataFrame(rows), num_rows="dynamic", use_container_width=True,
            column_config={
                "n_operazione": st.column_config.NumberColumn("N. Op. (0=no registro)", step=1),
                "tipologia": st.column_config.SelectboxColumn("Tipologia *", options=TIPOLOGIE),
                "costo": st.column_config.NumberColumn("Costo €", format="%.2f"),
                "costo_imballo": st.column_config.NumberColumn("Imballo €", format="%.2f"),
            })
        submit = st.form_submit_button("✅ Aggiungi al MAGAZZINO", type="primary")
    if submit:
        def _do(_p):
            ok, errs = [], []
            for _, r in edited.iterrows():
                d = r.to_dict()
                try:
                    n = int(d.get("n_operazione") or 0)
                except (ValueError, TypeError):
                    n = 0
                d["n_operazione"] = n if n > 0 else None
                res = stock_agent.add_carico(d)
                (ok if res["status"] == "success" else errs).append(res)
            return ok, errs
        ok, errs = with_workbook(_do)
        for e in errs:
            st.error(e["message"])
        if ok:
            st.session_state.pop("items_CARICO", None)
            st.success(f"{len(ok)} carichi salvati nel registro condiviso.")


def section_scarico(api_key):
    st.subheader("➖ Scarico (vendita)")
    items = _extract_block("SCARICO", api_key) or [{}]
    rows = [{
        "matr_arma": it.get("matr_arma") or "", "cliente": it.get("cliente") or "",
        "prezzo_vendita": float(it.get("prezzo_vendita") or 0),
        "data_scarico": it.get("data_fattura_vendita") or it.get("data_scarico") or "",
        "n_fattura": it.get("n_fattura") or "",
        "data_fattura_vendita": it.get("data_fattura_vendita") or "",
        "n_ata": it.get("n_ata") or "", "titolo_acquisto": it.get("titolo_acquisto") or "",
    } for it in items]
    with st.form("scarico_form"):
        edited = st.data_editor(
            pd.DataFrame(rows), num_rows="dynamic", use_container_width=True,
            column_config={"prezzo_vendita": st.column_config.NumberColumn("Prezzo €", format="%.2f")})
        submit = st.form_submit_button("✅ Registra scarico", type="primary")
    if submit:
        def _do(_p):
            ok, errs = [], []
            for _, r in edited.iterrows():
                res = stock_agent.add_scarico(r.to_dict())
                (ok if res["status"] == "success" else errs).append(res)
            return ok, errs
        ok, errs = with_workbook(_do)
        for e in errs:
            st.error(e["message"])
        if ok:
            st.session_state.pop("items_SCARICO", None)
            st.success(f"{len(ok)} scarichi registrati.")


def section_registro():
    st.subheader("📑 Esporta REGISTRO (per la P.S.)")
    st.caption("Genera il registro aggiornato pronto da stampare, ricostruito dai dati attuali.")
    if st.button("Genera registro", type="primary"):
        import tempfile
        xlsx_path = os.path.join(tempfile.gettempdir(), "REGISTRO_export.xlsx")
        pdf_path = os.path.join(tempfile.gettempdir(), "REGISTRO_export.pdf")
        res = with_workbook(
            lambda _p: stock_agent.export_registro(xlsx_path, pdf_path), read_only=True)
        st.success(f"Registro generato: {res['righe']} operazioni.")
        with open(xlsx_path, "rb") as f:
            st.download_button("⬇️ Scarica REGISTRO (Excel)", f.read(),
                               file_name=f"REGISTRO {datetime.now():%Y-%m-%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if res.get("pdf") and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                st.download_button("⬇️ Scarica REGISTRO (PDF stampa)", f.read(),
                                   file_name=f"REGISTRO {datetime.now():%Y-%m-%d}.pdf",
                                   mime="application/pdf")
        else:
            st.info("PDF non generato (reportlab non installato) — usa l'Excel.")

    st.divider()
    st.markdown("**📥 Scarica il file MAGAZZINO completo (tutti i fogli, aggiornato)**")
    st.caption("Scarica l'intero workbook così com'è nello storage condiviso — "
               "ISTRUZIONI, INVENTARIO, MAGAZZINO, REGISTRO, BASE, PRINT.")
    if st.button("Prepara download MAGAZZINO completo", key="dl_full_wb"):
        try:
            data = get_backend().read_bytes(storage.workbook_name())
            st.download_button(
                "⬇️ Scarica MAGAZZINO completo (.xlsx)", data,
                file_name=f"{os.path.splitext(storage.workbook_name())[0]} "
                          f"{datetime.now():%Y-%m-%d %H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_full_wb_btn")
        except Exception as e:
            st.error(f"Impossibile scaricare il workbook: {e}")


def section_licenses():
    st.subheader("📁 Archivio licenze (Export / Import / EUC)")
    arch = LicenseArchive(get_backend())

    exp = arch.expiring_soon(days=60)
    if exp:
        st.warning("In scadenza / scaduti: " +
                   ", ".join(f"{e['doc_type']} {e.get('number','')} ({e['days_left']}g)" for e in exp[:8]))

    with st.expander("➕ Carica nuova licenza / EUC"):
        up = st.file_uploader("File (PDF/immagine/qualsiasi)", key="lic_up")
        c1, c2 = st.columns(2)
        doc_type = c1.selectbox("Tipo documento", DOC_TYPES)
        number = c2.text_input("Numero")
        c3, c4 = st.columns(2)
        country = c3.text_input("Paese")
        counterparty = c4.text_input("Controparte / Cliente")
        c5, c6 = st.columns(2)
        issue = c5.text_input("Data emissione (GG/MM/AAAA)")
        expiry = c6.text_input("Data scadenza (GG/MM/AAAA)")
        c7, c8 = st.columns(2)
        n_op = c7.text_input("N. operazione collegata")
        inv = c8.text_input("N. fattura collegata")
        tags = st.text_input("Tag (separati da virgola)")
        notes = st.text_area("Note")
        if st.button("Salva in archivio", type="primary"):
            if not up:
                st.error("Seleziona un file.")
            else:
                arch.add(up.read(), up.name, {
                    "doc_type": doc_type, "number": number, "country": country,
                    "counterparty": counterparty, "issue_date": issue, "expiry_date": expiry,
                    "n_operazione": n_op, "invoice_no": inv,
                    "tags": [t.strip() for t in tags.split(",") if t.strip()], "notes": notes})
                st.success("Licenza archiviata.")
                st.rerun()

    st.markdown("##### Archivio")
    fc1, fc2 = st.columns([2, 1])
    query = fc1.text_input("Cerca (numero, paese, controparte, n.op, tag…)")
    ftype = fc2.selectbox("Filtra tipo", ["(tutti)"] + DOC_TYPES)
    entries = arch.list(query=query or None, doc_type=None if ftype == "(tutti)" else ftype)
    if not entries:
        st.info("Nessuna licenza in archivio.")
        return
    for e in entries:
        with st.container(border=True):
            cols = st.columns([3, 2, 2, 2, 1])
            cols[0].markdown(f"**{e['doc_type']}** · {e.get('number','—')}\n\n{e['filename']}")
            cols[1].write(f"🌍 {e.get('country','—')}\n\n{e.get('counterparty','—')}")
            cols[2].write(f"Emessa: {e.get('issue_date','—')}\n\nScad.: {e.get('expiry_date','—')}")
            cols[3].write(f"N.Op: {e.get('n_operazione','—')}\n\nFatt: {e.get('invoice_no','—')}")
            data, fname = arch.get_file(e["id"])
            cols[4].download_button("⬇️", data, file_name=fname, key=f"dl_{e['id']}")
            if e.get("tags"):
                st.caption("🏷️ " + ", ".join(e["tags"]))
            if st.button("Elimina", key=f"del_{e['id']}"):
                arch.delete(e["id"]); st.rerun()


def section_settings(backend):
    st.subheader("⚙️ Impostazioni")
    st.write(f"**Backend storage:** `{backend.name}`")
    st.write(f"**Workbook:** `{storage.workbook_name()}`")
    st.write(f"**Utente corrente:** `{current_user()}` "
             "_(modificabile nella barra laterale)_")
    if not backend_ready(backend):
        st.warning("Il workbook non è ancora presente nello storage condiviso.")
        seed = st.file_uploader("Carica il MAGAZZINO iniziale per inizializzare", type=["xlsx"])
        if seed and st.button("Inizializza storage"):
            backend.write_bytes(storage.workbook_name(), seed.read())
            st.success("Workbook caricato nello storage condiviso."); st.rerun()
    else:
        st.success("Workbook presente nello storage condiviso.")
        if st.button("Verifica struttura workbook"):
            res = with_workbook(lambda _p: stock_agent.validate_workbook(), read_only=True)
            st.json(res)


# ── Documenti (proforma / packing / declarations) ─────────────────────────────

def _doc_downloads(full_html, base_name):
    """Render HTML preview + PDF/Word download buttons for a generated document."""
    import streamlit.components.v1 as components
    st.markdown("##### Anteprima")
    components.html(full_html, height=560, scrolling=True)
    c1, c2, c3 = st.columns(3)
    try:
        pdf_bytes = doc_templates.to_pdf(full_html)
        c1.download_button("⬇️ PDF", pdf_bytes, file_name=f"{base_name}.pdf",
                           mime="application/pdf", type="primary")
    except Exception as e:
        c1.warning(f"PDF non disponibile ({e}). Usa Word o stampa l'anteprima.")
    doc_bytes = doc_templates.to_doc(full_html)
    c2.download_button("⬇️ Word", doc_bytes, file_name=f"{base_name}.doc",
                       mime="application/msword")
    c3.download_button("⬇️ HTML", full_html.encode("utf-8"),
                       file_name=f"{base_name}.html", mime="text/html")


def _buyer_inputs(prefix):
    st.markdown("**Destinatario / Buyer**")
    b1, b2 = st.columns(2)
    name = b1.text_input("Nome / Name", key=f"{prefix}_b_name")
    trade = b2.text_input("Ragione sociale / Trade", key=f"{prefix}_b_trade")
    addr = st.text_input("Indirizzo / Address", key=f"{prefix}_b_addr")
    b3, b4 = st.columns(2)
    country = b3.text_input("Paese / Country", key=f"{prefix}_b_country")
    phone = b4.text_input("Telefono / Phone", key=f"{prefix}_b_phone")
    return {"name": name, "trade": trade, "addr": addr, "country": country, "phone": phone}


def _stock_picker(target_key, map_fn, label="➕ Aggiungi da magazzino (in giacenza)"):
    """BCW-only: pick in-stock guns and append mapped rows to st.session_state[target_key]."""
    with st.expander(label):
        q = st.text_input("Filtra (matricola, marca, modello, calibro, tipologia)",
                          key=f"{target_key}_pick_q")
        try:
            stock = with_workbook(lambda _p: stock_agent.list_in_stock(q or None),
                                  read_only=True)
        except Exception as e:
            st.error(f"Impossibile leggere il magazzino: {e}")
            return
        if not stock:
            st.info("Nessun articolo in giacenza corrispondente.")
            return
        st.caption(f"{len(stock)} articoli in giacenza.")
        df = pd.DataFrame([{
            "Sel": False, "Tipologia": s["tipologia"], "Marca": s["marca"],
            "Modello": s["modello"], "Calibro": s["calibro"],
            "Matr. arma": s["matr_arma"], "Matr. canna": s["matr_canna"],
            "Costo €": s["costo_compl"],
        } for s in stock])
        edited = st.data_editor(df, hide_index=True, use_container_width=True,
                                key=f"{target_key}_pick_tbl",
                                column_config={"Sel": st.column_config.CheckboxColumn("Sel")})
        if st.button("Aggiungi selezionati", key=f"{target_key}_pick_add"):
            chosen = [stock[i] for i in edited.index[edited["Sel"]].tolist()]
            if not chosen:
                st.warning("Nessuna riga selezionata.")
            else:
                cur = list(st.session_state.get(target_key, []))
                cur.extend(map_fn(s) for s in chosen)
                st.session_state[target_key] = cur
                st.success(f"{len(chosen)} articoli aggiunti.")
                st.rerun()


def _doc_proforma():
    company = st.radio("Azienda", ["BME", "BCW"], horizontal=True, key="pf_company")
    co = doc_templates.COMPANIES[company]
    backend = get_backend()
    st.caption(f"Prossimo numero: **{doc_numbering.preview_next(backend, company, 'proforma')}** "
               f"· Valuta: {co['currency']}")
    date = st.date_input("Data", key="pf_date")
    buyer = _buyer_inputs("pf")

    items_key = f"pf_items_{company}"
    if items_key not in st.session_state:
        st.session_state[items_key] = []

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
    notes = st.text_area("Note", key="pf_notes")

    if st.button("📄 Genera proforma", type="primary", key="pf_gen"):
        items = [r for r in edited.to_dict("records")
                 if any(str(v).strip() for v in r.values())]
        number = doc_numbering.next_number(backend, company, "proforma", owner=current_user())
        f = {"company": company, "number": number, "date": date.isoformat(),
             "currency": co["currency"], "buyer": buyer, "items": items,
             "notes": notes or ""}
        html = doc_templates.document_html(doc_templates.render_proforma(f), number)
        st.session_state["pf_result"] = (html, number)
        st.session_state[items_key] = []
        st.success(f"Proforma {number} generata.")
    if st.session_state.get("pf_result"):
        html, number = st.session_state["pf_result"]
        _doc_downloads(html, number)


def _doc_packing():
    company = st.radio("Azienda", ["BME", "BCW"], horizontal=True, key="pk_company")
    backend = get_backend()
    st.caption(f"Prossimo numero: **{doc_numbering.preview_next(backend, company, 'packing')}**")
    date = st.date_input("Data", key="pk_date")
    buyer = _buyer_inputs("pk")

    n_parcels = st.number_input("Numero di colli / parcels", min_value=1, max_value=50,
                                value=int(st.session_state.get("pk_nparcels", 1)), step=1,
                                key="pk_nparcels")
    pcols = ["qty", "type", "brand", "model", "caliber", "serial1", "serial2"]
    parcels_input = []
    for i in range(int(n_parcels)):
        st.markdown(f"**Collo / Parcel {i + 1}**")
        pk_key = f"pk_items_{company}_{i}"
        if company == "BCW":
            _stock_picker(pk_key, lambda s: {
                "qty": s["quantita"], "type": s["tipologia"], "brand": s["marca"],
                "model": s["modello"], "caliber": s["calibro"],
                "serial1": s["matr_arma"], "serial2": s["matr_canna"]},
                label=f"➕ Aggiungi da magazzino → collo {i + 1}")
        base = st.session_state.get(pk_key) or [{c: ("" if c != "qty" else 1) for c in pcols}]
        df = pd.DataFrame(base, columns=pcols)
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True,
                                key=f"pk_editor_{i}",
                                column_config={"qty": st.column_config.NumberColumn("qty", step=1)})
        d1, d2 = st.columns(2)
        dims = d1.text_input("Dimensioni pacco (es. 60x40x30 CM)", key=f"pk_dims_{i}")
        weight = d2.text_input("Peso pacco (es. 20,5 KG)", key=f"pk_weight_{i}")
        items = [r for r in edited.to_dict("records")
                 if any(str(v).strip() for v in r.values())]
        parcels_input.append({"items": items, "dims": dims, "weight": weight})

    notes = st.text_area("Note", key="pk_notes")
    if st.button("📦 Genera packing list", type="primary", key="pk_gen"):
        number = doc_numbering.next_number(backend, company, "packing", owner=current_user())
        f = {"company": company, "number": number, "date": date.isoformat(),
             "buyer": buyer, "parcels": parcels_input, "notes": notes or ""}
        html = doc_templates.document_html(doc_templates.render_packing(f), number)
        st.session_state["pk_result"] = (html, number)
        for i in range(int(n_parcels)):
            st.session_state.pop(f"pk_items_{company}_{i}", None)
        st.success(f"Packing list {number} generata.")
    if st.session_state.get("pk_result"):
        html, number = st.session_state["pk_result"]
        _doc_downloads(html, number)


def _doc_declarations():
    st.caption("EUR1 + End User Certificate (BCW). Emessi sulla fattura finale — "
               "non consumano un numero di proforma.")
    buyer = _buyer_inputs("ef")
    st.markdown("**Dettagli dichiarazione** — i dati BCW sono già precompilati (modificabili).")
    BCW_SEAT = "Via Matteotti 311 – Gardone Val Trompia, 25063 (BS) – Italia"
    c1, c2 = st.columns(2)
    rep = c1.text_input("Legale rappresentante (firma)", value="Antoine Abi Saab",
                        key="ef_rep")
    sender = c2.text_input("Ditta mittente", value=doc_templates.COMPANIES["BCW"]["name"],
                           key="ef_sender")
    seat = st.text_input("Sede ditta", value=BCW_SEAT, key="ef_seat")
    c3, c4 = st.columns(2)
    inv_no = c3.text_input("N. fattura esportazione", key="ef_invno")
    inv_date = c4.date_input("Data fattura", key="ef_invdate")
    c5, c6 = st.columns(2)
    dest = c5.text_input("Paese di destinazione", key="ef_dest")
    contract = c6.text_input("N. contratto (opz.)", key="ef_contract")
    commodity = st.text_area("Descrizione dettagliata merce", key="ef_commodity")
    st.markdown("**Spedizioniere incaricato** (default BS Cargo — modificabile)")
    cf1, cf2 = st.columns(2)
    forwarder = cf1.text_input("Società spedizioniere", value="BS CARGO SCS SRL",
                               key="ef_forwarder")
    doganalista = cf2.text_input("Doganalista", value="MASSIMO TURINELLI",
                                 key="ef_doganalista")
    c7, c8 = st.columns(2)
    place = c7.text_input("Luogo firma", value="Gardone Val Trompia", key="ef_place")
    sign_date = c8.date_input("Data firma", key="ef_signdate")

    if st.button("📜 Genera dichiarazioni (EUR1 + EUC)", type="primary", key="ef_gen"):
        f = {"company": "BCW", "buyer": buyer, "ef": {
            "rep": rep, "sender": sender, "seat": seat, "invNo": inv_no,
            "invDate": inv_date.isoformat(), "dest": dest, "contract": contract,
            "commodity": commodity, "place": place, "signDate": sign_date.isoformat(),
            "forwarder": forwarder, "doganalista": doganalista}}
        html = doc_templates.document_html(doc_templates.render_forms(f),
                                           f"EUR1-EUC {inv_no or ''}".strip())
        st.session_state["ef_result"] = html
        st.success("Dichiarazioni generate.")
    if st.session_state.get("ef_result"):
        _doc_downloads(st.session_state["ef_result"],
                       f"EUR1-EUC {st.session_state.get('ef_invno','')}".strip())


def section_documents():
    st.subheader("📄 Documenti")
    sub = st.tabs(["🧾 Proforma", "📦 Packing list", "📜 Dichiarazioni (EUR1+EUC)"])
    with sub[0]:
        _doc_proforma()
    with sub[1]:
        _doc_packing()
    with sub[2]:
        _doc_declarations()


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    st.title("🔫 BCW – Gestione Magazzino & Registro")
    api_key = get_api_key()
    backend = get_backend()

    with st.sidebar:
        st.text_input("Il tuo nome", key="user_name", placeholder="es. Marco")
        if not api_key:
            k = st.text_input("Anthropic API Key", type="password")
            if st.button("Salva key"):
                st.session_state["api_key"] = k; st.rerun()
        st.caption(f"Storage: {backend.name}")

    if not backend_ready(backend):
        st.warning("Workbook non inizializzato — vai su **Impostazioni** per caricarlo.")
        section_settings(backend); return

    tabs = st.tabs(["📊 Dashboard", "➕ Carico", "➖ Scarico", "🔎 Cerca",
                    "📄 Documenti", "📑 Registro", "📁 Licenze", "⚙️ Impostazioni"])
    with tabs[0]:
        section_dashboard()
    with tabs[1]:
        if api_key:
            section_carico(api_key)
        else:
            st.info("Inserisci l'API key nella barra laterale per l'estrazione automatica.")
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
        section_licenses()
    with tabs[7]:
        section_settings(backend)


if __name__ == "__main__":
    main()
