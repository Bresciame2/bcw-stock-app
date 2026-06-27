"""
BCW Stock Management Agent  (enhanced)
======================================
Core engine that reads/writes the real MAGAZZINO BCW workbook, preserving the
legal REGISTRO sheet and the formula columns. No UI, no API dependency.

Public API (all return JSON-serialisable dicts with a `status` key):
    stato()                     -> totals + max n.operazione
    stato_per_prodotto()        -> per-TIPOLOGIA breakdown (count + value)
    cerca_item(data)            -> full row for an item
    add_carico(data)            -> register a purchase (+ REGISTRO row if n_operazione)
    add_scarico(data)           -> record a sale (+ REGISTRO scarico update)
    export_registro(out_xlsx, out_pdf=None)   -> print-ready REGISTRO file(s)
    validate_workbook()         -> structural sanity check of the target file

CLI:
    python stock_agent.py stato
    python stock_agent.py prodotti
    python stock_agent.py cerca   '{"matr_arma": "ABC123"}'
    python stock_agent.py carico  '<json>'
    python stock_agent.py scarico '<json>'
    python stock_agent.py registro out.xlsx [out.pdf]
    python stock_agent.py validate

Set EXCEL_FILE (module attribute) to point at the workbook you manage.
"""

import json
import sys
import os
import shutil
import tempfile
from datetime import datetime

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ── paths ────────────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(SCRIPT_DIR, "MAGAZZINO BCW fixed.xlsx")

PINK_FILL = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)

# ── valid dropdown values ─────────────────────────────────────────────────────
# The authoritative list of TIPOLOGIE is the real workbook's dropdown source:
# INVENTARIO!B (the typology column the MAGAZZINO data-validation menu reads from).
# load_valid_tipologie() reads it live so that when a colleague adds a new
# typology to INVENTARIO (as the sheet's own instructions require), validation
# and the UI dropdown pick it up automatically — no code change needed.
# The constant below is only a FALLBACK used if INVENTARIO can't be read; it is
# kept in sync with the current workbook (32 values).
VALID_TIPOLOGIE = {
    "ACCESSORI", "CANNA PER CARABINA", "CANNA PER CARABINA AD ARIA COMPRESSA",
    "CANNA PER FUCILE DOPPIETTA", "CANNA PER FUCILE MONOCANNA",
    "CANNA PER FUCILE SEMIAUTOMATICO", "CANNA PER FUCILE SOVRAPPOSTO",
    "CANNA PER PISTOLA", "CANNA RIGATA", "CARABINA", "CARABINA A LEVA",
    "CARABINA SEMIAUTOMATICA", "CARABINA AD ARIA COMPRESSA",
    "CARABINA AD ARIA COMPRESSA PCP", "CARABINA MONOCOLPO",
    "CARABINA AD OTTURATORE", "DOPPIETTA A CANNE RIGATE", "FUCILE A POMPA",
    "FUCILE AVANCARICA", "FUCILE DOPPIETTA", "FUCILE DOPPIETTA CANI ESTERNI",
    "FUCILE DOPPIETTA CANI INTERNI", "FUCILE DOPPIETTA AVANCARICA",
    "FUCILE MONOCANNA", "FUCILE SEMIAUTOMATICO", "FUCILE AD OTTURATORE",
    "FUCILE SOVRAPPOSTO", "PISTOLA AVANCARICA", "PISTOLA SEMIAUTOMATICA",
    "PISTOLA POMPA", "REVOLVER", "CARCASSA",
}

# Sheet + column that holds the live typology dropdown source in the workbook.
TIPOLOGIA_SOURCE_SHEET = "INVENTARIO"
TIPOLOGIA_SOURCE_COL = 2          # column B
TIPOLOGIA_SOURCE_FIRST_ROW = 3    # first typology row (row 1-2 are headers)


def load_valid_tipologie(wb=None):
    """Return the set of valid TIPOLOGIE, read live from INVENTARIO!B.

    Reads contiguously from INVENTARIO!B3 downward, stopping at the first blank
    cell (which separates the typology block from the totals/instruction rows).
    Falls back to the VALID_TIPOLOGIE constant if the sheet is missing or empty.
    """
    try:
        own = wb is None
        if own:
            wb = load_wb()
        if TIPOLOGIA_SOURCE_SHEET not in wb.sheetnames:
            return set(VALID_TIPOLOGIE)
        ws = wb[TIPOLOGIA_SOURCE_SHEET]
        values = set()
        r = TIPOLOGIA_SOURCE_FIRST_ROW
        while True:
            v = ws.cell(r, TIPOLOGIA_SOURCE_COL).value
            if v is None or str(v).strip() == "":
                break
            # ignore any stray formula cell (dropdown source is plain text)
            s = str(v).strip()
            if not s.startswith("="):
                values.add(s.upper())
            r += 1
        return values or set(VALID_TIPOLOGIE)
    except Exception:
        return set(VALID_TIPOLOGIE)

# Column map for MAGAZZINO (1-based). Keep identical to the workbook contract.
COL = {
    "GIACENZA": 1, "N_OPERAZIONE": 2, "DATA_CARICO": 3, "QUANTITA": 4,
    "TIPOLOGIA": 5, "CALIBRO": 6, "MARCA": 7, "MODELLO": 8,
    "MATR_ARMA": 9, "MATR_CANNA": 10, "MATR_AGG": 11, "FORNITORE": 12,
    "COSTO": 13, "COSTO_IMBALLO": 14, "COSTO_COMPL": 15, "DATA_DDT": 16,
    "DATA_FATTURA_ACQ": 17, "PREZZO_VENDITA": 18, "V": 19, "N_ATA": 20,
    "DATA_SCARICO": 21, "N_FATTURA_VEND": 22, "DATA_FATTURA_VEND": 23, "CLIENTE": 24,
}


# ── helpers ──────────────────────────────────────────────────────────────────

def parse_date(s):
    if not s:
        return None
    if isinstance(s, datetime):
        return s
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except ValueError:
            pass
    return None


def _fmt_date(dt):
    if isinstance(dt, datetime):
        return dt.strftime("%d/%m/%Y")
    return str(dt) if dt not in (None, "") else ""


def load_wb():
    return load_workbook(EXCEL_FILE, data_only=False)


def save_wb(wb):
    """Atomic save with timestamped backup.

    Writes to a temp file in the same directory, then os.replace() onto the
    target so a crash mid-write can never leave a half-written register.
    Keeps a *_backup.xlsx copy of the previous good version.
    """
    target_dir = os.path.dirname(os.path.abspath(EXCEL_FILE)) or "."
    if os.path.exists(EXCEL_FILE):
        backup = EXCEL_FILE.replace(".xlsx", "_backup.xlsx")
        shutil.copy2(EXCEL_FILE, backup)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=target_dir)
    os.close(fd)
    wb.save(tmp)
    os.replace(tmp, EXCEL_FILE)


def find_last_data_row(ws):
    """Last real PURCHASE row, detected via col C (date) + col L (fornitore) only.
    Reserved N.OP slots and summary/category rows are ignored."""
    last = 2
    for r in range(3, ws.max_row + 1):
        c = ws.cell(r, COL["DATA_CARICO"]).value
        l = ws.cell(r, COL["FORNITORE"]).value
        if any(v not in (None, "") for v in [c, l]):
            last = r
    return last


def get_max_operazione(ws):
    max_op = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, COL["N_OPERAZIONE"]).value
        if isinstance(v, (int, float)) and v > max_op:
            max_op = int(v)
    return max_op


def find_item_row(ws, data):
    matr = str(data.get("matr_arma", "")).strip().upper()
    n_op = data.get("n_operazione")
    for r in range(3, ws.max_row + 1):
        if matr and matr not in ("", "N/D"):
            cell_val = ws.cell(r, COL["MATR_ARMA"]).value
            if cell_val is not None and str(cell_val).strip().upper() == matr:
                return r
        if n_op:
            cell_val = ws.cell(r, COL["N_OPERAZIONE"]).value
            if cell_val is not None:
                try:
                    if int(cell_val) == int(n_op):
                        return r
                except (ValueError, TypeError):
                    pass
    return None


def _to_float(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def suggest_serials(ws, matr, limit=3, only_instock=True):
    """Return up to `limit` matricole closest to `matr` (likely OCR misreads).

    Uses difflib similarity with a digit/letter-confusion normalization
    (O↔0, I/L↔1, S↔5, B↔8, Z↔2, G↔6) so that e.g. 'F36490IT17' suggests
    'F364901T17'. Each result is a dict: {matricola, riga, venduto, score}."""
    import difflib
    target = str(matr or "").strip().upper()
    if not target:
        return []
    confuse = str.maketrans({"O": "0", "I": "1", "L": "1", "S": "5",
                             "B": "8", "Z": "2", "G": "6", "Q": "0"})
    def norm(s):
        return str(s).strip().upper().translate(confuse)
    tnorm = norm(target)
    cands = []
    for r in range(3, ws.max_row + 1):
        cv = ws.cell(r, COL["MATR_ARMA"]).value
        if cv in (None, "", "N/D"):
            continue
        sold = _is_sold(ws, r)
        if only_instock and sold:
            continue
        cs = str(cv).strip().upper()
        score = difflib.SequenceMatcher(None, tnorm, norm(cs)).ratio()
        cands.append({"matricola": cs, "riga": r, "venduto": sold, "score": score})
    cands.sort(key=lambda c: c["score"], reverse=True)
    # only keep reasonably-close suggestions
    return [c for c in cands if c["score"] >= 0.6][:limit]


def _norm_serial(s):
    confuse = str.maketrans({"O": "0", "I": "1", "L": "1", "S": "5",
                             "B": "8", "Z": "2", "G": "6", "Q": "0"})
    return str(s or "").strip().upper().translate(confuse)


def _field(it, *keys):
    for k in keys:
        v = it.get(k)
        if v not in (None, ""):
            return str(v).strip()
    return ""


def reconcile_export(invoice_items, permit_items, ws=None):
    """Cross-check the sales invoice against the export permit for customs.

    Confirms every gun on the invoice is on the permit (and vice-versa) and that
    matched serials agree on marca / modello / calibro. When `ws` (MAGAZZINO) is
    given, the invoice side is enriched with the authoritative registered
    marca/modello/calibro/tipologia so the permit is checked against the real
    stock record, not just the invoice text.

    Items are dicts with any of: matr_arma/matricola, marca/brand,
    modello/model, calibro/cal/caliber, tipologia/tipo.

    Returns {ok, matched, only_invoice[], only_permit[], mismatches[]}.
    Serials are matched after folding common OCR confusions (O↔0, I/L↔1, …);
    a matched pair whose raw serials still differ is flagged as a likely OCR
    discrepancy so it still gets a human look."""
    def serial(it):
        return _field(it, "matr_arma", "matricola", "serial")

    def base(it):
        return {"matr_arma": serial(it),
                "marca": _field(it, "marca", "brand"),
                "modello": _field(it, "modello", "model"),
                "calibro": _field(it, "calibro", "cal", "caliber"),
                "tipologia": _field(it, "tipologia", "tipo")}

    def ref(it):
        out = base(it)
        if ws is not None and out["matr_arma"]:
            row = find_item_row(ws, {"matr_arma": out["matr_arma"]})
            if row:
                for label, col in (("marca", "MARCA"), ("modello", "MODELLO"),
                                   ("calibro", "CALIBRO"), ("tipologia", "TIPOLOGIA")):
                    if not out[label]:
                        out[label] = str(ws.cell(row, COL[col]).value or "").strip()
        return out

    inv, per = {}, {}
    for it in invoice_items:
        if serial(it):
            inv[_norm_serial(serial(it))] = ref(it)
    for it in permit_items:
        if serial(it):
            per[_norm_serial(serial(it))] = base(it)

    only_invoice = [inv[k] for k in inv if k not in per]
    only_permit = [per[k] for k in per if k not in inv]
    mismatches, matched = [], 0
    for k, a in inv.items():
        b = per.get(k)
        if not b:
            continue
        matched += 1
        if a["matr_arma"].strip().upper() != b["matr_arma"].strip().upper():
            mismatches.append({"matricola": a["matr_arma"], "campo": "matricola",
                               "fattura": a["matr_arma"], "permesso": b["matr_arma"],
                               "nota": "possibile errore di lettura (OCR)"})
        for label in ("marca", "modello", "calibro"):
            va, vb = a[label], b[label]
            if va and vb and va.upper() != vb.upper():
                mismatches.append({"matricola": a["matr_arma"], "campo": label,
                                   "fattura": va, "permesso": vb})
    ok = not only_invoice and not only_permit and not mismatches
    return {"ok": ok, "matched": matched, "only_invoice": only_invoice,
            "only_permit": only_permit, "mismatches": mismatches}


# ── VALIDATION / HARDENING ────────────────────────────────────────────────────

def validate_workbook():
    """Structural sanity check. Returns status + list of warnings/errors."""
    if not os.path.exists(EXCEL_FILE):
        return {"status": "error", "message": f"File non trovato: {EXCEL_FILE}"}
    try:
        wb = load_wb()
    except Exception as e:
        return {"status": "error", "message": f"Impossibile aprire il file: {e}"}
    problems = []
    for sheet in ("MAGAZZINO", "REGISTRO"):
        if sheet not in wb.sheetnames:
            problems.append(f"Foglio mancante: {sheet}")
    if "MAGAZZINO" in wb.sheetnames:
        ws = wb["MAGAZZINO"]
        # Spot-check that the formula columns are still formulas on a data row.
        last = find_last_data_row(ws)
        if last >= 3:
            a = ws.cell(last, COL["GIACENZA"]).value
            o = ws.cell(last, COL["COSTO_COMPL"]).value
            if not (isinstance(a, str) and a.startswith("=")):
                problems.append(f"Col A (GIACENZA) riga {last} non è una formula: {a!r}")
            if not (isinstance(o, str) and o.startswith("=")):
                problems.append(f"Col O (COSTO COMPL.) riga {last} non è una formula: {o!r}")
    return {
        "status": "ok" if not problems else "warning",
        "file": EXCEL_FILE,
        "problems": problems,
    }


def _validate_carico(data, valid_tipologie=None):
    """Return list of human-readable problems (empty = ok).

    valid_tipologie: the live set from INVENTARIO. If None, it is read from the
    workbook (falls back to the VALID_TIPOLOGIE constant).
    """
    if valid_tipologie is None:
        valid_tipologie = load_valid_tipologie()
    errs = []
    tip = (data.get("tipologia") or "").strip().upper()
    if not tip:
        errs.append("TIPOLOGIA mancante.")
    elif tip not in valid_tipologie:
        errs.append(f"TIPOLOGIA '{tip}' non è un valore valido del menù a tendina.")
    if not parse_date(data.get("data_carico")):
        errs.append("DATA DI CARICO mancante o non valida (atteso GG/MM/AAAA).")
    if data.get("costo") is not None and _to_float(data.get("costo")) is None:
        errs.append("COSTO non numerico.")
    return errs


# ── CARICO ───────────────────────────────────────────────────────────────────

def find_instock_serial(ws, matr):
    """Return the row of an item with this matricola that is currently IN STOCK
    (not sold), or None. Used to block loading the same serial twice."""
    target = str(matr or "").strip().upper()
    if not target or target == "N/D":
        return None
    for r in range(3, ws.max_row + 1):
        cell_val = ws.cell(r, COL["MATR_ARMA"]).value
        if cell_val is not None and str(cell_val).strip().upper() == target:
            if not _is_sold(ws, r):
                return r
    return None


def add_carico(data, validate=True, allow_duplicate=False):
    wb = load_wb()
    ws = wb["MAGAZZINO"]

    if validate:
        problems = _validate_carico(data, load_valid_tipologie(wb))
        if problems:
            return {"status": "error", "message": " ".join(problems), "problems": problems}

    if not allow_duplicate:
        dup = find_instock_serial(ws, data.get("matr_arma"))
        if dup:
            matr = str(data.get("matr_arma") or "").strip()
            return {"status": "error",
                    "message": f"Matricola {matr} è già IN GIACENZA alla riga {dup}. "
                               f"Carico annullato per evitare un duplicato. "
                               f"(Se è un'arma usata rientrata, registra prima lo scarico "
                               f"della precedente o usa una matricola corretta.)"}

    n_op = data.get("n_operazione")
    if n_op:
        n_op = int(n_op)
        max_op = get_max_operazione(ws)
        if n_op <= max_op:
            return {"status": "error",
                    "message": f"N. OPERAZIONE {n_op} già esiste o è ≤ del massimo ({max_op}). "
                               f"Usa un valore > {max_op}."}

    r = find_last_data_row(ws) + 1

    ws.cell(r, COL["GIACENZA"]).value = f"=+D{r}-S{r}"

    existing_n_op = ws.cell(r, COL["N_OPERAZIONE"]).value
    if existing_n_op not in (None, ""):
        n_op = int(existing_n_op)
    elif n_op:
        ws.cell(r, COL["N_OPERAZIONE"]).value = n_op

    ws.cell(r, COL["DATA_CARICO"]).value = parse_date(data.get("data_carico"))
    ws.cell(r, COL["QUANTITA"]).value = int(data.get("quantita", 1) or 1)
    ws.cell(r, COL["TIPOLOGIA"]).value = (data.get("tipologia") or "").strip().upper()
    ws.cell(r, COL["CALIBRO"]).value = data.get("calibro", "")
    ws.cell(r, COL["MARCA"]).value = data.get("marca", "")
    ws.cell(r, COL["MODELLO"]).value = data.get("modello", "")
    ws.cell(r, COL["MATR_ARMA"]).value = data.get("matr_arma") or "N/D"
    ws.cell(r, COL["MATR_CANNA"]).value = data.get("matr_canna") or "N/D"
    ws.cell(r, COL["MATR_AGG"]).value = data.get("matr_agg") or "N/D"
    ws.cell(r, COL["FORNITORE"]).value = data.get("fornitore", "")

    costo = _to_float(data.get("costo"))
    if costo is not None:
        ws.cell(r, COL["COSTO"]).value = costo
    costo_imb = _to_float(data.get("costo_imballo"))
    if costo_imb is not None:
        ws.cell(r, COL["COSTO_IMBALLO"]).value = costo_imb

    ws.cell(r, COL["COSTO_COMPL"]).value = f"=+M{r}+N{r}"
    ws.cell(r, COL["DATA_DDT"]).value = parse_date(data.get("data_ddt"))
    ws.cell(r, COL["DATA_FATTURA_ACQ"]).value = parse_date(data.get("data_fattura"))

    if n_op:
        _add_to_registro(wb, n_op, r)

    save_wb(wb)
    return {"status": "success",
            "message": f"Carico aggiunto alla riga {r} del foglio MAGAZZINO.",
            "row": r, "n_operazione": n_op}


def _add_to_registro(wb, n_operazione, mag_row):
    ws_reg = wb["REGISTRO"]
    next_row = 4
    for r in range(4, ws_reg.max_row + 2):
        if ws_reg.cell(r, 1).value is None:
            next_row = r
            break
    upper = max(mag_row + 200, 5000)
    ws_reg.cell(next_row, 1).value = n_operazione
    for offset in range(1, 12):
        ws_reg.cell(next_row, offset + 1).value = (
            f"=+VLOOKUP($A{next_row},MAGAZZINO!$B$2:$M${upper},{offset + 1},FALSE)"
        )


# ── SCARICO ──────────────────────────────────────────────────────────────────

def add_scarico(data):
    wb = load_wb()
    ws = wb["MAGAZZINO"]

    row = find_item_row(ws, data)
    if row is None:
        _matr = (str(data.get("matr_arma") or "")).strip()
        _nop = data.get("n_operazione") or ""
        _id = f"matricola '{_matr or '(vuota)'}'" + (f" / n.op {_nop}" if _nop else "")
        sugg = suggest_serials(ws, _matr) if _matr else []
        msg = (f"Articolo non trovato in MAGAZZINO ({_id}). "
               "Può essere una matricola letta male dal PDF, oppure l'arma non è "
               "in giacenza BCW.")
        if sugg:
            hint = ", ".join(
                f"{s['matricola']}" + (" [già venduto]" if s["venduto"] else "")
                for s in sugg)
            msg += f" Forse intendevi: {hint}?"
        return {"status": "error", "message": msg, "suggestions": sugg}

    if ws.cell(row, COL["V"]).value not in (None, 0, ""):
        existing_client = ws.cell(row, COL["CLIENTE"]).value
        return {"status": "error",
                "message": f"Articolo alla riga {row} già venduto a: {existing_client}."}

    pv = _to_float(data.get("prezzo_vendita"))
    if pv is not None:
        ws.cell(row, COL["PREZZO_VENDITA"]).value = pv
    ws.cell(row, COL["V"]).value = int(data.get("qty_venduta", 1) or 1)
    if data.get("n_ata"):
        ws.cell(row, COL["N_ATA"]).value = str(data["n_ata"])
    ws.cell(row, COL["DATA_SCARICO"]).value = parse_date(data.get("data_scarico"))
    if data.get("n_fattura"):
        ws.cell(row, COL["N_FATTURA_VEND"]).value = str(data["n_fattura"])
    ws.cell(row, COL["DATA_FATTURA_VEND"]).value = parse_date(data.get("data_fattura_vendita"))
    if data.get("cliente"):
        ws.cell(row, COL["CLIENTE"]).value = data["cliente"]

    for c in range(COL["PREZZO_VENDITA"], COL["CLIENTE"] + 1):
        ws.cell(row, c).fill = PINK_FILL

    n_op = ws.cell(row, COL["N_OPERAZIONE"]).value
    if n_op:
        _update_registro_scarico(wb, int(n_op), data)

    save_wb(wb)
    return {"status": "success",
            "message": f"Scarico registrato alla riga {row}.",
            "row": row, "n_operazione": int(n_op) if n_op else None}


def _update_registro_scarico(wb, n_operazione, data):
    ws_reg = wb["REGISTRO"]
    for r in range(4, ws_reg.max_row + 1):
        v = ws_reg.cell(r, 1).value
        if v is not None:
            try:
                if int(v) == n_operazione:
                    ws_reg.cell(r, 14).value = parse_date(data.get("data_scarico"))
                    if data.get("cliente"):
                        ws_reg.cell(r, 15).value = data["cliente"]
                    if data.get("titolo_acquisto"):
                        ws_reg.cell(r, 16).value = data["titolo_acquisto"]
                    break
            except (ValueError, TypeError):
                pass


def _clear_registro_scarico(wb, n_operazione):
    """Clear only the SCARICO columns (DATA/ACQUIRENTE/TITOLO = 14/15/16) of the
    REGISTRO row for this operazione, leaving the carico entry intact."""
    ws_reg = wb["REGISTRO"]
    for r in range(4, ws_reg.max_row + 1):
        v = ws_reg.cell(r, 1).value
        if v is not None:
            try:
                if int(v) == n_operazione:
                    for c in (14, 15, 16):
                        ws_reg.cell(r, c).value = None
                    return True
            except (ValueError, TypeError):
                pass
    return False


def _clear_registro_row(wb, n_operazione):
    """Clear the ENTIRE REGISTRO row for this operazione (carico mistake)."""
    ws_reg = wb["REGISTRO"]
    for r in range(4, ws_reg.max_row + 1):
        v = ws_reg.cell(r, 1).value
        if v is not None:
            try:
                if int(v) == n_operazione:
                    for c in range(1, 17):
                        ws_reg.cell(r, c).value = None
                    return True
            except (ValueError, TypeError):
                pass
    return False


def reverse_scarico(data):
    """Undo a sale recorded by mistake: clear the sold columns so the gun goes
    back IN GIACENZA, and clear the REGISTRO scarico columns. Safe and reversible
    (only clears cells, never shifts rows). Match by matr_arma or n_operazione."""
    wb = load_wb()
    ws = wb["MAGAZZINO"]
    row = find_item_row(ws, data)
    if row is None:
        return {"status": "error",
                "message": "Articolo non trovato in MAGAZZINO. Controlla matr_arma o n_operazione."}
    if ws.cell(row, COL["V"]).value in (None, 0, ""):
        return {"status": "error",
                "message": f"Articolo alla riga {row} non risulta venduto: niente da annullare."}
    for c in range(COL["PREZZO_VENDITA"], COL["CLIENTE"] + 1):
        ws.cell(row, c).value = None
        ws.cell(row, c).fill = NO_FILL
    n_op = ws.cell(row, COL["N_OPERAZIONE"]).value
    if n_op:
        _clear_registro_scarico(wb, int(n_op))
    save_wb(wb)
    return {"status": "success",
            "message": f"Scarico annullato (riga {row}). L'articolo è di nuovo in giacenza.",
            "row": row, "n_operazione": int(n_op) if n_op else None}


def delete_carico(data, allow_registro_gap=False):
    """Cancel a load entered by mistake. Clears the MAGAZZINO row IN PLACE — no
    row shifting, so the per-row formulas of every other row stay valid — and
    clears the matching REGISTRO row if present.

    Refuses if the item is already sold (annulla lo scarico prima). If the item
    is in the REGISTRO and is NOT the last operazione, deleting leaves a gap in
    the legal sequence: it returns a 'confirm' status unless allow_registro_gap
    is True."""
    wb = load_wb()
    ws = wb["MAGAZZINO"]
    row = find_item_row(ws, data)
    if row is None:
        return {"status": "error",
                "message": "Articolo non trovato in MAGAZZINO. Controlla matr_arma o n_operazione."}
    if _is_sold(ws, row):
        return {"status": "error",
                "message": f"Articolo alla riga {row} risulta VENDUTO. Annulla prima lo "
                           f"scarico, poi elimina il carico."}

    n_op = ws.cell(row, COL["N_OPERAZIONE"]).value
    n_op_i = None
    try:
        n_op_i = int(n_op) if n_op not in (None, "") else None
    except (ValueError, TypeError):
        n_op_i = None

    if n_op_i is not None and not allow_registro_gap:
        if n_op_i != get_max_operazione(ws):
            return {"status": "confirm", "needs_confirm": True, "row": row,
                    "n_operazione": n_op_i,
                    "message": f"L'operazione {n_op_i} non è l'ultima del registro: "
                               f"eliminandola resterà un buco nella sequenza legale. "
                               f"Confermi l'eliminazione?"}

    if n_op_i is not None:
        _clear_registro_row(wb, n_op_i)

    for c in range(1, COL["CLIENTE"] + 1):
        ws.cell(row, c).value = None
        ws.cell(row, c).fill = NO_FILL

    save_wb(wb)
    return {"status": "success",
            "message": f"Carico eliminato (riga {row}).",
            "row": row, "n_operazione": n_op_i}


# ── CERCA ────────────────────────────────────────────────────────────────────

HEADERS = [
    "GIACENZA", "N_OPERAZIONE", "DATA_CARICO", "QUANTITA", "TIPOLOGIA",
    "CALIBRO", "MARCA", "MODELLO", "MATR_ARMA", "MATR_CANNA", "MATR_AGG",
    "FORNITORE", "COSTO", "COSTO_IMBALLO", "COSTO_COMPL",
    "DATA_DDT", "DATA_FATTURA_ACQ", "PREZZO_VENDITA", "V", "N_ATA",
    "DATA_SCARICO", "N_FATTURA_VEND", "DATA_FATTURA_VEND", "CLIENTE",
]


def cerca_item(data):
    wb = load_wb()
    ws = wb["MAGAZZINO"]
    row = find_item_row(ws, data)
    if row is None:
        return {"status": "not_found", "message": "Articolo non trovato."}
    result = {}
    for col, header in enumerate(HEADERS, start=1):
        val = ws.cell(row, col).value
        if isinstance(val, datetime):
            val = val.strftime("%d/%m/%Y")
        result[header] = val
    result["row"] = row
    result["status"] = "found"
    return result


def search_items(term, limit=300):
    """Free-text search across MARCA / MODELLO / MATR_ARMA / MATR_CANNA /
    CALIBRO / TIPOLOGIA, case-insensitive substring. Returns EVERY matching
    row (in stock or sold) so colleagues can look a gun up by brand or model,
    not only by matricola. Skips empty/summary rows. For the Cerca tab."""
    term = (term or "").strip().lower()
    if not term:
        return {"status": "empty", "count": 0, "results": []}
    wb = load_wb(); ws = wb["MAGAZZINO"]
    out = []
    for r in range(3, ws.max_row + 1):
        tip = ws.cell(r, COL["TIPOLOGIA"]).value
        if tip in (None, ""):
            continue
        if ws.cell(r, COL["DATA_CARICO"]).value in (None, "") and \
           ws.cell(r, COL["FORNITORE"]).value in (None, ""):
            continue
        fields = [
            str(ws.cell(r, COL["MARCA"]).value or ""),
            str(ws.cell(r, COL["MODELLO"]).value or ""),
            str(ws.cell(r, COL["MATR_ARMA"]).value or ""),
            str(ws.cell(r, COL["MATR_CANNA"]).value or ""),
            str(ws.cell(r, COL["CALIBRO"]).value or ""),
            str(tip),
        ]
        if not any(term in f.lower() for f in fields):
            continue
        dc = ws.cell(r, COL["DATA_CARICO"]).value
        if isinstance(dc, datetime):
            dc = dc.strftime("%d/%m/%Y")
        out.append({
            "row": r,
            "n_operazione": ws.cell(r, COL["N_OPERAZIONE"]).value,
            "data_carico": dc,
            "tipologia": str(tip).strip(),
            "calibro": str(ws.cell(r, COL["CALIBRO"]).value or "").strip(),
            "marca": str(ws.cell(r, COL["MARCA"]).value or "").strip(),
            "modello": str(ws.cell(r, COL["MODELLO"]).value or "").strip(),
            "matr_arma": str(ws.cell(r, COL["MATR_ARMA"]).value or "").strip(),
            "matr_canna": str(ws.cell(r, COL["MATR_CANNA"]).value or "").strip(),
            "fornitore": str(ws.cell(r, COL["FORNITORE"]).value or "").strip(),
            "stato": "VENDUTO" if _is_sold(ws, r) else "IN GIACENZA",
        })
        if len(out) >= limit:
            break
    return {"status": "ok", "count": len(out), "results": out}


# ── STATO / PER-PRODOTTO ──────────────────────────────────────────────────────

def _is_sold(ws, r):
    v = ws.cell(r, COL["V"]).value
    return bool(v and int(v) > 0)


def stato():
    wb = load_wb(); ws = wb["MAGAZZINO"]
    total = in_stock = sold = 0
    for r in range(3, ws.max_row + 1):
        tip = ws.cell(r, COL["TIPOLOGIA"]).value
        if tip in (None, ""):
            continue
        # skip summary/category rows (no date AND no fornitore)
        if ws.cell(r, COL["DATA_CARICO"]).value in (None, "") and \
           ws.cell(r, COL["FORNITORE"]).value in (None, ""):
            continue
        total += 1
        if _is_sold(ws, r):
            sold += 1
        else:
            in_stock += 1
    return {"status": "ok", "totale_articoli": total, "in_giacenza": in_stock,
            "venduti": sold, "max_n_operazione": get_max_operazione(ws),
            "ultimo_row_magazzino": find_last_data_row(ws)}


def stato_per_prodotto():
    """Per-TIPOLOGIA breakdown: counts and €-value of stock on hand and sold."""
    wb = load_wb(); ws = wb["MAGAZZINO"]
    cats = {}
    tot = {"in_stock": 0, "sold": 0, "cost_in_stock": 0.0,
           "cost_sold": 0.0, "sale_value_sold": 0.0}
    for r in range(3, ws.max_row + 1):
        tip = ws.cell(r, COL["TIPOLOGIA"]).value
        if tip in (None, ""):
            continue
        if ws.cell(r, COL["DATA_CARICO"]).value in (None, "") and \
           ws.cell(r, COL["FORNITORE"]).value in (None, ""):
            continue
        tip = str(tip).strip().upper()
        c = cats.setdefault(tip, {"in_stock": 0, "sold": 0, "cost_in_stock": 0.0,
                                  "cost_sold": 0.0, "sale_value_sold": 0.0})
        costo = _to_float(ws.cell(r, COL["COSTO"]).value) or 0.0
        imb = _to_float(ws.cell(r, COL["COSTO_IMBALLO"]).value) or 0.0
        costo_compl = costo + imb
        if _is_sold(ws, r):
            pv = _to_float(ws.cell(r, COL["PREZZO_VENDITA"]).value) or 0.0
            c["sold"] += 1; c["cost_sold"] += costo_compl; c["sale_value_sold"] += pv
            tot["sold"] += 1; tot["cost_sold"] += costo_compl; tot["sale_value_sold"] += pv
        else:
            c["in_stock"] += 1; c["cost_in_stock"] += costo_compl
            tot["in_stock"] += 1; tot["cost_in_stock"] += costo_compl
    rows = []
    for tip in sorted(cats):
        d = cats[tip]; d["tipologia"] = tip
        d["margin_sold"] = round(d["sale_value_sold"] - d["cost_sold"], 2)
        for k in ("cost_in_stock", "cost_sold", "sale_value_sold"):
            d[k] = round(d[k], 2)
        rows.append(d)
    for k in ("cost_in_stock", "cost_sold", "sale_value_sold"):
        tot[k] = round(tot[k], 2)
    tot["margin_sold"] = round(tot["sale_value_sold"] - tot["cost_sold"], 2)
    return {"status": "ok", "per_prodotto": rows, "totale": tot}


def list_in_stock(query=None):
    """List items currently IN STOCK (not sold), for the document line-item picker.

    Applies the same in-stock filter as stato()/stato_per_prodotto():
      - skip rows with empty TIPOLOGIA
      - skip summary/category rows (no DATA_CARICO AND no FORNITORE)
      - exclude sold rows (_is_sold)
    Optional `query` filters case-insensitively across matricola / marca /
    modello / calibro / tipologia.
    Returns a list of dicts ready to map onto invoice / packing line items.
    """
    wb = load_wb(); ws = wb["MAGAZZINO"]
    q = (query or "").strip().lower()
    out = []
    for r in range(3, ws.max_row + 1):
        tip = ws.cell(r, COL["TIPOLOGIA"]).value
        if tip in (None, ""):
            continue
        if ws.cell(r, COL["DATA_CARICO"]).value in (None, "") and \
           ws.cell(r, COL["FORNITORE"]).value in (None, ""):
            continue
        if _is_sold(ws, r):
            continue
        costo = _to_float(ws.cell(r, COL["COSTO"]).value) or 0.0
        imb = _to_float(ws.cell(r, COL["COSTO_IMBALLO"]).value) or 0.0
        qta = ws.cell(r, COL["QUANTITA"]).value
        try:
            qta = int(qta) if qta not in (None, "") else 1
        except (ValueError, TypeError):
            qta = 1
        item = {
            "row": r,
            "n_operazione": ws.cell(r, COL["N_OPERAZIONE"]).value,
            "tipologia": str(tip).strip(),
            "calibro": str(ws.cell(r, COL["CALIBRO"]).value or "").strip(),
            "marca": str(ws.cell(r, COL["MARCA"]).value or "").strip(),
            "modello": str(ws.cell(r, COL["MODELLO"]).value or "").strip(),
            "matr_arma": str(ws.cell(r, COL["MATR_ARMA"]).value or "").strip(),
            "matr_canna": str(ws.cell(r, COL["MATR_CANNA"]).value or "").strip(),
            "quantita": qta,
            "costo": costo,
            "costo_imballo": imb,
            "costo_compl": round(costo + imb, 2),
            "fornitore": str(ws.cell(r, COL["FORNITORE"]).value or "").strip(),
        }
        if q:
            hay = " ".join(str(item[k]) for k in
                           ("matr_arma", "marca", "modello", "calibro", "tipologia")).lower()
            if q not in hay:
                continue
        out.append(item)
    return out


# ── REGISTRO EXPORT (print-ready) ─────────────────────────────────────────────

# REGISTRO columns rebuilt directly from MAGAZZINO (source of truth) so the
# export never depends on cached VLOOKUP values.
_REG_COLS = [
    ("N. OP.", "N_OPERAZIONE"),
    ("DATA CARICO", "DATA_CARICO"),
    ("Q.TA", "QUANTITA"),
    ("TIPOLOGIA", "TIPOLOGIA"),
    ("CALIBRO", "CALIBRO"),
    ("MARCA", "MARCA"),
    ("MODELLO", "MODELLO"),
    ("MATR. ARMA", "MATR_ARMA"),
    ("MATR. CANNA", "MATR_CANNA"),
    ("MATR. AGG.", "MATR_AGG"),
    ("FORNITORE", "FORNITORE"),
    ("DATA SCARICO", "DATA_SCARICO"),
    ("CLIENTE", "CLIENTE"),
]


def _registro_rows():
    wb = load_wb(); ws = wb["MAGAZZINO"]
    out = []
    for r in range(3, ws.max_row + 1):
        n_op = ws.cell(r, COL["N_OPERAZIONE"]).value
        if not isinstance(n_op, (int, float)):
            continue
        # only rows that are real operations (have a carico date)
        if ws.cell(r, COL["DATA_CARICO"]).value in (None, ""):
            continue
        rowvals = []
        for _label, key in _REG_COLS:
            v = ws.cell(r, COL[key]).value
            if isinstance(v, datetime):
                v = _fmt_date(v)
            rowvals.append(v if v is not None else "")
        rowvals[0] = int(n_op)
        out.append(rowvals)
    out.sort(key=lambda x: x[0])
    return out


def export_registro(out_xlsx, out_pdf=None, intestazione="REGISTRO OPERAZIONI"):
    rows = _registro_rows()
    # ---- Excel ----
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "REGISTRO"
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title = f"{intestazione} — generato il {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws.cell(1, 1).value = title
    ws.cell(1, 1).font = Font(bold=True, size=12)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(_REG_COLS))
    hdr = [lab for lab, _ in _REG_COLS]
    for c, lab in enumerate(hdr, start=1):
        cell = ws.cell(3, c); cell.value = lab
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="404040")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for i, rowvals in enumerate(rows, start=4):
        for c, val in enumerate(rowvals, start=1):
            cell = ws.cell(i, c); cell.value = val; cell.border = border
            cell.alignment = Alignment(vertical="center")
    widths = [7, 12, 5, 22, 12, 14, 18, 16, 16, 14, 20, 12, 22]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    wb.save(out_xlsx)

    pdf_made = False
    if out_pdf:
        pdf_made = _registro_pdf(rows, hdr, out_pdf, title)

    return {"status": "success", "xlsx": out_xlsx,
            "pdf": out_pdf if pdf_made else None, "righe": len(rows)}


def _registro_pdf(rows, hdr, out_pdf, title):
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
    except ImportError:
        return False
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(out_pdf, pagesize=landscape(A4),
                            leftMargin=8 * mm, rightMargin=8 * mm,
                            topMargin=10 * mm, bottomMargin=10 * mm)
    data = [hdr] + [[("" if v is None else str(v)) for v in r] for r in rows]
    tbl = Table(data, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#404040")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    doc.build([Paragraph(title, styles["Heading4"]), Spacer(1, 4), tbl])
    return True


# ── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "message": "carico | scarico | cerca | stato | prodotti | registro | validate"}))
        sys.exit(1)
    cmd = sys.argv[1].lower()

    if cmd == "stato":
        print(json.dumps(stato(), ensure_ascii=False, indent=2)); sys.exit(0)
    if cmd == "prodotti":
        print(json.dumps(stato_per_prodotto(), ensure_ascii=False, indent=2)); sys.exit(0)
    if cmd == "validate":
        print(json.dumps(validate_workbook(), ensure_ascii=False, indent=2)); sys.exit(0)
    if cmd == "registro":
        out_xlsx = sys.argv[2] if len(sys.argv) > 2 else "REGISTRO_export.xlsx"
        out_pdf = sys.argv[3] if len(sys.argv) > 3 else None
        print(json.dumps(export_registro(out_xlsx, out_pdf), ensure_ascii=False, indent=2)); sys.exit(0)

    if len(sys.argv) < 3:
        print(json.dumps({"status": "error", "message": "JSON dati mancante"})); sys.exit(1)
    try:
        data = json.loads(sys.argv[2])
    except json.JSONDecodeError as e:
        print(json.dumps({"status": "error", "message": f"JSON non valido: {e}"})); sys.exit(1)

    if cmd == "carico":
        result = add_carico(data)
    elif cmd == "scarico":
        result = add_scarico(data)
    elif cmd == "cerca":
        result = cerca_item(data)
    else:
        result = {"status": "error", "message": f"Comando sconosciuto: {cmd}"}
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
