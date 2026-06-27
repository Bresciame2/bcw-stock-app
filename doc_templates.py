"""
Brescia document templates — faithful Python port of the HTML generator.
========================================================================
This module reproduces, byte-for-byte where it matters, the proforma / packing
/ export-form layouts and CSS from "Brescia Document Generator.html", so the
hosted Python app renders documents identical to the ones the shop already uses.

Two companies:
    BME = Brescia Middle East S.a.l. (Lebanon, USD $)
    BCW = Brescia Custom Works S.R.L. (Italy, EUR €)

Public API:
    COMPANIES                      -> config dict
    render_proforma(f)             -> HTML string for one proforma invoice
    render_packing(f)              -> HTML string for a packing list (parcels)
    render_forms(f)                -> HTML string for EUR1 + End User Certificate
    document_html(body, title)     -> wrap a render in a full standalone HTML page
    to_pdf(full_html) -> bytes     -> PDF (WeasyPrint)  [optional dependency]
    to_doc(full_html) -> bytes     -> Word-openable .doc (HTML + Office namespace)

The render functions take a plain dict `f` (the "form"); see each function's
docstring for the expected keys. Helpers mirror the original JS exactly
(esc/fmt/fmtdate/num_to_words/amount_words/sum_by_type/by_type_table/parse_*).
"""

import os
import re
import math
import base64

# ── logos (embedded as data URIs, read from assets/) ───────────────────────────

_ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")


def _logo_uri(fname):
    path = os.path.join(_ASSETS, fname)
    try:
        with open(path, "rb") as fh:
            b64 = base64.b64encode(fh.read()).decode("ascii")
        return "data:image/png;base64," + b64
    except Exception:
        return ""  # render still works, just without the logo


BME_LOGO = _logo_uri("bme_logo.png")
BCW_LOGO = _logo_uri("bcw_logo.png")

# ── company config (mirrors COMPANIES in the HTML app) ──────────────────────────

COMPANIES = {
    "BME": {
        "id": "BME", "name": "Brescia Middle East S.a.l.", "prefix": "BME",
        "currency": "USD", "symbol": "$", "curWord": "United States Dollars",
        "style": "bme", "logo": BME_LOGO,
        "seller": [
            "Brescia Middle East S.a.l.",
            "Lebanon - jounieh main road rizk building",
            "Phone : 9619636896",
            "info@bresciame.com - www.bresciame.com",
            "V.A.T. Reg. 3194673-601",
        ],
        "shipper": ("Brescia Middle East S.a.l.\n"
                    "Lebanon - Jounieh main road, Rizk building\n"
                    "Tel: 9619636896   info@bresciame.com"),
        "bank": "",
    },
    "BCW": {
        "id": "BCW", "name": "BCW - Brescia Custom Works S.R.L.", "prefix": "BCW",
        "currency": "EUR", "symbol": "€", "curWord": "Euro",
        "style": "bcw", "logo": BCW_LOGO,
        "seller": [
            "BCW- BRESCIA CUSTOM WORKS",
            "Via Matteotti 311 - Gardone Val Trompia",
            "25063 (BS) - Italia - P.IVA-C.F 04425720986",
            "Tel: +390307285583   Email: info@bresciacw.com",
        ],
        "shipper": ("Brescia Custom Works srl\n"
                    "Via Matteotti 311\n"
                    "25063 Gardone Val Trompia (Brescia)\n"
                    "Italy"),
        "bank": ("BCW -BRESCIA CUSTOM WORKS S.R.L\n"
                 "BANCA CASSA PADANA FILIALE DI MARCHENO\n"
                 "IBAN : IT28 W083 4054 7400 0000 2137 960\n"
                 "BIC : CCRTIT2TPAD"),
    },
}

# ── helpers (faithful ports) ────────────────────────────────────────────────────


def esc(s):
    if s is None:
        s = ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def fmt(n):
    n = round(_num(n) * 100) / 100
    return f"{n:,.2f}"  # en-US thousands + 2 decimals, e.g. 1,234.50


def fmtdate(iso):
    if not iso:
        return ""
    p = str(iso).split("-")
    if len(p) == 3:
        return f"{int(p[1])}/{int(p[2])}/{p[0]}"  # M/D/YYYY
    return str(iso)


def num_to_words(n):
    n = int(n)
    if n == 0:
        return "zero"
    ones = ["", "one", "two", "three", "four", "five", "six", "seven", "eight",
            "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen",
            "sixteen", "seventeen", "eighteen", "nineteen"]
    tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy",
            "eighty", "ninety"]

    def below1000(x):
        s = ""
        if x >= 100:
            s += ones[x // 100] + " hundred"
            x %= 100
            if x:
                s += " "
        if x >= 20:
            s += tens[x // 10]
            x %= 10
            if x:
                s += "-" + ones[x]
        elif x > 0:
            s += ones[x]
        return s

    scales = ["", "thousand", "million", "billion"]
    chunks = []
    num = n
    while num > 0:
        chunks.append(num % 1000)
        num //= 1000
    parts = []
    for j in range(len(chunks) - 1, -1, -1):
        if chunks[j] == 0:
            continue
        seg = below1000(chunks[j])
        if scales[j]:
            seg += " " + scales[j]
        parts.append(seg)
    return " ".join(parts)


def amount_words(amount, cur_word):
    amount = _num(amount)
    whole = int(math.floor(amount))
    cents = int(round((amount - whole) * 100))
    w = num_to_words(whole)
    cap = w[:1].upper() + w[1:]
    return f"Only {cap} and {str(cents).zfill(2)}/100 {cur_word}"


def sum_by_type(items, type_key):
    m = {}
    order = []
    tot = 0.0
    for v in items:
        t = (str(v.get(type_key, "") or "").strip()) or "—"
        q = _num(v.get("qty"))
        if t not in m:
            m[t] = 0.0
            order.append(t)
        m[t] += q
        tot += q
    return [(t, m[t]) for t in order], tot


def _qnum(q):
    """Render a quantity the way JS would (no trailing .0 for whole numbers)."""
    if q == int(q):
        return str(int(q))
    return str(q)


def by_type_table(items, type_key, head_label):
    rows, total = sum_by_type(items, type_key)
    if not total:
        return ""
    trs = "".join(
        f'<tr><td style="border:1px solid #555;padding:5px 9px">{esc(t)}</td>'
        f'<td class="ctr" style="border:1px solid #555;padding:5px 9px">{_qnum(q)}</td></tr>'
        for t, q in rows
    )
    return (
        '<table style="width:auto;min-width:300px;margin-top:14px;border-collapse:collapse;font-size:12px">\n'
        '    <thead><tr><th style="background:#000;color:#fff;padding:6px 9px;text-align:left;border:1px solid #000">'
        f'{esc(head_label)}</th><th style="background:#000;color:#fff;padding:6px 9px;border:1px solid #000">QTY</th></tr></thead>\n'
        f'    <tbody>{trs}</tbody>\n'
        '    <tfoot><tr><td style="border:1px solid #555;padding:5px 9px;font-weight:700">TOTAL NO. OF GUNS</td>'
        f'<td class="ctr" style="border:1px solid #555;padding:5px 9px;font-weight:700">{_qnum(total)}</td></tr></tfoot>\n'
        '  </table>'
    )


def parse_weight_kg(s):
    m = re.search(r"([0-9]+(?:[.,][0-9]+)?)", re.sub(r"\s", "", str(s or "")))
    return float(m.group(1).replace(",", ".")) if m else 0.0


def parse_dims_cm(s):
    t = re.sub(r"\s", "", str(s or "").upper())
    t = t.replace("CM", "")
    t = re.sub(r"MT|M(?=X|$)", "", t)
    parts = []
    for x in t.split("X"):
        try:
            parts.append(float(x.replace(",", ".")))
        except ValueError:
            pass
    return parts[:3] if len(parts) >= 3 else None


def _seller_lines(co):
    return "".join(
        f'<div class="nm">{esc(l)}</div>' if i == 0 else f"<div>{esc(l)}</div>"
        for i, l in enumerate(co["seller"])
    )


def _doc_total(f):
    if f.get("total") is not None:
        return _num(f["total"])
    return sum(_num(v.get("qty")) * _num(v.get("price")) for v in f.get("items", []))


# ── renderers ───────────────────────────────────────────────────────────────────


def render_bme_invoice(f, co):
    items = f.get("items", [])
    has_serial = any((str(v.get("serial", "") or "").strip() != "") for v in items)
    total = _doc_total(f)
    rows = []
    for v in items:
        line = _num(v.get("qty")) * _num(v.get("price"))
        serial_td = f'<td class="ctr">{esc(v.get("serial"))}</td>' if has_serial else ""
        rows.append(
            "<tr>\n"
            f'      <td>{esc(v.get("type"))}</td><td class="ctr">{esc(v.get("gauge"))}</td><td class="ctr">{esc(v.get("brand"))}</td>\n'
            f'      {serial_td}\n'
            f'      <td class="ctr">{esc(v.get("qty"))}</td>\n'
            f'      <td><table style="width:100%"><tr><td style="border:none;padding:0">{co["symbol"]}</td><td class="num" style="border:none;padding:0">{fmt(v.get("price"))}</td></tr></table></td>\n'
            f'      <td><table style="width:100%"><tr><td style="border:none;padding:0">{co["symbol"]}</td><td class="num" style="border:none;padding:0">{fmt(line)}</td></tr></table></td>\n'
            "    </tr>"
        )
    rows = "".join(rows)
    buyer = f.get("buyer", {})
    b_trade = f'<div>{esc(buyer.get("trade"))}</div>' if buyer.get("trade") else ""
    b_addr = f'<div>{esc(buyer.get("addr"))}</div>' if buyer.get("addr") else ""
    b_country = f'<div>{esc(buyer.get("country"))}</div>' if buyer.get("country") else ""
    b_phone = f'<div>Phone : {esc(buyer.get("phone"))}</div>' if buyer.get("phone") else ""
    serial_th = "<th>Serial No.</th>" if has_serial else ""
    notes = f'<div style="font-size:12px;padding:6px 2px;white-space:pre-line">{esc(f.get("notes"))}</div>' if f.get("notes") else ""
    bank = f'<div style="font-size:12px;padding:6px 2px;white-space:pre-line"><b>Bank details</b><br>{esc(f.get("bank"))}</div>' if f.get("bank") else ""
    return (
        '<div class="doc bme">\n'
        '    <div class="head">\n'
        f'      <img class="logo" src="{co["logo"]}" style="width:200px">\n'
        '      <div style="min-width:300px">\n'
        '        <div class="pi-title">PROFORMA INVOICE</div>\n'
        f'        <div class="pi-field"><span></span><div class="pi-box">{esc(f.get("number"))}</div></div>\n'
        f'        <div class="pi-field"><span>Date</span><div class="pi-box">{esc(fmtdate(f.get("date")))}</div></div>\n'
        f'        <div class="pi-field"><span>Currency</span><div class="pi-box">{esc(f.get("currency"))}</div></div>\n'
        '      </div>\n'
        '    </div>\n'
        '    <div class="sb">\n'
        f'      <div><div class="blackbar">SELLER</div><div class="body">\n        {_seller_lines(co)}\n      </div></div>\n'
        '      <div><div class="blackbar">BUYER</div><div class="body">\n'
        f'        <div class="nm">{esc(buyer.get("name"))}</div>\n'
        f'        {b_trade}\n        {b_addr}\n        {b_country}\n        {b_phone}\n'
        '      </div></div>\n'
        '    </div>\n'
        '    <table class="it" style="margin-top:0">\n'
        f'      <thead><tr><th>Type</th><th>Gauge</th><th>Brand</th>{serial_th}<th>QTY</th><th>UNIT PRICE</th><th>TOTAL {esc(f.get("currency"))}</th></tr></thead>\n'
        f'      <tbody>{rows}</tbody>\n'
        '    </table>\n'
        '    <div class="totbox">\n'
        f'      <div class="words">{esc(amount_words(total, co["curWord"]))}</div>\n'
        '      <div class="totcol">\n'
        f'        <div class="tr"><div class="l">TOTAL</div><div class="v">{co["symbol"]} {fmt(total)}</div></div>\n'
        f'        <div class="tr"><div class="l">NET TOTAL {esc(f.get("currency"))}</div><div class="v">{co["symbol"]} {fmt(total)}</div></div>\n'
        '      </div>\n'
        '    </div>\n'
        '    <div class="terms">Terms &amp; Conditions</div>\n'
        f'    {notes}\n    {bank}\n'
        f'    {by_type_table(items, "type", "TYPE OF GUN")}\n'
        f'    <div class="signrow"><div class="sign"><img src="{co["logo"]}" style="width:150px;opacity:.85"><br>Authorized signature &amp; stamp</div></div>\n'
        '  </div>'
    )


def render_bcw_invoice(f, co):
    items = f.get("items", [])
    has_serial = any((str(v.get("serial", "") or "").strip() != "") for v in items)
    total = _doc_total(f)
    rows = []
    for v in items:
        line = _num(v.get("qty")) * _num(v.get("price"))
        serial_td = f'<td class="ctr">{esc(v.get("serial"))}</td>' if has_serial else ""
        rows.append(
            "<tr>\n"
            f'      <td class="ctr">{esc(v.get("tipo"))}</td><td class="ctr">{esc(v.get("cal"))}</td><td class="ctr">{esc(v.get("marca"))}</td>\n'
            f'      <td class="ctr">{esc(v.get("desc"))}</td>{serial_td}<td class="ctr">{esc(v.get("qty"))}</td>\n'
            f'      <td class="num">{fmt(v.get("price"))}</td><td class="num">{fmt(line)}</td>\n'
            "    </tr>"
        )
    rows = "".join(rows)
    buyer = f.get("buyer", {})
    b_trade = f'<div>{esc(buyer.get("trade"))}</div>' if buyer.get("trade") else ""
    b_addr = f'<div>{esc(buyer.get("addr"))}</div>' if buyer.get("addr") else ""
    b_country = f'<div>{esc(buyer.get("country"))}</div>' if buyer.get("country") else ""
    b_phone = f'<div>Tel: {esc(buyer.get("phone"))}</div>' if buyer.get("phone") else ""
    serial_th = "<th>matricola</th>" if has_serial else ""
    bank = f.get("bank") or co["bank"]
    bank_lines = "".join(f"<div>{esc(l)}</div>" for l in bank.split("\n"))
    notes = f'<div class="grey" style="white-space:pre-line">{esc(f.get("notes"))}</div>' if f.get("notes") else ""
    return (
        '<div class="doc bcw">\n'
        f'    <img class="logo" src="{co["logo"]}" style="width:230px">\n'
        f'    <div class="seller">\n      {_seller_lines(co)}\n    </div>\n'
        '    <div class="pi-title">PROFORMA INVOICE</div>\n'
        '    <div class="meta">\n'
        '      <div class="buyer">\n'
        f'        <div class="nm">{esc(buyer.get("name"))}</div>\n'
        f'        {b_trade}\n        {b_addr}\n        {b_country}\n        {b_phone}\n'
        '      </div>\n'
        f'      <div class="ref"><div><b>INVOICE #</b> {esc(f.get("number"))}</div><div>Date: {esc(fmtdate(f.get("date")))}</div></div>\n'
        '    </div>\n'
        '    <table class="it bcw" style="margin-top:10px">\n'
        f'      <thead><tr><th>tipo</th><th>cal</th><th>marca</th><th>descrizione</th>{serial_th}<th>QTA</th><th>PREZZO</th><th>TOTALE</th></tr></thead>\n'
        f'      <tbody>{rows}</tbody>\n'
        '    </table>\n'
        f'    <div class="toteuro"><div class="l">TOTAL {esc(f.get("currency"))}</div><div class="v">{co["symbol"]} {fmt(total)}</div></div>\n'
        f'    {by_type_table(items, "tipo", "TIPO ARMA / TYPE OF GUN")}\n'
        '    <div class="footer">\n'
        '      <div class="bankbox">\n'
        '        <div class="hd">COORDINATE BANCARIE</div>\n'
        f'        {bank_lines}\n'
        '      </div>\n'
        '      <div class="notes">\n'
        f'        {notes}\n'
        '        <div style="margin-top:10px">Thank you for your business!</div>\n'
        '      </div>\n'
        '    </div>\n'
        '  </div>'
    )


def render_proforma(f):
    """f: {company:'BME'|'BCW', number, date(ISO), currency, total(optional),
           buyer:{name,trade,addr,country,phone}, items:[...], notes, bank}
       BME items: type,gauge,brand,serial,qty,price
       BCW items: tipo,cal,marca,desc,serial,qty,price"""
    co = COMPANIES[f["company"]]
    f.setdefault("currency", co["currency"])
    return render_bme_invoice(f, co) if f["company"] == "BME" else render_bcw_invoice(f, co)


def render_packing(f):
    """f: {company, number, date, buyer{...}, shipper(optional), notes,
           parcels:[{items:[{qty,type,brand,model,caliber,serial1,serial2}], dims, weight}]}"""
    co = COMPANIES[f["company"]]
    parcels = f.get("parcels", [])
    buyer = f.get("buyer", {})
    buyer_block = "\n".join(
        x for x in [buyer.get("name"), buyer.get("trade"), buyer.get("addr"),
                    buyer.get("country"), buyer.get("phone")] if x
    )
    shipper = f.get("shipper") or co["shipper"]
    is_bme = f["company"] == "BME"
    if is_bme:
        pl_header = (
            f'<div class="head"><img class="logo" src="{co["logo"]}" style="width:190px">\n'
            '        <div style="min-width:280px"><div class="pi-title">PACKING LIST</div>\n'
            f'        <div class="pi-field"><span></span><div class="pi-box">{esc(f.get("number"))}</div></div>\n'
            f'        <div class="pi-field"><span>Date</span><div class="pi-box">{esc(fmtdate(f.get("date")))}</div></div></div></div>'
        )
    else:
        pl_header = (
            f'<img class="logo" src="{co["logo"]}" style="width:220px">\n'
            f'        <div class="seller">{_seller_lines(co)}</div>\n'
            '        <div class="pi-title">PACKING LIST</div>\n'
            f'        <div class="meta"><div class="buyer"></div><div class="ref"><div><b>PL #</b> {esc(f.get("number"))}</div><div>Date: {esc(fmtdate(f.get("date")))}</div></div></div>'
        )

    pages = []
    for idx, p in enumerate(parcels):
        rows = "".join(
            "<tr>\n"
            f'      <td class="ctr">{esc(v.get("qty"))}</td><td>{esc(v.get("type"))}</td><td class="ctr">{esc(v.get("brand"))}</td>\n'
            f'      <td class="ctr">{esc(v.get("model"))}</td><td class="ctr">{esc(v.get("caliber"))}</td>\n'
            f'      <td class="ctr">{esc(v.get("serial1"))}</td><td class="ctr">{esc(v.get("serial2"))}</td></tr>'
            for v in p.get("items", [])
        )
        pb = "pagebreak" if idx > 0 else ""
        tbl_cls = "" if is_bme else "bcw"
        pages.append(
            f'<div class="pl-parcel {pb}">\n'
            f'      {pl_header}\n'
            '      <div class="pl-cs">\n'
            f'        <div><div class="hd">CUSTOMER:</div>{esc(buyer_block)}</div>\n'
            f'        <div><div class="hd">SHIPPER:</div>{esc(shipper)}</div>\n'
            '      </div>\n'
            f'      <div class="pl-pn">PARCEL NUMBER {idx + 1}</div>\n'
            f'      <table class="it {tbl_cls}" style="margin-top:0">\n'
            '        <thead><tr><th>Quantity</th><th>Type</th><th>Brand</th><th>Model</th><th>Caliber</th><th>Serial number 1</th><th>Serial number 2</th></tr></thead>\n'
            f'        <tbody>{rows}</tbody>\n'
            '      </table>\n'
            '      <div class="pl-dims">\n'
            f'        <div>DIMENSIONI PACCO:&nbsp; {esc(p.get("dims"))}</div>\n'
            f'        <div>PESO PACCO:&nbsp; {esc(p.get("weight"))}</div>\n'
            '      </div>\n'
            '    </div>'
        )
    pages = "".join(pages)

    # shipment summary
    all_items = []
    tot_w = 0.0
    tot_v = 0.0
    for p in parcels:
        for it in p.get("items", []):
            all_items.append(it)
        tot_w += parse_weight_kg(p.get("weight"))
        d = parse_dims_cm(p.get("dims"))
        if d:
            tot_v += d[0] * d[1] * d[2]
    pkgs = len(parcels)
    vol_m3 = tot_v / 1e6
    tot_guns = sum(_num(v.get("qty")) for v in all_items)
    pb = "pagebreak" if pkgs > 0 else ""
    tbl_cls = "" if is_bme else "bcw"
    s_notes = f'<div style="font-size:12px;padding:12px 2px;white-space:pre-line">{esc(f.get("notes"))}</div>' if f.get("notes") else ""
    summary = (
        f'<div class="pl-parcel {pb}">\n'
        f'    {pl_header}\n'
        '    <div class="pl-cs">\n'
        f'      <div><div class="hd">CUSTOMER:</div>{esc(buyer_block)}</div>\n'
        f'      <div><div class="hd">SHIPPER:</div>{esc(shipper)}</div>\n'
        '    </div>\n'
        '    <div class="pl-pn">SHIPMENT SUMMARY</div>\n'
        f'    <table class="it {tbl_cls}" style="margin-top:0;width:auto;min-width:360px">\n'
        '      <tbody>\n'
        f'        <tr><td style="font-weight:700">Total number of packages</td><td class="ctr">{pkgs}</td></tr>\n'
        f'        <tr><td style="font-weight:700">Total volume</td><td class="ctr">{vol_m3:.3f} m³ &nbsp;({fmt(tot_v / 1000)} dm³)</td></tr>\n'
        f'        <tr><td style="font-weight:700">Total gross weight</td><td class="ctr">{fmt(tot_w)} KG</td></tr>\n'
        f'        <tr><td style="font-weight:700">Total number of guns</td><td class="ctr">{_qnum(tot_guns)}</td></tr>\n'
        '      </tbody>\n'
        '    </table>\n'
        f'    {by_type_table(all_items, "type", "TYPE OF GUN")}\n'
        f'    {s_notes}\n'
        '  </div>'
    )
    return f'<div class="doc {co["style"]}">\n    {pages}\n    {summary}\n  </div>'


def _bcw_letterhead(co):
    return (
        f'<img class="logo" src="{co["logo"]}" style="width:220px">\n'
        f'    <div class="seller">{_seller_lines(co)}</div>'
    )


def render_forms(f):
    """EUR1 + End User Certificate (BCW). f: {company:'BCW', buyer{...},
       ef:{rep,sender,seat,invNo,invDate,dest,place,signDate,contract,commodity}}"""
    co = COMPANIES[f["company"]]
    e = f.get("ef", {})
    buyer = f.get("buyer", {})
    buyer_str = ", ".join(
        x for x in [buyer.get("name"), buyer.get("trade"), buyer.get("addr"),
                    buyer.get("country"), buyer.get("phone")] if x
    )
    inv_date = fmtdate(e.get("invDate")) if e.get("invDate") else "…"
    sign_bit = (", " + esc(fmtdate(e.get("signDate")))) if e.get("signDate") else ""
    eur1 = (
        '<div class="doc bcw ef-doc">\n'
        f'    {_bcw_letterhead(co)}\n'
        '    <div class="ef-h1">DICHIARAZIONE DI ORIGINE MERCE</div>\n'
        '    <div class="ef-h2">LETTERA DI INCARICO PER LA RICHIESTA<br>DI CERTIFICATO DI CIRCOLAZIONE EUR1</div>\n'
        f'    <p>Il sottoscritto <b>{esc(e.get("rep") or "…")}</b> nella qualità di legale rappresentante della ditta <b>{esc(e.get("sender") or co["name"])}</b> con sede in <b>{esc(e.get("seat") or "")}</b> consapevole della responsabilità e degli obblighi stabiliti dalla vigente normativa comunitaria e nazionale</p>\n'
        '    <div class="ef-decl">DICHIARA</div>\n'
        f'    <p>Che le merci meglio descritte nella fattura di esportazione nr <b>{esc(e.get("invNo") or "…")}</b> del <b>{esc(inv_date)}</b> soddisfano le condizioni richieste per ottenere il certificato EUR1</p>\n'
        f'    <p>In particolare dichiara che le merci di cui sopra sono di origine preferenziale comunitaria in base a quanto previsto negli accordi tra U.E. e <b>{esc(e.get("dest") or "…")}</b></p>\n'
        '    <p>A riscontro delle condizioni sopra dichiarate, oltre alla documentazione prodotta contestualmente alla domanda di rilascio del certificato EUR1, si impegna espressamente a fornire all\'Autorità Doganale qualsiasi altra prova documentale o giustificazione che quest\'ultima richiede, nonché ad accettare ogni eventuale controllo.</p>\n'
        '    <p>Per quanto sopra con la presente, si conferisce espresso incarico a formulare alla Dogana, domanda di rilascio del certificato EUR1 in relazione alle merci di cui sopra, alla Società BS CARGO SCS SRL</p>\n'
        '    <p>La Società BS CARGO SCS ed il doganalista MASSIMO TURINELLI vengono autorizzati a compiere tutto quanto necessario per l\'ottenimento del certificato EUR1 e sono fin d\'ora espressamente manlevati da qualsiasi responsabilità legale direttamente ed indirettamente all\'espletamento della procedura oggetto del presente incarico</p>\n'
        f'    <p style="margin-top:24px">Luogo e data: <b>{esc(e.get("place") or "")}{sign_bit}</b></p>\n'
        '    <div class="ef-sign">\n'
        f'      <div>Nome e cognome di chi firma<br><b>{esc(e.get("rep") or "")}</b></div>\n'
        '      <div>Timbro e firma</div>\n'
        '    </div>\n'
        '    <p style="margin-top:18px;font-size:11px;font-style:italic">Allegare documento di identità</p>\n'
        '  </div>'
    )
    contract_bit = (f' / Contract No. <b>{esc(e.get("contract"))}</b>'
                    if e.get("contract") else " / Contract No. …")
    euc = (
        '<div class="doc bcw ef-doc pagebreak">\n'
        f'    {_bcw_letterhead(co)}\n'
        '    <div class="ef-h1">END USER CERTIFICATE</div>\n'
        f'    <p>In accordance with the invoice No. <b>{esc(e.get("invNo") or "…")}</b> of <b>{esc(inv_date)}</b>{contract_bit}</p>\n'
        f'    <p><b>End user / buyer:</b> {esc(buyer_str or "…")}</p>\n'
        f'    <p><b>Detailed commodity description:</b> {esc(e.get("commodity") or "…")}</p>\n'
        f'    <p>will remain in <b>{esc(e.get("dest") or "…")}</b> and will not be used for human rights violations such as torture, slavery, cruel and inhuman punishment.</p>\n'
        '    <p>In addition, re-export of goods is definitely excluded. As well as will not be transferred to the Russian Federation nor Belarus, in accordance with the previsions of Article 12-octies of Regulation (EU) No.833/2014, as introduced by Regulation (EU) No.2023/2878 and Reg.UE 1865/2024 Article 8-octies.</p>\n'
        '    <p>The goods may be transferred to a third party/company only on the condition that this third party/company accepts the obligations contained in the above declaration and that this third party/company is known to be reliable and law abiding.</p>\n'
        '    <p>The goods will only be sold to a third party provided that this third party has all necessary legal documents in accordance with local law and we, as the importer of the goods, are obliged to verify this.</p>\n'
        f'    <p>If <b>{esc(e.get("sender") or co["name"])}</b> becomes aware that the buyer has violated the contractual prohibitions, it is obliged to inform the Italian competent authority immediately, in order ensure the prompt adoption of appropriate measures. A proven infringement of this prohibition precludes the possibility of further contracts with the buyer.</p>\n'
        f'    <p style="margin-top:24px">Place &amp; date: <b>{esc(e.get("place") or "")}{sign_bit}</b></p>\n'
        '    <div class="ef-sign">\n'
        f'      <div>Name and signature<br><b>{esc(e.get("rep") or "")}</b></div>\n'
        '      <div>Stamp and signature</div>\n'
        '    </div>\n'
        '  </div>'
    )
    return eur1 + euc


# ── document CSS (verbatim from the HTML app, .doc rules + print) ────────────────

DOC_CSS = """
.doc{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff;width:100%;padding:34px 38px;font-size:12.5px;box-sizing:border-box}
  .doc *{box-sizing:border-box}
  .doc img.logo{display:block}
  .doc .blackbar{background:#000;color:#fff;font-weight:700;font-size:11px;letter-spacing:.5px;padding:3px 8px}
  .doc table{width:100%;border-collapse:collapse}
  .doc .it th{background:#000;color:#fff;font-size:11px;letter-spacing:.4px;padding:7px 8px;text-align:center;border:1px solid #000}
  .doc .it td{border:1px solid #555;padding:7px 9px;font-size:12.5px}
  .doc .num{text-align:right;font-weight:700;font-variant-numeric:tabular-nums;white-space:nowrap}
  .doc .ctr{text-align:center}
  /* BME */
  .doc.bme .head{display:flex;justify-content:space-between;align-items:flex-start}
  .doc.bme .pi-title{font-weight:700;font-size:15px;text-align:right;margin-bottom:6px}
  .doc.bme .pi-field{display:flex;align-items:center;justify-content:flex-end;gap:10px;margin-bottom:6px}
  .doc.bme .pi-field span{font-size:12px}
  .doc.bme .pi-box{border:1px solid #000;min-width:150px;text-align:center;padding:3px 6px;font-weight:700}
  .doc.bme .sb{display:flex;gap:0;margin-top:14px}
  .doc.bme .sb > div{flex:1}
  .doc.bme .sb .body{padding:6px 8px;font-size:12px;line-height:1.5}
  .doc.bme .sb .body .nm{font-weight:700}
  .doc.bme .totbox{display:flex;margin-top:0}
  .doc.bme .words{flex:1;border:1px solid #000;border-top:none;padding:8px 10px;font-size:12px;display:flex;align-items:center}
  .doc.bme .totcol{width:300px}
  .doc.bme .totcol .tr{display:flex;border:1px solid #000;border-top:none}
  .doc.bme .totcol .tr .l{flex:1;font-weight:700;padding:5px 8px;border-right:1px solid #000}
  .doc.bme .totcol .tr .v{width:150px;text-align:right;font-weight:700;padding:5px 10px;font-variant-numeric:tabular-nums}
  .doc.bme .terms{background:#000;color:#fff;font-weight:700;font-size:11px;padding:3px 8px;margin-top:14px}
  .doc.bme .signrow{display:flex;justify-content:flex-end;margin-top:26px}
  .doc.bme .sign{text-align:center;font-size:11px;color:#333}
  /* BCW */
  .doc.bcw .seller{font-size:12px;line-height:1.5;margin-top:6px}
  .doc.bcw .seller .nm{font-weight:700}
  .doc.bcw .pi-title{text-align:center;font-weight:700;font-size:22px;margin:16px 0 10px}
  .doc.bcw .meta{display:flex;justify-content:space-between;align-items:flex-start;font-size:12.5px;line-height:1.55}
  .doc.bcw .meta .buyer .nm{font-weight:700}
  .doc.bcw .meta .ref{text-align:right}
  .doc.bcw .it.bcw th{background:#fff;color:#000;border:1px solid #000;font-weight:700;font-size:12px}
  .doc.bcw .it.bcw td{border:1px solid #000}
  .doc.bcw .toteuro{display:flex;justify-content:flex-end;margin-top:0}
  .doc.bcw .toteuro .l{border:1px solid #000;border-top:none;padding:6px 12px;font-weight:700}
  .doc.bcw .toteuro .v{border:1px solid #000;border-top:none;border-left:none;width:170px;text-align:right;padding:6px 12px;font-weight:700;font-variant-numeric:tabular-nums}
  .doc.bcw .footer{display:flex;justify-content:space-between;gap:20px;margin-top:22px;font-size:12px}
  .doc.bcw .bankbox{border:1px solid #000;padding:8px 10px;line-height:1.55;max-width:340px}
  .doc.bcw .bankbox .hd{font-weight:700}
  .doc.bcw .notes{text-align:right;line-height:1.7}
  .doc.bcw .notes .grey{background:#eee;padding:6px 10px;display:inline-block;text-align:center}
  /* Packing list */
  .doc .pl-parcel{margin-top:16px}
  .doc .pl-cs{display:flex;gap:0;margin-top:12px;border:1px solid #000}
  .doc .pl-cs > div{flex:1;padding:8px 10px;font-size:11.5px;line-height:1.5;white-space:pre-line}
  .doc .pl-cs > div:first-child{border-right:1px solid #000}
  .doc .pl-cs .hd{font-weight:700;margin-bottom:3px}
  .doc .pl-pn{background:#000;color:#fff;font-weight:700;font-size:12px;letter-spacing:.5px;padding:4px 10px;margin-top:14px}
  .doc .pl-dims{display:flex;gap:0;margin-top:0;border:1px solid #000;border-top:none}
  .doc .pl-dims > div{flex:1;padding:6px 10px;font-weight:700;font-size:11.5px}
  .doc .pl-dims > div:first-child{border-right:1px solid #000}
  /* export forms (EUR1 / EUC) */
  .doc.ef-doc{font-size:12.5px;line-height:1.55;color:#111}
  .doc.ef-doc p{margin:9px 0;text-align:justify}
  .doc.ef-doc .ef-h1{font-weight:800;font-size:16px;text-align:center;margin:18px 0 6px;letter-spacing:.5px}
  .doc.ef-doc .ef-h2{font-weight:700;font-size:13px;text-align:center;margin:0 0 16px}
  .doc.ef-doc .ef-decl{font-weight:800;font-size:14px;text-align:center;letter-spacing:3px;margin:14px 0}
  .doc.ef-doc .ef-sign{display:flex;justify-content:space-between;margin-top:46px;font-size:12px}
  .doc.ef-doc .ef-sign > div{width:45%}
  .doc .pagebreak,.doc.pagebreak{page-break-before:always}
  @media print{
    body{background:#fff;margin:0}
    .doc{padding:14mm 12mm}
    @page{size:A4;margin:0}
  }
"""


def document_html(body, title="Document"):
    """Wrap a render output into a complete standalone HTML page."""
    return (
        "<!DOCTYPE html>\n<html><head><meta charset=\"utf-8\">\n"
        f"<title>{esc(title)}</title>\n<style>{DOC_CSS}</style>\n"
        "</head><body>\n" + body + "\n</body></html>"
    )


def to_pdf(full_html):
    """Render a full HTML page to PDF bytes via WeasyPrint (optional dependency)."""
    from weasyprint import HTML  # imported lazily so the app loads without it
    return HTML(string=full_html).write_pdf()


def to_doc(full_html):
    """Return Word-openable .doc bytes (HTML body + Office namespace header).
    Word opens HTML-with-mso just like the original app's Word export."""
    head = (
        '<html xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:w="urn:schemas-microsoft-com:office:word" '
        'xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8">'
        "<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View>"
        "<w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]-->"
        f"<style>{DOC_CSS}</style></head><body>"
    )
    m = re.search(r"<body>(.*)</body>", full_html, re.S)
    body = m.group(1) if m else full_html
    return (head + body + "</body></html>").encode("utf-8")
