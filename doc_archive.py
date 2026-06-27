"""
Generated-document archive + ledger.
====================================
Every proforma / packing list / declaration the app generates is saved here so
it can be looked up, re-downloaded and audited later. Two things are stored in
the SAME shared backend as the register (storage.py):

  * the rendered files   ->  documents/files/<id>__<name>.pdf / .html
  * a searchable ledger  ->  documents/ledger.json

The ledger is the "log chart": one row per issued document, with its number,
company, type, date, buyer, total and the colleague who issued it. Because it
lives in shared storage and is written under the backend lock, the log stays
consistent even when several people generate documents at the same time.

API:
    arch = DocArchive(backend)
    entry = arch.add(base_name, html, pdf_bytes, meta={...})
    rows  = arch.list(query=None, company=None, doc_type=None)
    data, fname = arch.get_file(entry_id, kind="pdf")   # or "html"
    arch.delete(entry_id)

Entry schema:
    id, number, company, doc_type, buyer, total, currency, doc_date,
    user, issued_at, base_name, pdf_key, html_key, size
"""

import json
import uuid
import contextlib
from datetime import datetime, timezone

LEDGER_KEY = "documents/ledger.json"
FILES_PREFIX = "documents/files/"

# human labels for the document kinds we archive
DOC_TYPE_LABELS = {
    "proforma": "Proforma invoice",
    "packing": "Packing list",
    "declaration": "Dichiarazione (EUR1 + EUC)",
}


class DocArchive:
    def __init__(self, backend):
        self.backend = backend

    # ---- ledger io ----
    def _load(self):
        if not self.backend.exists(LEDGER_KEY):
            return []
        try:
            return json.loads(self.backend.read_bytes(LEDGER_KEY).decode("utf-8"))
        except Exception:
            return []

    def _save(self, entries):
        self.backend.write_bytes(
            LEDGER_KEY,
            json.dumps(entries, ensure_ascii=False, indent=2).encode("utf-8"),
        )

    # ---- crud ----
    def add(self, base_name, html, pdf_bytes=None, meta=None):
        """Archive one generated document. `html` is the full document HTML
        (always stored); `pdf_bytes` is optional (stored if WeasyPrint produced
        a PDF). Returns the ledger entry."""
        meta = meta or {}
        entry_id = uuid.uuid4().hex[:12]
        safe = (base_name or "documento").replace("/", "_").replace("\\", "_").strip()

        html_key = f"{FILES_PREFIX}{entry_id}__{safe}.html"
        html_bytes = html.encode("utf-8") if isinstance(html, str) else html
        self.backend.write_bytes(html_key, html_bytes)

        pdf_key = ""
        size = len(html_bytes)
        if pdf_bytes:
            pdf_key = f"{FILES_PREFIX}{entry_id}__{safe}.pdf"
            self.backend.write_bytes(pdf_key, pdf_bytes)
            size = len(pdf_bytes)

        entry = {
            "id": entry_id,
            "number": meta.get("number", ""),
            "company": meta.get("company", ""),
            "doc_type": meta.get("doc_type", ""),
            "buyer": meta.get("buyer", ""),
            "total": meta.get("total", ""),
            "currency": meta.get("currency", ""),
            "doc_date": meta.get("doc_date", ""),
            "user": meta.get("user", ""),
            "issued_at": datetime.now(timezone.utc).isoformat(),
            "base_name": safe,
            "pdf_key": pdf_key,
            "html_key": html_key,
            "size": size,
        }

        owner = meta.get("user") or "app"
        locked = False
        with contextlib.suppress(Exception):
            self.backend.acquire_lock(LEDGER_KEY, owner)
            locked = True
        try:
            entries = self._load()
            entries.append(entry)
            self._save(entries)
        finally:
            if locked:
                with contextlib.suppress(Exception):
                    self.backend.release_lock(LEDGER_KEY)
        return entry

    def list(self, query=None, company=None, doc_type=None):
        entries = self._load()
        if company:
            entries = [e for e in entries if e.get("company") == company]
        if doc_type:
            entries = [e for e in entries if e.get("doc_type") == doc_type]
        if query:
            q = query.lower()

            def hit(e):
                hay = " ".join(str(e.get(k, "")) for k in (
                    "number", "company", "doc_type", "buyer",
                    "doc_date", "user", "base_name")).lower()
                return q in hay

            entries = [e for e in entries if hit(e)]
        entries.sort(key=lambda e: e.get("issued_at", ""), reverse=True)
        return entries

    def get(self, entry_id):
        for e in self._load():
            if e["id"] == entry_id:
                return e
        return None

    def get_file(self, entry_id, kind="pdf"):
        """kind: 'pdf' or 'html'. Falls back to html if pdf wasn't stored."""
        e = self.get(entry_id)
        if not e:
            return None, None
        key = e.get("pdf_key") if kind == "pdf" else e.get("html_key")
        if not key:
            key = e.get("html_key")
            kind = "html"
        if not key:
            return None, None
        ext = "pdf" if kind == "pdf" else "html"
        return self.backend.read_bytes(key), f"{e.get('base_name', 'documento')}.{ext}"

    def delete(self, entry_id):
        entries = self._load()
        keep, removed = [], None
        for e in entries:
            if e["id"] == entry_id:
                removed = e
            else:
                keep.append(e)
        if removed:
            for k in (removed.get("pdf_key"), removed.get("html_key")):
                if k:
                    with contextlib.suppress(Exception):
                        self.backend.delete(k)
            self._save(keep)
            return True
        return False
