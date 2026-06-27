"""
Export / Import licence + EUC archive.
======================================
Stores licence and End-User-Certificate files in the same shared backend as the
register, with a searchable JSON index. Each entry can be tagged and linked to
an operation (N. OPERAZIONE) or an invoice number for later retrieval.

Index lives at  licenses/index.json ; files at  licenses/files/<id>__<name>.
Designed to use the same storage.Backend as the workbook, so everything is in
one durable, shared place.

API:
    arch = LicenseArchive(backend)
    arch.add(file_bytes, filename, meta={...})   -> entry dict
    arch.list(query=None, doc_type=None)         -> [entry, ...]
    arch.get_file(entry_id)                       -> (bytes, filename)
    arch.update(entry_id, meta)                   -> entry
    arch.delete(entry_id)                         -> bool

Entry schema:
    id, filename, doc_type, number, country, counterparty,
    issue_date, expiry_date, n_operazione, invoice_no, tags[], notes,
    uploaded_at, size, storage_key
"""

import json
import uuid
import contextlib
from datetime import datetime, timezone

INDEX_KEY = "licenses/index.json"
FILES_PREFIX = "licenses/files/"

DOC_TYPES = [
    "Licenza Esportazione",      # export licence
    "Licenza Importazione",      # import licence
    "EUC (End User Certificate)",
    "Autorizzazione UAMA",
    "Transito / Brokering",
    "Altro",
]


class LicenseArchive:
    def __init__(self, backend):
        self.backend = backend

    # ---- index io ----
    def _load_index(self):
        if not self.backend.exists(INDEX_KEY):
            return []
        try:
            return json.loads(self.backend.read_bytes(INDEX_KEY).decode("utf-8"))
        except Exception:
            return []

    def _save_index(self, entries):
        self.backend.write_bytes(
            INDEX_KEY, json.dumps(entries, ensure_ascii=False, indent=2).encode("utf-8"))

    # ---- crud ----
    def add(self, file_bytes, filename, meta=None):
        meta = meta or {}
        entry_id = uuid.uuid4().hex[:12]
        safe_name = filename.replace("/", "_").replace("\\", "_")
        storage_key = f"{FILES_PREFIX}{entry_id}__{safe_name}"
        self.backend.write_bytes(storage_key, file_bytes)
        entry = {
            "id": entry_id,
            "filename": safe_name,
            "doc_type": meta.get("doc_type", "Altro"),
            "number": meta.get("number", ""),
            "country": meta.get("country", ""),
            "counterparty": meta.get("counterparty", ""),
            "issue_date": meta.get("issue_date", ""),
            "expiry_date": meta.get("expiry_date", ""),
            "n_operazione": meta.get("n_operazione", ""),
            "invoice_no": meta.get("invoice_no", ""),
            "tags": meta.get("tags", []),
            "notes": meta.get("notes", ""),
            "uploaded_at": datetime.now(timezone.utc).isoformat(),
            "size": len(file_bytes),
            "storage_key": storage_key,
        }
        entries = self._load_index()
        entries.append(entry)
        self._save_index(entries)
        return entry

    def list(self, query=None, doc_type=None):
        entries = self._load_index()
        if doc_type:
            entries = [e for e in entries if e.get("doc_type") == doc_type]
        if query:
            q = query.lower()
            def hit(e):
                hay = " ".join(str(e.get(k, "")) for k in (
                    "filename", "number", "country", "counterparty",
                    "n_operazione", "invoice_no", "notes")).lower()
                hay += " " + " ".join(str(t) for t in e.get("tags", [])).lower()
                return q in hay
            entries = [e for e in entries if hit(e)]
        entries.sort(key=lambda e: e.get("uploaded_at", ""), reverse=True)
        return entries

    def get(self, entry_id):
        for e in self._load_index():
            if e["id"] == entry_id:
                return e
        return None

    def get_file(self, entry_id):
        e = self.get(entry_id)
        if not e:
            return None, None
        return self.backend.read_bytes(e["storage_key"]), e["filename"]

    def update(self, entry_id, meta):
        entries = self._load_index()
        for e in entries:
            if e["id"] == entry_id:
                e.update({k: v for k, v in meta.items() if k in e})
                self._save_index(entries)
                return e
        return None

    def delete(self, entry_id):
        entries = self._load_index()
        keep, removed = [], None
        for e in entries:
            if e["id"] == entry_id:
                removed = e
            else:
                keep.append(e)
        if removed:
            with contextlib.suppress(Exception):
                self.backend.delete(removed["storage_key"])
            self._save_index(keep)
            return True
        return False

    def expiring_soon(self, days=60):
        """Entries whose expiry_date is within `days` (or already past)."""
        out = []
        today = datetime.now().date()
        for e in self._load_index():
            d = e.get("expiry_date")
            if not d:
                continue
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    exp = datetime.strptime(d, fmt).date()
                    delta = (exp - today).days
                    if delta <= days:
                        ee = dict(e); ee["days_left"] = delta
                        out.append(ee)
                    break
                except ValueError:
                    continue
        out.sort(key=lambda e: e.get("days_left", 9999))
        return out
