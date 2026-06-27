"""
Central gapless document numbering for the Brescia document generator.
=====================================================================
Official numbered documents (proforma invoices, packing lists) MUST get a
unique, gapless, per-company / per-year sequence even when several colleagues
generate documents at the same time. A per-browser counter (the old HTML app's
localStorage) cannot guarantee that — two people would mint the same number.

So counters live in the SAME durable shared storage backend as the MAGAZZINO
workbook (see storage.py), and every allocation is done under the backend lock.
That makes "give me the next number" atomic across all sessions/devices.

Format (identical to the original HTML app):
    proforma   ->  BME-2026-001   / BCW-2026-001
    packing    ->  BME-PL-2026-001 / BCW-PL-2026-001
Sequence resets to 001 at the start of each year, independently per company and
per document type.

Export forms (EUR1 + End User Certificate) do NOT consume a number — they are
issued against an already-existing final-invoice number, so they are not handled
here.

Usage:
    import storage, doc_numbering
    backend = storage.get_backend()
    num = doc_numbering.next_number(backend, "BME", "proforma")   # -> "BME-2026-001"
"""

import json
from datetime import date

COUNTERS_KEY = "documents/counters.json"

DOC_TYPES = ("proforma", "packing")


def _load(backend):
    if backend.exists(COUNTERS_KEY):
        try:
            return json.loads(backend.read_bytes(COUNTERS_KEY).decode("utf-8"))
        except Exception:
            return {}
    return {}


def _save(backend, counters):
    backend.write_bytes(
        COUNTERS_KEY,
        json.dumps(counters, indent=2, ensure_ascii=False).encode("utf-8"),
    )


def format_number(company, doc_type, year, seq):
    s = f"{int(seq):03d}"
    if doc_type == "packing":
        return f"{company}-PL-{year}-{s}"
    return f"{company}-{year}-{s}"


def _key(company, doc_type, year):
    return f"{company}|{doc_type}|{year}"


def next_number(backend, company, doc_type, year=None, owner="app"):
    """Atomically allocate and persist the next number. Returns the formatted id."""
    if doc_type not in DOC_TYPES:
        raise ValueError(f"doc_type sconosciuto: {doc_type} (atteso uno di {DOC_TYPES})")
    year = year or date.today().year
    backend.acquire_lock(COUNTERS_KEY, owner)
    try:
        counters = _load(backend)
        key = _key(company, doc_type, year)
        seq = int(counters.get(key, 0)) + 1
        counters[key] = seq
        _save(backend, counters)
        return format_number(company, doc_type, year, seq)
    finally:
        backend.release_lock(COUNTERS_KEY)


def peek(backend, company, doc_type, year=None):
    """Return the LAST issued sequence (0 if none). Does not allocate."""
    year = year or date.today().year
    counters = _load(backend)
    return int(counters.get(_key(company, doc_type, year), 0))


def preview_next(backend, company, doc_type, year=None):
    """What the next number WOULD be, without consuming it (for UI preview)."""
    year = year or date.today().year
    return format_number(company, doc_type, year, peek(backend, company, doc_type, year) + 1)


def list_counters(backend):
    """All counters as {company|doc_type|year: last_seq} — for an admin view."""
    return _load(backend)


def set_counter(backend, company, doc_type, year, seq, owner="app"):
    """Manual override (e.g. to align with numbers already issued on paper)."""
    backend.acquire_lock(COUNTERS_KEY, owner)
    try:
        counters = _load(backend)
        counters[_key(company, doc_type, year)] = int(seq)
        _save(backend, counters)
    finally:
        backend.release_lock(COUNTERS_KEY)
