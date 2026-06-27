"""
Shared-storage backend for the BCW register.
============================================
Streamlit Cloud's local disk is EPHEMERAL — it is wiped on every restart,
sleep, or redeploy. A legal register must therefore live in durable shared
storage that every colleague's session reads from and writes to.

This module gives the app a single authoritative workbook plus a lock so two
people can never save at the same time and corrupt it.

Supported backends (pick one via env var / Streamlit secret  BCW_STORAGE):
    local      -> a path on disk            (dev / single machine only)
    s3         -> any S3-compatible bucket  (AWS S3, Backblaze B2, Cloudflare R2, MinIO)
    dropbox    -> a Dropbox app folder

Typical app usage:

    import storage, stock_agent
    backend = storage.get_backend()
    with storage.workbook_session(backend) as local_path:
        stock_agent.EXCEL_FILE = local_path
        stock_agent.add_carico(...)        # mutates the local copy
    # on clean exit the workbook is uploaded back and the lock released

Config keys (env vars OR st.secrets):
    BCW_STORAGE                local | s3 | dropbox
    BCW_WORKBOOK_NAME          object/key/filename of the workbook (default "MAGAZZINO BCW V45.xlsx")
    # local:
    BCW_LOCAL_DIR
    # s3:
    BCW_S3_BUCKET, BCW_S3_PREFIX, AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY,
    AWS_REGION (optional), BCW_S3_ENDPOINT (optional, for B2/R2/MinIO)
    # dropbox:
    BCW_DROPBOX_TOKEN (or refresh-token trio), BCW_DROPBOX_DIR
"""

import os
import io
import time
import json
import tempfile
import contextlib
from datetime import datetime, timezone

DEFAULT_WORKBOOK = "MAGAZZINO BCW fixed.xlsx"
LOCK_STALE_SECONDS = 120          # auto-break a lock older than this
LOCK_WAIT_TIMEOUT = 30            # how long to wait for a busy lock
LOCK_POLL = 1.0


def _cfg(key, default=None):
    """Read from Streamlit secrets first, then env."""
    try:
        import streamlit as st
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key, default)


# ── Backend base ──────────────────────────────────────────────────────────────

class Backend:
    name = "base"

    def read_bytes(self, key):              raise NotImplementedError
    def write_bytes(self, key, data):       raise NotImplementedError
    def exists(self, key):                  raise NotImplementedError
    def delete(self, key):                  raise NotImplementedError

    # ---- lock (implemented in terms of read/write/exists/delete) ----
    def _lock_key(self, workbook):
        return workbook + ".lock"

    def acquire_lock(self, workbook, owner):
        lk = self._lock_key(workbook)
        deadline = time.time() + LOCK_WAIT_TIMEOUT
        while True:
            if self.exists(lk):
                try:
                    info = json.loads(self.read_bytes(lk).decode("utf-8"))
                    ts = datetime.fromisoformat(info["ts"])
                    age = (datetime.now(timezone.utc) - ts).total_seconds()
                except Exception:
                    age = LOCK_STALE_SECONDS + 1
                if age <= LOCK_STALE_SECONDS:
                    if time.time() > deadline:
                        holder = info.get("owner", "?") if 'info' in dir() else "?"
                        raise TimeoutError(
                            f"Registro occupato da un altro utente ({holder}). Riprova tra poco.")
                    time.sleep(LOCK_POLL)
                    continue
            # free or stale -> take it
            payload = json.dumps({"owner": owner,
                                  "ts": datetime.now(timezone.utc).isoformat()}).encode()
            self.write_bytes(lk, payload)
            time.sleep(0.3)  # settle; re-check we still own it
            try:
                info = json.loads(self.read_bytes(lk).decode("utf-8"))
                if info.get("owner") == owner:
                    return True
            except Exception:
                pass

    def release_lock(self, workbook):
        with contextlib.suppress(Exception):
            self.delete(self._lock_key(workbook))


# ── Local ─────────────────────────────────────────────────────────────────────

class LocalBackend(Backend):
    name = "local"

    def __init__(self, directory):
        self.dir = directory
        os.makedirs(self.dir, exist_ok=True)

    def _p(self, key):
        return os.path.join(self.dir, key)

    def read_bytes(self, key):
        with open(self._p(key), "rb") as f:
            return f.read()

    def write_bytes(self, key, data):
        dest = self._p(key)
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        tmp = dest + ".tmp"
        with open(tmp, "wb") as f:
            f.write(data)
        os.replace(tmp, dest)

    def exists(self, key):
        return os.path.exists(self._p(key))

    def delete(self, key):
        with contextlib.suppress(FileNotFoundError):
            os.remove(self._p(key))


# ── S3-compatible ─────────────────────────────────────────────────────────────

class S3Backend(Backend):
    name = "s3"

    def __init__(self, bucket, prefix="", endpoint=None, region=None):
        import boto3
        self.bucket = bucket
        self.prefix = prefix.strip("/")
        kwargs = {}
        if endpoint:
            kwargs["endpoint_url"] = endpoint
        if region:
            kwargs["region_name"] = region
        self.s3 = boto3.client("s3", **kwargs)

    def _k(self, key):
        return f"{self.prefix}/{key}" if self.prefix else key

    def read_bytes(self, key):
        obj = self.s3.get_object(Bucket=self.bucket, Key=self._k(key))
        return obj["Body"].read()

    def write_bytes(self, key, data):
        self.s3.put_object(Bucket=self.bucket, Key=self._k(key), Body=data)

    def exists(self, key):
        from botocore.exceptions import ClientError
        try:
            self.s3.head_object(Bucket=self.bucket, Key=self._k(key))
            return True
        except ClientError:
            return False

    def delete(self, key):
        self.s3.delete_object(Bucket=self.bucket, Key=self._k(key))


# ── Dropbox ───────────────────────────────────────────────────────────────────

class DropboxBackend(Backend):
    name = "dropbox"

    def __init__(self, directory="/bcw"):
        import dropbox
        self.dir = "/" + directory.strip("/")
        token = _cfg("BCW_DROPBOX_TOKEN")
        if token:
            self.dbx = dropbox.Dropbox(token)
        else:
            self.dbx = dropbox.Dropbox(
                app_key=_cfg("BCW_DROPBOX_APP_KEY"),
                app_secret=_cfg("BCW_DROPBOX_APP_SECRET"),
                oauth2_refresh_token=_cfg("BCW_DROPBOX_REFRESH_TOKEN"),
            )

    def _p(self, key):
        return f"{self.dir}/{key}"

    def read_bytes(self, key):
        _md, res = self.dbx.files_download(self._p(key))
        return res.content

    def write_bytes(self, key, data):
        import dropbox
        self.dbx.files_upload(data, self._p(key),
                              mode=dropbox.files.WriteMode.overwrite)

    def exists(self, key):
        import dropbox
        try:
            self.dbx.files_get_metadata(self._p(key))
            return True
        except dropbox.exceptions.ApiError:
            return False

    def delete(self, key):
        import dropbox
        with contextlib.suppress(dropbox.exceptions.ApiError):
            self.dbx.files_delete_v2(self._p(key))


# ── factory + session ─────────────────────────────────────────────────────────

def get_backend():
    kind = (_cfg("BCW_STORAGE", "local") or "local").lower()
    if kind == "local":
        return LocalBackend(_cfg("BCW_LOCAL_DIR", os.path.join(os.getcwd(), "data")))
    if kind == "s3":
        return S3Backend(bucket=_cfg("BCW_S3_BUCKET"),
                         prefix=_cfg("BCW_S3_PREFIX", ""),
                         endpoint=_cfg("BCW_S3_ENDPOINT"),
                         region=_cfg("AWS_REGION"))
    if kind == "dropbox":
        return DropboxBackend(_cfg("BCW_DROPBOX_DIR", "/bcw"))
    raise ValueError(f"BCW_STORAGE sconosciuto: {kind}")


def workbook_name():
    return _cfg("BCW_WORKBOOK_NAME", DEFAULT_WORKBOOK)


def download_workbook(backend, dest_path=None):
    name = workbook_name()
    data = backend.read_bytes(name)
    if dest_path is None:
        fd, dest_path = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    with open(dest_path, "wb") as f:
        f.write(data)
    return dest_path


def upload_workbook(backend, local_path):
    with open(local_path, "rb") as f:
        backend.write_bytes(workbook_name(), f.read())


@contextlib.contextmanager
def workbook_session(backend, owner="app", read_only=False):
    """Lock -> pull workbook to a temp file -> yield path -> push back -> unlock.

    On read_only=True the workbook is pulled but not pushed back and no lock
    is taken (safe for dashboards/exports)."""
    name = workbook_name()
    if not read_only:
        backend.acquire_lock(name, owner)
    local = None
    try:
        local = download_workbook(backend)
        yield local
        if not read_only:
            upload_workbook(backend, local)
    finally:
        if not read_only:
            backend.release_lock(name)
        if local and os.path.exists(local):
            with contextlib.suppress(Exception):
                os.remove(local)
