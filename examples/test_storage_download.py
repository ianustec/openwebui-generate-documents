"""Local smoke test: remote-storage downloads run off the event loop.

With S3/GCS/Azure storage, Open WebUI's ``Storage.get_file`` is a sync call
that downloads the object to a local path before it returns. Called on the
event loop it stalls every other request on that worker for the length of
the download.

    python examples/test_storage_download.py
"""
from __future__ import annotations

import asyncio
import importlib.util
import os
import tempfile
import threading
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

_spec = importlib.util.spec_from_file_location(
    "generate_documents", BASE / "generate_documents.py"
)
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)

FILE_ID = "0b9f2c1e-aaaa-bbbb-cccc-1234567890ab"


class _Row:
    id = FILE_ID
    path = "s3://bucket/uploads/logo.png"
    filename = "logo.png"
    meta: dict = {}


class _Files:
    @staticmethod
    def get_file_by_id(file_id):
        return _Row() if file_id == FILE_ID else None


def _with_storage(storage, fn):
    saved = (mod._OwuiFiles, mod._OwuiStorage, mod._HAS_OWUI_STORAGE)
    mod._OwuiFiles, mod._OwuiStorage, mod._HAS_OWUI_STORAGE = _Files, storage, True
    try:
        return fn()
    finally:
        mod._OwuiFiles, mod._OwuiStorage, mod._HAS_OWUI_STORAGE = saved


def _download(tmp: str, uri: str) -> str:
    out = os.path.join(tmp, "downloaded.png")
    with open(out, "wb") as fh:
        fh.write(b"PNGDATA")
    return out


def test_sync_get_file_runs_in_a_worker_thread():
    seen: dict = {}
    with tempfile.TemporaryDirectory() as tmp:
        class _Storage:
            @staticmethod
            def get_file(uri):
                seen["thread"] = threading.get_ident()
                seen["uri"] = uri
                return _download(tmp, uri)

        async def run():
            seen["loop_thread"] = threading.get_ident()
            return await mod._read_owui_file_bytes(FILE_ID)

        data = _with_storage(_Storage, lambda: asyncio.run(run()))
        assert data == b"PNGDATA"
        assert seen["uri"] == _Row.path
        assert seen["thread"] != seen["loop_thread"]
        # the downloaded copy of a remote object is cleaned up
        assert not os.path.exists(os.path.join(tmp, "downloaded.png"))
    print("OK Storage.get_file runs in a worker thread")


def test_async_get_file_still_works():
    with tempfile.TemporaryDirectory() as tmp:
        class _Storage:
            @staticmethod
            async def get_file(uri):
                return _download(tmp, uri)

        data = _with_storage(
            _Storage, lambda: asyncio.run(mod._read_owui_file_bytes(FILE_ID))
        )
        assert data == b"PNGDATA"
    print("OK an async Storage.get_file is awaited")


def main():
    test_sync_get_file_runs_in_a_worker_thread()
    test_async_get_file_still_works()
    print("ALL SMOKE TESTS PASSED")


if __name__ == "__main__":
    main()
