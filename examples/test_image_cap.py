"""Local smoke tests: image downloads stop at ``_MAX_IMAGE_BYTES``.

``![alt](https://...)`` lets the model (or a pasted document) point the tool
at any URL, and a brand logo can be read from the user's terminal. Without a
cap the whole response was buffered in memory, however large. The URL test
uses an in-process httpx transport; the terminal test fakes the Open WebUI
helpers and serves the file from a local aiohttp server.

    python examples/test_image_cap.py
"""
from __future__ import annotations

import asyncio
import importlib.util
import sys
import types
from pathlib import Path

import httpx

BASE = Path(__file__).resolve().parent.parent

_spec = importlib.util.spec_from_file_location(
    "generate_documents", BASE / "generate_documents.py"
)
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)

CAP = mod._MAX_IMAGE_BYTES


async def _chunks(total: int, size: int = 1024 * 1024):
    sent = 0
    while sent < total:
        n = min(size, total - sent)
        sent += n
        yield b"\0" * n


def _handler(request: httpx.Request) -> httpx.Response:
    path = request.url.path
    if path == "/small.png":
        return httpx.Response(200, headers={"content-type": "image/png"}, content=b"PNG" * 10)
    if path == "/declared-huge.png":
        return httpx.Response(
            200,
            headers={"content-type": "image/png", "content-length": str(CAP + 1)},
            content=b"x",
        )
    if path == "/streamed-huge.png":
        # no content-length: the size is only known while reading
        return httpx.Response(
            200, headers={"content-type": "image/png"}, content=_chunks(CAP + 1),
        )
    if path == "/streamed-at-cap.png":
        return httpx.Response(
            200, headers={"content-type": "image/png"}, content=_chunks(CAP),
        )
    if path == "/page.html":
        return httpx.Response(200, headers={"content-type": "text/html"}, content=b"<html>")
    return httpx.Response(404)


def _fetch(url: str):
    real = httpx.AsyncClient

    class _Client(real):
        def __init__(self, *args, **kwargs):
            kwargs["transport"] = httpx.MockTransport(_handler)
            super().__init__(*args, **kwargs)

    mod.httpx.AsyncClient = _Client
    try:
        return asyncio.run(mod._fetch_image_url(url))
    finally:
        mod.httpx.AsyncClient = real


def test_url_image_cap():
    assert _fetch("https://img.test/small.png") == b"PNG" * 10
    assert len(_fetch("https://img.test/streamed-at-cap.png")) == CAP
    assert _fetch("https://img.test/declared-huge.png") is None
    assert _fetch("https://img.test/streamed-huge.png") is None
    assert _fetch("https://img.test/page.html") is None
    assert _fetch("https://img.test/missing.png") is None
    print("OK URL images are capped at _MAX_IMAGE_BYTES")


def _install_terminal_fakes(base_url: str) -> None:
    def _mod(name, **attrs):
        m = types.ModuleType(name)
        m.__dict__.update(attrs)
        sys.modules[name] = m

    class _Config:
        @staticmethod
        async def get(key, default=None):
            return [{"id": "term", "enabled": True}]

    class _Groups:
        @staticmethod
        async def get_groups_by_member_id(user_id, include_inherited=True):
            return []

    class _User:
        def __init__(self, **kw):
            self.id = kw.get("id")

    async def _has_connection_access(user, conn, group_ids):
        return True

    async def _build_tool_server_headers(conn, request, user, metadata=None):
        return {}, {}

    for pkg in ("open_webui", "open_webui.models", "open_webui.utils"):
        sys.modules.setdefault(pkg, types.ModuleType(pkg))
    _mod("open_webui.models.config", Config=_Config)
    _mod("open_webui.models.groups", Groups=_Groups)
    _mod("open_webui.models.users", UserModel=_User)
    _mod("open_webui.utils.terminals",
         terminal_context_id=lambda conn, metadata, kind: None,
         get_terminal_server_url=lambda conn: base_url,
         TERMINAL_CONTEXT_HEADER="X-Terminal-Context")
    _mod("open_webui.utils.access_control", has_connection_access=_has_connection_access)
    _mod("open_webui.utils.tools", build_tool_server_headers=_build_tool_server_headers)


async def _terminal_case(size: int, chunked: bool):
    from aiohttp import web

    async def read(request):
        resp = web.StreamResponse(headers={"Content-Type": "image/png"})
        if not chunked:
            resp.content_length = size
        await resp.prepare(request)
        left = size
        try:
            while left:
                n = min(256 * 1024, left)
                await resp.write(b"\0" * n)
                left -= n
            await resp.write_eof()
        except ConnectionError:
            pass  # the tool hung up at the cap
        return resp

    app = web.Application()
    app.router.add_get("/files/read", read)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "127.0.0.1", 0)
    await site.start()
    port = site._server.sockets[0].getsockname()[1]
    _install_terminal_fakes(f"http://127.0.0.1:{port}")
    try:
        return await mod._read_terminal_file(
            "logo.png", request=object(), user_dict={"id": "u1"},
            metadata={"terminal_id": "term"},
        )
    finally:
        await runner.cleanup()


def test_terminal_image_cap():
    try:
        import aiohttp  # noqa: F401
    except ImportError:
        print("SKIP terminal image cap (aiohttp not installed)")
        return
    assert len(asyncio.run(_terminal_case(CAP, chunked=True))) == CAP
    assert len(asyncio.run(_terminal_case(3 * 1024 * 1024, chunked=False))) == 3 * 1024 * 1024
    assert asyncio.run(_terminal_case(CAP + 1, chunked=False)) is None
    assert asyncio.run(_terminal_case(CAP + 1, chunked=True)) is None
    print("OK terminal images are capped at _MAX_IMAGE_BYTES")


def main():
    test_url_image_cap()
    test_terminal_image_cap()
    print("ALL SMOKE TESTS PASSED")


if __name__ == "__main__":
    main()
