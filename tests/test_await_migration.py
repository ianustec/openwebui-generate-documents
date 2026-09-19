#!/usr/bin/env python3
"""Guard: no OpenWebUI coroutine may be called without ``await``.

WHY
---
Open WebUI turned nearly every ``open_webui.models.*`` method -- and every
``open_webui.routers.*`` handler -- into an ``async def``, with no
compatibility shim. Calling one without ``await`` does not raise: it returns a
coroutine object, which is **truthy**, so the guard right after it

    item = upload_file_handler(...)                      # coroutine object
    fid = getattr(item, "id", None) if item else None    # truthy, fid is None

silently does nothing and the tool still reports success to the user, having
saved nothing. Because the tool must also keep working on older Open WebUI
releases (``required_open_webui_version: 0.4.0``), where these are plain
``def``, the call sites go through ``_maybe_await`` rather than a bare
``await``.

WHAT THIS GUARD DOES
--------------------
1. ``test_no_unawaited_open_webui_call`` -- AST-scans the tool for any call
   reached through an ``open_webui`` import. Fail-closed: it must be awaited
   (directly, or via ``await _maybe_await(...)``) unless it is named in
   ``SYNC_ALLOWLIST`` with its reason.
2. ``test_every_local_coroutine_is_awaited`` -- catches a *half* migration: a
   local ``async def`` whose callers were not updated.
3. ``test_scanner_catches_the_pre_fix_shape`` -- the negative control. The
   scanner must FAIL on a snippet reproducing the original defect; without it
   a scanner that matches nothing would pass silently.

Run:  python -m pytest tests/test_await_migration.py -q
"""
from __future__ import annotations

import ast
from pathlib import Path

TARGET = Path(__file__).resolve().parent.parent / "generate_documents.py"

# Calls that are genuinely synchronous upstream and must NOT be awaited.
# Verified against open-webui v0.11.0:
#   storage/provider.py -- ``get_file`` is a plain ``def`` on every provider
#     (Local, S3, GCS, Azure, and the dispatching Storage).
#   routers/images.py:CreateImageForm and models/users.py:UserModel are
#     pydantic classes; calling one constructs it.
SYNC_ALLOWLIST = {
    "_OwuiStorage.get_file",
    "_OwuiCreateImageForm",
    "_OwuiUserModel",
}


def _awaited_call_ids(tree: ast.AST) -> set[int]:
    """Ids of Call nodes that are awaited, directly or via ``_maybe_await``."""
    ids: set[int] = set()
    for n in ast.walk(tree):
        if not (isinstance(n, ast.Await) and isinstance(n.value, ast.Call)):
            continue
        call = n.value
        ids.add(id(call))
        func = call.func
        name = getattr(func, "id", None) or getattr(func, "attr", None)
        if name == "_maybe_await":
            # ``await _maybe_await(X())`` awaits X() too.
            ids.update(id(a) for a in call.args if isinstance(a, ast.Call))
    return ids


def unawaited_open_webui_calls(source: str) -> list[tuple[int, str]]:
    """Every call reached through an ``open_webui`` import that is not awaited."""
    tree = ast.parse(source)
    imported: set[str] = set()
    for n in ast.walk(tree):
        if isinstance(n, ast.ImportFrom) and n.module and n.module.startswith("open_webui"):
            imported.update(a.asname or a.name for a in n.names)

    awaited = _awaited_call_ids(tree)
    hits = []
    for n in ast.walk(tree):
        if not isinstance(n, ast.Call) or id(n) in awaited:
            continue
        if isinstance(n.func, ast.Attribute):
            base = ast.unparse(n.func.value)
            if base.split(".")[-1] not in imported:
                continue
            label = f"{base}.{n.func.attr}"
        elif isinstance(n.func, ast.Name) and n.func.id in imported:
            label = n.func.id
        else:
            continue
        if label not in SYNC_ALLOWLIST:
            hits.append((n.lineno, label))
    return sorted(hits)


def unawaited_local_coroutine_calls(source: str) -> list[tuple[int, str]]:
    """Calls to a local ``async def`` that are not awaited."""
    tree = ast.parse(source)
    coros = {n.name for n in ast.walk(tree) if isinstance(n, ast.AsyncFunctionDef)}
    awaited = _awaited_call_ids(tree)
    hits = []
    for n in ast.walk(tree):
        if not isinstance(n, ast.Call) or id(n) in awaited:
            continue
        name = (
            n.func.attr if isinstance(n.func, ast.Attribute)
            else n.func.id if isinstance(n.func, ast.Name)
            else None
        )
        if name in coros:
            hits.append((n.lineno, name))
    return sorted(hits)


SOURCE = TARGET.read_text()


def test_no_unawaited_open_webui_call():
    hits = unawaited_open_webui_calls(SOURCE)
    assert not hits, (
        "un-awaited OpenWebUI coroutine call(s) -- these fail SILENTLY "
        f"(a coroutine object is truthy): {hits}"
    )


def test_every_local_coroutine_is_awaited():
    hits = unawaited_local_coroutine_calls(SOURCE)
    assert not hits, f"local `async def` called without await: {hits}"


def test_scanner_catches_the_pre_fix_shape():
    """Negative control: the scanner must FAIL on the original defect."""
    bad = (
        "from open_webui.models.users import Users\n"
        "from open_webui.routers.files import upload_file_handler\n"
        "def _save(user_dict, request, upload):\n"
        "    user_model = Users.get_user_by_id(user_dict['id'])\n"
        "    return upload_file_handler(request=request, file=upload, user=user_model)\n"
    )
    found = {label for _, label in unawaited_open_webui_calls(bad)}
    assert found == {"Users.get_user_by_id", "upload_file_handler"}, found


def test_scanner_accepts_the_maybe_await_shape():
    """Positive control: the fixed shape must pass."""
    good = (
        "from open_webui.models.users import Users\n"
        "async def _maybe_await(v):\n"
        "    return v\n"
        "async def _save(user_dict):\n"
        "    return await _maybe_await(Users.get_user_by_id(user_dict['id']))\n"
    )
    assert unawaited_open_webui_calls(good) == []
    assert unawaited_local_coroutine_calls(good) == []


def test_scanner_negative_control_on_the_local_coroutine_check():
    bad = "async def helper():\n    return 1\nasync def caller():\n    return helper()\n"
    assert unawaited_local_coroutine_calls(bad) == [(4, "helper")]
