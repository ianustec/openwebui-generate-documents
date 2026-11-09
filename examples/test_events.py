"""Local smoke tests for the events the tool sends (no Open WebUI required).

- A .docx saved through the Files API is attached to the assistant message
  with a ``files`` event, which Open WebUI stores with the message and
  renders as a file card in both function-calling modes. The card opens the
  file by id, so the ``/cache/files`` fallback (no file row) sends none.
- The live ``message`` event with the link is sent only on Legacy function
  calling; in Native mode (the default) the streamed completion snapshots
  replace the message content.

    python examples/test_events.py
"""
from __future__ import annotations

import asyncio
import importlib.util
import tempfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

_spec = importlib.util.spec_from_file_location(
    "generate_documents", BASE / "generate_documents.py"
)
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
SOURCE = "---\ntemplate: blank\ntitle: Events\n---\n\n# Hello\n\nBody.\n"


def _capture():
    events: list[dict] = []

    async def emitter(event):
        events.append(event)

    return events, emitter


def test_file_card_for_files_api_save():
    events, emitter = _capture()
    url = "/api/v1/files/0b9f2c1e-aaaa-bbbb-cccc-1234567890ab/content"
    asyncio.run(mod.Tools._emit_file(emitter, "Q3 Review.docx", url, size=42850))
    assert len(events) == 1 and events[0]["type"] == "files"
    (f,) = events[0]["data"]["files"]
    assert f == {
        "type": "file",
        "name": "Q3 Review.docx",
        "url": url,
        "size": 42850,
        "content_type": DOCX,
        "id": "0b9f2c1e-aaaa-bbbb-cccc-1234567890ab",
    }
    print("OK file card for a Files API save")


def test_no_file_card_for_cache_fallback():
    events, emitter = _capture()
    asyncio.run(mod.Tools._emit_file(emitter, "x.docx", "/cache/files/x_ab12cd.docx", size=1))
    assert events == []
    print("OK no file card for the /cache/files fallback")


def test_emitter_missing_or_failing():
    url = "/api/v1/files/0b9f2c1e-aaaa-bbbb-cccc-1234567890ab/content"
    asyncio.run(mod.Tools._emit_file(None, "x.docx", url, size=1))

    async def boom(_event):
        raise RuntimeError("socket gone")

    asyncio.run(mod.Tools._emit_file(boom, "x.docx", url, size=1))
    print("OK no emitter / failing emitter is ignored")


def test_legacy_detection():
    legacy = mod._is_legacy_function_calling
    assert legacy({"params": {"function_calling": "legacy"}})
    for meta in (None, {}, {"params": {}}, {"params": {"function_calling": "native"}},
                 {"params": None}, "legacy"):
        assert not legacy(meta), meta
    print("OK Legacy detection")


def _run(metadata):
    events, emitter = _capture()
    tool = mod.Tools()
    tool.valves.emit_status = False
    with tempfile.TemporaryDirectory() as tmp:
        tool.valves.docx_export_dir = tmp
        reply = asyncio.run(tool.generate_document(
            SOURCE, __event_emitter__=emitter, __metadata__=metadata,
        ))
    return reply, [e["type"] for e in events]


def test_generate_document_events_per_mode():
    # outside Open WebUI the document goes to the /cache/files fallback, so
    # no file card here; this checks the live link per mode
    reply, types = _run(None)
    assert "/cache/files/" in reply
    assert types == [], types
    _, types = _run({"params": {"function_calling": "legacy"}})
    assert types == ["message"], types
    print("OK live link only on Legacy function calling")


def main():
    test_file_card_for_files_api_save()
    test_no_file_card_for_cache_fallback()
    test_emitter_missing_or_failing()
    test_legacy_detection()
    test_generate_document_events_per_mode()
    print("ALL SMOKE TESTS PASSED")


if __name__ == "__main__":
    main()
