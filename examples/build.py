"""Render examples/report.md into a .docx locally (no Open WebUI required).

    pip install python-docx pillow markdown-it-py mdit-py-plugins PyYAML
    python examples/build.py
"""
import asyncio
import importlib.util
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

_spec = importlib.util.spec_from_file_location(
    "generate_documents", BASE / "generate_documents.py"
)
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)

source = (BASE / "examples" / "report.md").read_text()

tool = mod.Tools()
raw = tool._parse_content(source)
raw.setdefault("template", "report")
spec = mod._resolve_template(raw)

doc = asyncio.run(tool._build_document(spec))
out = BASE / "examples" / "demo_report.docx"
doc.save(out)
print(f"OK: {out} ({out.stat().st_size} bytes)")

# Client brand: preset examples/brands/ianustec.json, logo next to it
tool.valves.brands_dir = str(BASE / "examples" / "brands")
raw = tool._parse_content((BASE / "examples" / "brand.md").read_text())
spec = asyncio.run(tool._prepare_spec(raw))
doc = asyncio.run(tool._build_document(spec))
out = BASE / "examples" / "demo_brand.docx"
doc.save(out)
print(f"OK: {out} ({out.stat().st_size} bytes)")
