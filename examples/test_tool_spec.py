"""Local smoke test: the function-calling spec Open WebUI builds from the tool.

Open WebUI does not read ``self.tools``. It walks the public methods of
``Tools`` and builds each spec from the signature and the reST docstring:
the text before the first ``:param``/``:return`` line is the description and
``:param name:`` (plus its continuation lines, until the next line starting
with ``:``) describes the argument. The two parsers below mirror
``open_webui.utils.tools.parse_description`` / ``parse_docstring``.

    python examples/test_tool_spec.py
"""
from __future__ import annotations

import importlib.util
import re
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

_spec = importlib.util.spec_from_file_location(
    "generate_documents", BASE / "generate_documents.py"
)
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)


def parse_description(docstring: str) -> str:
    lines = [line.strip() for line in docstring.strip().split("\n")]
    out = []
    for line in lines:
        if re.match(r":param", line) or re.match(r":return", line):
            break
        out.append(line)
    return "\n".join(out)


def parse_docstring(docstring: str) -> dict:
    pattern = re.compile(r":param (\w+):\s*(.*)")
    params: dict = {}
    current = None
    for line in docstring.splitlines():
        line = line.strip()
        m = pattern.match(line)
        if m:
            name, text = m.groups()
            current = None if name.startswith("__") else name
            if current:
                params[current] = text
            continue
        if line.startswith(":"):
            current = None
            continue
        if current and line:
            params[current] = "\n".join(filter(None, [params[current], line]))
    return params


def test_content_parameter_is_described():
    doc = mod.Tools.generate_document.__doc__
    params = parse_docstring(doc)
    assert set(params) == {"content"}, params
    content = params["content"]
    # every section reaches the parameter (none is cut off by a line that
    # starts with ":")
    for section in ("FRONTMATTER", "LETTERHEAD", "CLIENT BRAND", "BODY",
                    "FENCED DIVS", "JSON blocks"):
        assert section in content, section
    assert "\\newpage" in content
    print("OK content parameter carries the format reference")


def test_description_stops_before_params():
    desc = parse_description(mod.Tools.generate_document.__doc__)
    assert desc.startswith("Generate a native Word (.docx) document")
    assert ":param" not in desc and "Args:" not in desc
    print("OK description is the summary only")


def test_no_dead_descriptor():
    tool = mod.Tools()
    assert not hasattr(tool, "tools")
    assert not hasattr(mod.Tools, "_tool_descriptor")
    print("OK no unused tool descriptor")


def main():
    test_content_parameter_is_described()
    test_description_stops_before_params()
    test_no_dead_descriptor()
    print("ALL SMOKE TESTS PASSED")


if __name__ == "__main__":
    main()
