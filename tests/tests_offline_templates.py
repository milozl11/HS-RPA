"""Offline guard: every DOM id the inline scripts touch must exist in the template."""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ROOT / "server" / "templates"

SCRIPT_RE = re.compile(
    r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", re.DOTALL | re.IGNORECASE
)
STYLE_RE = re.compile(r"<style[^>]*>(.*?)</style>", re.DOTALL | re.IGNORECASE)
ID_ATTR_RE = re.compile(r"""\bid\s*=\s*["']([A-Za-z0-9_-]+)["']""")
LOOKUP_RE = re.compile(
    r"""(?:\$\(|document\.getElementById\()\s*["']([A-Za-z0-9_-]+)["']\s*\)"""
)
# ids created at runtime by template literals, not present in the static markup
DYNAMIC_ID_RE = re.compile(r"""\bid\s*=\s*\\?["']\$\{""")


def check(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    scripts = "\n".join(SCRIPT_RE.findall(text))
    markup = SCRIPT_RE.sub("", STYLE_RE.sub("", text))
    declared = set(ID_ATTR_RE.findall(markup)) | set(ID_ATTR_RE.findall(scripts))
    used = set(LOOKUP_RE.findall(scripts))
    return sorted(used - declared)


def main() -> int:
    failures = 0
    for path in sorted(TEMPLATES.glob("*.html")):
        missing = check(path)
        status = "PASS" if not missing else "FAIL"
        if missing:
            failures += 1
        print(f"{status} {path.name}" + (f" missing={missing}" if missing else ""))
    print("RESULT=", "PASS" if not failures else "FAIL")
    return 0 if not failures else 1


if __name__ == "__main__":
    sys.exit(main())
