"""Conversion-fidelity check: compare the words in the Pandoc conversion of the docx
with the words on the rendered page, ignoring order (sections were moved).

Usage: python _source/check_fidelity.py   (after quarto render)
"""

import html
import re
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

raw = (ROOT / "_source" / "raw.md").read_text(encoding="utf-8")
raw = raw[raw.index("# Abstract"):]
# Removed on purpose: the Word TOC and the 2012 About the Author (replaced by the PDF text).
raw = raw[: raw.index("# About the Author")] + raw[raw.index("# Introduction"):]
raw = re.sub(r"!\[[^\]]*\]\([^)]*\)\{[^}]*\}", "", raw)  # images; alt text moves to attributes
raw = re.sub(r"\]\(([^)]*)\)", "]", raw)  # link targets
raw = re.sub(r"\{[^}]*\}", "", raw)

page = (ROOT / "_site" / "index.html").read_text(encoding="utf-8")
main = page[page.index("<main"): page.index("</main>")]
main = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", main, flags=re.S)
out = html.unescape(re.sub(r"<[^>]+>", " ", main))


def tokens(s):
    s = s.replace("’", "'").replace("‘", "'").replace("\\", "")
    words = (w.lower().strip("'.:,;") for w in re.findall(r"[A-Za-z0-9][A-Za-z0-9'.-]*", s))
    return Counter(w for w in words if w)


src, dst = tokens(raw), tokens(out)
missing, added = src - dst, dst - src
print(f"source tokens {sum(src.values())}, page tokens {sum(dst.values())}")
print(f"\nMISSING from page ({sum(missing.values())}):\n{dict(missing.most_common())}")
print(f"\nADDED on page ({sum(added.values())}):\n{dict(added.most_common())}")
