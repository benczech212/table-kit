"""Builds docs/hero-ledger/index.html for GitHub Pages from hero-ledger.html. The artifact version stays untouched."""
import os, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
subprocess.check_call([sys.executable, os.path.join(HERE, "build_page.py")])
page = open(os.path.join(HERE, "hero-ledger.html"), encoding="utf-8").read()
nav = '<style>.suite{font:500 .95rem "Alegreya Sans",system-ui,sans-serif;padding:8px 16px;display:flex;flex-wrap:wrap;gap:4px 16px}.suite a{color:var(--blue)}.suite b{color:var(--accent)}</style><nav class="suite" aria-label="Table Kit"><a href="../">Table Kit</a><b>Hero Ledger</b><a href="../map-builder/">Map Builder</a><a href="../story-manager/">Story Manager</a></nav>'
html = '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"></head><body style="margin:0">' + nav + page + "</body></html>"
os.makedirs(os.path.join(HERE, "docs", "hero-ledger"), exist_ok=True)
open(os.path.join(HERE, "docs", "hero-ledger", "index.html"), "w", encoding="utf-8").write(html)
print("docs/hero-ledger/index.html", len(html) // 1024, "KB")
