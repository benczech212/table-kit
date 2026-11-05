"""Assembles hero-ledger.html: template + rules data + engine + the Excel workbook embedded for download."""
import base64, os
HERE = os.path.dirname(os.path.abspath(__file__))
r = lambda n, m="r": open(os.path.join(HERE, n), m, **({} if "b" in m else {"encoding": "utf-8"})).read()
html = r("page.template.html")
html = html.replace("/*DATA*/", "window.HERO_DATA=" + r("data.json").replace("</", "<\\/") + ";window.HERO_XLSX=\"" + base64.b64encode(r("Hero-Ledger.xlsx", "rb")).decode() + "\";")
html = html.replace("/*ENGINE*/", r("engine.js"))
open(os.path.join(HERE, "hero-ledger.html"), "w", encoding="utf-8").write(html)
print(len(html) // 1024, "KB")
