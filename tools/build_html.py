# -*- coding: utf-8 -*-
"""Byggjer app/analyse.html frå app/src/template.html + spec.py + kontoplanen.

    python tools/build_html.py
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
import spec  # noqa: E402
from formula import render, tokenize  # noqa: E402
from bokmal import nb  # noqa: E402

ROOT = os.path.join(os.path.dirname(__file__), "..")
TEMPLATE = os.path.join(ROOT, "app", "src", "template.html")
OUT = os.path.join(ROOT, "app", "analyse.html")
WEB = os.path.join(ROOT, "public", "index.html")

SHEETS = spec.COMMON + spec.PRODS
labels = {}
for sh in SHEETS:
    for r in sh["rows"]:
        if r["t"] in ("input", "calc"):
            labels["%s.%s" % (sh["key"], r["id"])] = r["label"]


def full(name, key):
    return name if "." in name else "%s.%s" % (key, name)


def compile_sheet(sh):
    key = sh["key"]
    rows = []
    for r in sh["rows"]:
        o = {k: r.get(k) for k in ("t", "id", "label", "unit", "fmt", "kind", "acc", "sign", "src", "note", "kpi")}
        o = {k: v for k, v in o.items() if v not in (None, "", False)}
        if r.get("label"):
            o["label_nb"] = nb(r["label"])
        if r["t"] == "input":
            o["ex"] = r["val"]
        if r["t"] == "calc":
            o["expr"] = r["expr"]
            toks = tokenize(r["expr"])
            o["js"] = render(toks, lambda n: 'V("%s")' % full(n, key), lambda k: 'ACTIVE("%s")' % k, "js")
            o["txt"] = render(toks, lambda n: labels[full(n, key)] + ("" if "." not in n else " (%s)" % dict(
                (s["key"], s["sheet"]) for s in SHEETS)[n.split(".")[0]]),
                lambda k: "1 dersom %s er med" % spec.PNAME[k], "txt")
        rows.append(o)
    out = {k: sh.get(k) for k in ("key", "sheet", "title", "dg", "group", "unit_label", "per_label", "desc")}
    for k in ("sheet", "title", "unit_label", "per_label", "desc"):
        if sh.get(k):
            out[k + "_nb"] = nb(sh[k])
    out["rows"] = rows
    return out


def kontoplan():
    K = json.load(open(os.path.join(os.path.dirname(__file__), "kontoplan_full.json"), encoding="utf-8"))
    out, groups, avds = [], {}, {}
    for k in K:
        if k["inaktiv"]:
            continue
        g = (k.get("grp") or "")[:2]
        groups[g] = (k.get("grp") or "")[3:]
        avd = (k.get("avd") or "").strip()
        code = avd.split(" ")[0] if avd else ""
        if code:
            avds[code] = avd[len(code):].replace("(Inaktiv)", "").strip()
        out.append([k["nr"], k["navn"], k["type"][:1] if k["type"] else "", k["m1"], k["m2"], g, code])
    return {"accounts": out, "groups": groups, "avds": avds}


SPECJSON = {
    "prods": [compile_sheet(p) for p in spec.PRODS],
    "common": [compile_sheet(c) for c in spec.COMMON],
    "prodKeys": spec.PROD_KEYS, "grovforKeys": spec.GROVFOR_KEYS, "husdyrKeys": spec.HUSDYR_KEYS,
    "suffixProd": spec.SUFFIX_PROD, "accountProd": spec.ACCOUNT_PROD,
}

html = open(TEMPLATE, encoding="utf-8").read()
html = html.replace("/*__SPEC__*/null", json.dumps(SPECJSON, ensure_ascii=False, separators=(",", ":")))
html = html.replace("/*__KONTOPLAN__*/null", json.dumps(kontoplan(), ensure_ascii=False, separators=(",", ":")))
open(OUT, "w", encoding="utf-8").write(html)
print("Lagra", os.path.abspath(OUT), "(%d kB)" % (len(html.encode()) // 1024))
# Same fil som startside for nettstaden (Vercel publiserer mappa public/)
os.makedirs(os.path.dirname(WEB), exist_ok=True)
open(WEB, "w", encoding="utf-8").write(html)
print("Lagra", os.path.abspath(WEB))
