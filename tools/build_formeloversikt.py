# -*- coding: utf-8 -*-
"""Byggjer excel/Formeloversikt_jordbruk.xlsx frå spec.py: éi lettlest liste per produksjon
med alle postar, kontointervall og formlar (på bokmål), i same rekkjefølgje som analysen.

    python tools/build_formeloversikt.py
"""
import os
import re
import shutil
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, os.path.dirname(__file__))
import spec  # noqa: E402
from bokmal import nb  # noqa: E402
from formula import render, tokenize  # noqa: E402

ROOT = os.path.join(os.path.dirname(__file__), "..")
OUT = os.path.join(ROOT, "excel", "Formeloversikt_jordbruk.xlsx")

F = "Arial"
TITLE = Font(name=F, size=15, bold=True, color="1F3B2D")
SUB = Font(name=F, size=10, color="404040")
HEAD = Font(name=F, size=10, bold=True, color="FFFFFF")
SEC = Font(name=F, size=10, bold=True, color="1F3B2D")
TXT = Font(name=F, size=10)
KPI = Font(name=F, size=10, bold=True)
MUTED = Font(name=F, size=9, color="595959", italic=True)
FILL_HEAD = PatternFill("solid", fgColor="2E5E45")
FILL_SEC = PatternFill("solid", fgColor="E4EDE6")
FILL_KPI = PatternFill("solid", fgColor="F3F6F2")
TYPE_FILL = {"Konto": "EAF2FB", "Produksjonsdata": "FFF4E0", "Fôrnorm": "FFF4E0", "Fordelingsnøkkel": "F1ECF8", "Formel": None}
LINE = Border(bottom=Side(style="thin", color="DDDDDD"))
WRAP = Alignment(wrap_text=True, vertical="top")

SHEETS = spec.COMMON + spec.PRODS
ROWS = {"%s.%s" % (sh["key"], r["id"]): r for sh in SHEETS for r in sh["rows"] if r["t"] != "head"}
SHEET_NAME = {sh["key"]: nb(sh["sheet"]) for sh in SHEETS}
POOL = {"POOL_AREAL": "Fellespott areal (fordeles etter areal)", "POOL_GF": "Fellespott grovfôr (fordeles etter vinterfôr)",
        "POOL_HUSDYR": "Fellespott husdyr (fordeles etter kraftfôr og veterinær)", "POOL_TILSK": "Fellespott tilskudd (fordeles etter tilskudd)",
        "FASTE": "Faste kostnader", "INNMEK": "Innandørsmekanisering (herav-post i faste kostnader)", "TOTAL": "Totaløkonomi (ikke fordelt)"}
PROD_NAME = {p["key"]: nb(p["sheet"]) for p in spec.PRODS}


def lbl(full):
    return nb(ROWS[full]["label"])


def fullid(n, key):
    return n if "." in n else "%s.%s" % (key, n)


def per_prod(expr):
    parts = expr.replace(" ", "").split("+")
    ms = [re.match(r"^(\w+)\.(\w+)\*ACTIVE\((\w+)\)$", x) for x in parts]
    if len(parts) > 1 and all(m and m.group(1) == m.group(3) for m in ms):
        seen = []
        for m in ms:
            if m.group(2) not in [x[1] for x in seen]:
                seen.append((m.group(1), m.group(2)))
        return seen
    return None


def formel(expr, key):
    pp = per_prod(expr)
    if pp:
        return "Sum av %s for alle produksjoner som er med i analysen" % " / ".join("[%s]" % lbl("%s.%s" % x) for x in pp)
    def ref(n):
        f = fullid(n, key)
        if f not in ROWS:
            return n
        s = "[%s]" % lbl(f)
        return s + (" (%s)" % SHEET_NAME[f.split(".")[0]] if f.split(".")[0] != key else "")
    out = render(tokenize(expr), ref, lambda k: "(1 hvis %s er med)" % PROD_NAME.get(k, k), "txt")
    # deling mellom ledd (operator) blir ÷, men ikkje skråstrek inne i namn
    parts = re.split(r"(\[[^\]]*\])", out)
    return "".join(x if x.startswith("[") else x.replace(" / ", " ÷ ") for x in parts)


def acc_txt(acc):
    if not acc:
        return ""
    parts = [x.strip() for x in acc.split(",") if x.strip()]
    inc = [p for p in parts if not p.startswith("!")]
    exc = [p[1:] for p in parts if p.startswith("!")]
    return ", ".join(inc) + (" (unntatt %s)" % ", ".join(exc) if exc else "")


def row_info(sh, r):
    """(type, kontointervall / kilde, formel med namn, formel med id)"""
    key, full = sh["key"], "%s.%s" % (sh["key"], r["id"])
    if r["t"] == "calc":
        return "Formel", "", formel(r["expr"], key), r["expr"]
    kind = r.get("kind")
    if kind in ("tilskot", "inn", "kost", "gf", "fast") and r.get("acc"):
        sign = "inntekt (kredit = +)" if r.get("sign") == -1 else "kostnad (debet = +)"
        return "Konto", acc_txt(r["acc"]), "Sum av kontoene, " + sign, ""
    if kind == "pct":
        return "Fordelingsnøkkel", nb(r.get("src", "")), "Andel fra fordelingsnøkkelen (kan overstyres)", ""
    if kind == "norm":
        return "Fôrnorm", nb(r.get("src", "")), "Standardverdi %s (kan endres)" % str(r.get("val", "")).replace(".", ","), ""
    if r["id"] in ("andre_inn", "andre_kost") and not r.get("acc"):
        return "Konto", "Øvrige %sxxx-kontoer koblet til produksjonen" % ("3" if r["id"] == "andre_inn" else "4"), "Rest som ikke passer i andre poster", ""
    if kind == "fast":
        return "Konto", nb(r.get("src", "")), "Hentes automatisk fra saldobalansen", ""
    if kind == "info" and r.get("acc"):
        return "Konto", acc_txt(r["acc"]), "Herav-post (informasjon), hentes fra kontoene", ""
    src = nb(r.get("src", "")) or "Fylles inn"
    auto = re.search(r"[Mm]engde|1451|1452|1453|1454|1455|1456|konto", src)
    return "Produksjonsdata", src, "Hentes fra mengde i regnskapet (kan overstyres)" if auto else "Fylles inn", ""


def expand(expr, key, depth=0):
    """Formel med kontointervall sette inn: utrekna postar i same ark blir opna i parentes."""
    def ref(n):
        f = fullid(n, key)
        r = ROWS.get(f)
        if not r:
            return n
        sk = f.split(".")[0]
        if r["t"] == "calc" and sk in (key, "faste") and depth < 6 and not per_prod(r["expr"]):
            inner = expand(r["expr"], sk, depth + 1)
            return inner if re.fullmatch(r"[^ ]+", inner) else "(" + inner + ")"
        if r["t"] == "input":
            kind = r.get("kind")
            if kind in ("tilskot", "inn", "kost", "gf", "fast", "info") and r.get("acc"):
                a = acc_txt(r["acc"]).replace(", ", "+")
                return "(konto %s)" % a if "+" in a else "konto " + a
            if kind == "innmek":
                return "(konto %s)" % acc_txt(spec.INNMEK_ACC).replace(", ", "+")
            if kind == "pct":
                return lbl(f) + " (fordelingsnøkkel)"
            if r["id"] in ("andre_inn", "andre_kost"):
                return "øvrige %sxxx" % ("3" if r["id"] == "andre_inn" else "4")
        return lbl(f) + ("" if sk == key else " (%s)" % SHEET_NAME[sk])
    pp = per_prod(expr)
    if pp:
        return "sum av " + " / ".join(lbl("%s.%s" % x) for x in pp) + " for produksjonene som er med"
    out = render(tokenize(expr), ref, lambda k: "(1 hvis %s er med)" % PROD_NAME.get(k, k), "txt")
    return dedup_parens(out.replace(" / ", " ÷ "))


def dedup_parens(s):
    """Fjernar doble parentesar «((x))» og parentes rundt heile uttrykket."""
    while True:
        pairs, st = {}, []
        for i, ch in enumerate(s):
            if ch == "(":
                st.append(i)
            elif ch == ")" and st:
                pairs[st.pop()] = i
        drop = None
        for o, c in pairs.items():
            if s[o + 1:o + 2] == "(" and pairs.get(o + 1) == c - 1:
                drop = (o, c)
                break
        if not drop:
            break
        s = s[:drop[0]] + s[drop[0] + 1:drop[1]] + s[drop[1] + 1:]
    return s


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Nøkkeltall"
    ws["A1"] = "Nøkkeltall og formler – jordbruksanalyse"
    ws["A1"].font = TITLE
    ws["A2"] = "Første rad er formelen med navn. Raden under viser formelen med kontointervallene satt inn. Enhet er tom når tallet er i kroner."
    ws["A2"].font = MUTED
    for i, h in enumerate(["Nøkkeltall", "Formel", "Enhet"], 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font, c.fill = HEAD, FILL_HEAD
    row = 5
    for sh in spec.PRODS + [spec.COMMON[1], spec.COMMON[0], spec.COMMON[2]]:
        key = sh["key"]
        kp = [r for r in sh["rows"] if r["t"] == "calc" and r.get("kpi") and r.get("fmt") != "txt"]
        if not kp:
            continue
        c = ws.cell(row=row, column=1, value=nb(sh.get("title") if key in ("grovfor", "faste", "total") else sh["sheet"]))
        c.font = SEC
        for col in range(1, 4):
            ws.cell(row=row, column=col).fill = FILL_SEC
        row += 1
        for r in kp:
            unit = r.get("unit", "")
            unit = "" if unit in ("kr", "kr/eining", "tekst") else ("%" if unit == "%" else nb(unit))
            ws.cell(row=row, column=1, value=nb(r["label"])).font = KPI
            ws.cell(row=row, column=2, value=formel(r["expr"], key)).font = TXT
            ws.cell(row=row, column=3, value=unit).font = TXT
            for col in range(1, 4):
                ws.cell(row=row, column=col).alignment = WRAP
            row += 1
            ex = expand(r["expr"], key)
            c = ws.cell(row=row, column=2, value=ex)
            c.font, c.alignment = MUTED, WRAP
            for col in range(1, 4):
                ws.cell(row=row, column=col).border = LINE
            row += 1
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 120
    ws.column_dimensions["C"].width = 9
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    pub = os.path.join(ROOT, "public", os.path.basename(OUT))
    shutil.copyfile(OUT, pub)
    print("Lagra", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
