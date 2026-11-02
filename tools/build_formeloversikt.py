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


def section_name(label):
    s = nb(label)
    for a, b in (("1. INPUT – ", ""), ("2. INPUT – ", ""), ("3. INPUT – ", ""), ("4. INPUT – ", ""), ("5. INPUT – ", ""), ("6. INPUT – ", ""),
                 ("3. UTREKNING – ", ""), ("4. UTREKNING – ", ""), ("5. UTREKNING – ", ""), ("1. UTREKNING – ", ""), ("4b. UTREKNING – ", ""),
                 ("4. BEREGNING – ", ""), ("3. BEREGNING – ", ""), ("5. BEREGNING – ", ""), ("1. BEREGNING – ", ""), ("4b. BEREGNING – ", ""),
                 ("6. INPUT/UTREKNING – ", ""), ("6. INPUT/BEREGNING – ", ""), ("5. OUTPUT – ", ""), ("6. OUTPUT – ", ""), ("7. OUTPUT – ", "")):
        if s.startswith(a):
            return s[len(a):]
    return s


def write_sheet(wb, sh):
    ws = wb.create_sheet(SHEET_NAME[sh["key"]][:31])
    ws["A1"] = nb(sh.get("title", sh["sheet"]))
    ws["A1"].font = TITLE
    ws["A2"] = nb(sh.get("desc", "")) if sh.get("desc") else ""
    ws["A2"].font = SUB
    if sh.get("per_label"):
        ws["A3"] = "Per enhet = %s. Nøkkeltall er uthevet." % nb(sh["per_label"])
        ws["A3"].font = MUTED
    heads = ["Post", "Type", "Kontointervall / kilde", "Formel", "Enhet", "I rapporten", "Formel (tekniske navn)"]
    r0 = 5
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=r0, column=i, value=h)
        c.font, c.fill, c.alignment = HEAD, FILL_HEAD, Alignment(vertical="center")
    row = r0 + 1
    for r in sh["rows"]:
        if r["t"] == "head":
            c = ws.cell(row=row, column=1, value=section_name(r["label"]))
            for col in range(1, len(heads) + 1):
                ws.cell(row=row, column=col).fill = FILL_SEC
            c.font = SEC
            row += 1
            continue
        typ, src, form, tech = row_info(sh, r)
        unit = {"kr/eining": "kr", "tekst": ""}.get(r.get("unit", ""), nb(r.get("unit", "")))
        vals = [nb(r["label"]), typ, src, form, unit, "Nøkkeltall" if r.get("kpi") else "", tech]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = KPI if (r.get("kpi") and i == 1) else (MUTED if i == 7 else TXT)
            c.alignment, c.border = WRAP, LINE
            if r.get("kpi"):
                c.fill = FILL_KPI
        tf = TYPE_FILL.get(typ)
        if tf:
            ws.cell(row=row, column=2).fill = PatternFill("solid", fgColor=tf)
        row += 1
    for col, w in zip("ABCDEFG", (46, 17, 34, 70, 8, 12, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B6"
    ws.auto_filter.ref = "A%d:G%d" % (r0, row - 1)
    ws.sheet_view.showGridLines = False


def les_meg(wb):
    ws = wb.active
    ws.title = "Les meg"
    lines = [
        ("Formeloversikt – nøkkeltallsanalyse jordbruk", TITLE),
        ("Alle poster, kontointervaller og formler i analysen, én fane per produksjon. Lages automatisk fra samme oppsett som nettsiden og analyse-arket.", SUB),
        ("", TXT),
        ("Slik leser du fanene", SEC),
        ("• Post: navnet slik det står i analysen og rapporten. Nøkkeltall er uthevet og merket i kolonnen «I rapporten».", TXT),
        ("• Type: Konto = summen av Tripletex-kontoene i kontointervallet. Produksjonsdata = antall, mengder og areal. Fôrnorm = standardverdi. Formel = regnes ut fra andre poster.", TXT),
        ("• Kontointervall: f.eks. «4000-4019» betyr alle kontoer fra 4000 til 4019. «unntatt» betyr at kontoene holdes utenfor. Inntekter regnes som positive tall.", TXT),
        ("• Formel: postene står i [hakeparentes]. ÷ er delt på, × er ganger. En post fra en annen fane har fanenavnet i parentes, f.eks. [Maskinkostnader – Melk] (Faste kostnader).", TXT),
        ("• (1 hvis Melk er med) er 1 når produksjonen er med i analysen, ellers 0. Brukes i totalene.", TXT),
        ("• Per enhet: postene delt på antall dyr/dekar (se teksten øverst i hver fane).", TXT),
        ("", TXT),
        ("Rekkefølge i analysen", SEC),
        ("1. Kontoene i saldobalansen kobles til produksjoner (fanen Kontokobling), fellespotter fordeles med fordelingsnøkler.", TXT),
        ("2. Hver produksjon: inntekter − variable kostnader = DB uten tilskudd; + tilskudd = DB med tilskudd.", TXT),
        ("3. Faste kostnader trekkes fra samlet DB → driftsoverskudd jordbruk; avstemmes mot skattemessig resultat (Totaløkonomi).", TXT),
        ("", TXT),
        ("Formlene kan endres i nettsiden under Innstillinger → Konteringsoppsett. Denne oversikten viser standardoppsettet.", MUTED),
    ]
    for i, (t, f) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=t)
        c.font = f
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 130
    ws.sheet_view.showGridLines = False


def kontokobling(wb):
    ws = wb.create_sheet("Kontokobling")
    ws["A1"] = "Kontokobling – hvor kontoene havner"
    ws["A1"].font = TITLE
    ws["A2"] = "Standard kobling fra Tripletex-konto til produksjon. Første produksjon i lista som er med i analysen får kontoen. Kan overstyres i Kontoplan i nettsiden."
    ws["A2"].font = SUB
    heads = ["Kontointervall", "Går til", "Merknad"]
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font, c.fill = HEAD, FILL_HEAD
    row = 5
    for acc, targets in spec.ACCOUNT_PROD:
        names = [PROD_NAME.get(t, POOL.get(t, t)) for t in targets]
        note = "Fordeles med fordelingsnøkkel" if any(t.startswith("POOL") for t in targets) else ("Første som er med" if len(targets) > 1 else "")
        for i, v in enumerate([acc_txt(acc), " → ".join(names), note], 1):
            c = ws.cell(row=row, column=i, value=v)
            c.font, c.alignment, c.border = TXT, WRAP, LINE
        row += 1
    row += 1
    ws.cell(row=row, column=1, value="Kontoer med produktkode (suffiks), f.eks. 1451.20.511").font = SEC
    row += 1
    for suf, targets in spec.SUFFIX_PROD:
        for i, v in enumerate([".%s" % suf, " → ".join(PROD_NAME.get(t, POOL.get(t, t)) for t in targets), ""], 1):
            c = ws.cell(row=row, column=i, value=v)
            c.font, c.alignment, c.border = TXT, WRAP, LINE
        row += 1
    for col, w in zip("ABC", (48, 60, 32)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False


def main():
    wb = Workbook()
    les_meg(wb)
    for sh in spec.PRODS:
        write_sheet(wb, sh)
    for sh in spec.COMMON:
        write_sheet(wb, sh)
    kontokobling(wb)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    pub = os.path.join(ROOT, "public", os.path.basename(OUT))
    shutil.copyfile(OUT, pub)
    print("Lagra", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
