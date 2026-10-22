# -*- coding: utf-8 -*-
"""
Felles spesifikasjon for nøkkeltalsanalysen i jordbruket.

Same spesifikasjon blir brukt av
  * build_excel.py  -> excel/Analyse_jordbruk_produksjonar.xlsx
  * build_html.py   -> app/analyse.html

Kvar "ark" (produksjon eller fellesark) er ei liste med rader:
  input : verdi som skal fyllast inn (produksjonsdata eller kroner frå rekneskapen)
  calc  : utrekning, skriven i eit enkelt formelspråk
  head  : seksjonsoverskrift

Formelspråk:
  id              -> rad i same ark
  ark.id          -> rad i anna ark (t.d. faste.maskin_sau, grovfor.da_vk)
  DIV(a,b)        -> a/b, men 0 dersom b = 0
  IF(v,a,b), MAX, MIN, ABS, ROUND(x,n)
  ACTIVE(ark)     -> 1 dersom produksjonen er med i totalen (Oversikt), elles 0
Forteikn: alle inntekter og kostnader er positive kronebeløp.
"""

# ---------------------------------------------------------------- hjelparar

def H(title):
    return {"t": "head", "label": title}


def I(id, label, unit, val, src="", note="", fmt=None, kind=None, acc=None, sign=None):
    """Inputrad. kind: prod | tilskot | inn | kost | gf | info | pct | fast
    acc  = Tripletex-kontoar (4-sifra basekonto) som blir summert til posten i HTML-appen.
    sign = +1: verdi = saldo (kostnad), -1: verdi = -saldo (inntekt / auke i lager)."""
    if fmt is None:
        fmt = "kr" if unit == "kr" else ("pct" if unit == "%" else "n")
    return {"t": "input", "id": id, "label": label, "unit": unit, "val": val,
            "src": src, "note": note, "fmt": fmt, "kind": kind or "prod",
            "acc": acc, "sign": sign}


def C(id, label, expr, unit="kr", fmt=None, note="", kpi=False):
    if fmt is None:
        fmt = {"kr": "kr", "kr/eining": "kr2", "%": "pct", "tekst": "txt"}.get(unit, "n2")
    return {"t": "calc", "id": id, "label": label, "expr": expr, "unit": unit,
            "fmt": fmt, "note": note, "kpi": kpi}


TILSKOT_ACC = "3400-3439, 3450-3499"
TILSKOT_SRC = "34xx (ikkje 344x avløysartilskot)"

# Felles kostnadsposter for grovfôrbaserte husdyrproduksjonar --------------
GF_POSTS = [
    ("gf_safro", "Såfrø / grasfrø", "4001", "4001"),
    ("gf_plantevern", "Plantevernmiddel", "4110, 4191", "4110, 4191"),
    ("gf_plast", "Rundballeplast mm.", "4140, 4194", "4140, 4194"),
    ("gf_gjodsel", "Handelsgjødsel", "4050, 4095", "4050, 4095"),
    ("gf_kalk", "Kalk", "4057", "4057"),
    ("gf_konserv", "Konserveringsmiddel", "4080, 4098", "4080, 4098"),
    ("gf_lager", "Endring anna varelager (variable kostn.)", "4090", "4090"),
    ("gf_endring", "Endring eigenprodusert grovfôr (+ = nedgang i lager)", "4297", "4297"),
]
GF_COST_POSTS = [
    ("kjop_grovfor", "Kjøp grovfôr", "4030", "4030"),
    ("beiteleige", "Beiteleige / dyr på beite", "4031", "4031"),
    ("lager_innkjopt", "Endring innkjøpt grovfôrlager", "4093", "4093"),
    ("kraftfor", "Kraftfôr mm.", "4020-4029, 4040-4049, 4092, 4094", "4020-4029, 4040-4049, 4092, 4094"),
    ("vet", "Veterinær, medisin, semin mm.", "4120, 4192, 4520-4529", "4120, 4192, 4520-4529"),
    ("innkjop_dyr", "Innkjøp dyr", "4600-4689", "4600-4689"),
    ("andre_kost", "Andre variable kostnader", "4130-4139, 4193, 4530-4549, 4560-4569 + andre 4xxx", "4130-4139, 4193, 4530-4549, 4560-4569"),
]
KF_COST_POSTS = [
    ("kraftfor", "Kraftfôr mm.", "4020-4029, 4040-4049, 4092, 4094", "4020-4029, 4040-4049, 4092, 4094"),
    ("vet", "Veterinær, medisin, inseminering", "4120, 4192, 4520-4529", "4120, 4192, 4520-4529"),
    ("innkjop_dyr", "Innkjøp dyr", "4600-4689", "4600-4689"),
    ("andre_kost", "Andre variable kostnader", "4130-4139, 4193, 4530-4549, 4560-4569 + andre 4xxx", "4130-4139, 4193, 4530-4549, 4560-4569"),
]

INN_POSTS = {
    "mjolk": ("Mjølk", "3070-3073, 3163/3173 (.482)", "3070-3073, 3163, 3173"),
    "geitmjolk": ("Mjølk", "3075, 3079, 3163/3173 (.482)", "3075, 3079, 3163, 3173"),
    "slakt": ("Slakt", "3040-3049, 3162, 3172, 3078", "3040-3049, 3162, 3172, 3078"),
    "livdyr": ("Livdyrsal", "3050-3059", "3050-3059"),
    "ull": ("Ull", "3074, 3163/3173 (.484)", "3074"),
    "egg": ("Egg", "3076, 3097, 3163/3173 (.486)", "3076, 3097"),
    "buskap": ("Buskap, beholdningsendring (+ = auke)", "4690-4698 (forteikn snudd)", "4690-4698"),
    "andre_inn": ("Andre inntekter", "andre 3xxx kopla til produksjonen", ""),
}


def inn(id, val, label=None):
    lab, src, acc = INN_POSTS[id if id in INN_POSTS else id]
    real_id = "mjolk" if id == "geitmjolk" else id
    return I(real_id, label or lab, "kr", val, src=src, kind="inn", acc=acc, sign=-1)


def husdyr_cost_rows(vals, posts):
    return [I(p[0], p[1], "kr", vals.get(p[0], 0), src=p[2], kind="gf" if p[0].startswith("gf_") else "kost",
              acc=p[3], sign=1) for p in posts]


def per_unit(rows, n="n"):
    """Lagar per-eining-rader: (label, expr) -> calc med DIV(expr, n)."""
    out = []
    for i, (lab, ex) in enumerate(rows):
        if ex is None:
            out.append(H(lab))
            continue
        out.append(C("pe_%d" % i, lab, "DIV(%s,%s)" % (ex, n), "kr/eining"))
    return out


# ---------------------------------------------------------------- grovfôr-husdyr

def grovfor_husdyr(key, sheet, title, dg, unit_label, per_label, prod_rows, inn_ids, inn_vals,
                   info_rows, cost_vals, kpis, extra_calc=(), kueiningar=None, desc=""):
    rows = [H("1. INPUT – Produksjonsdata (frå produksjonstabell / teljingar)")]
    rows += prod_rows
    rows.append(H("2. INPUT – Rekneskapstal (kr) frå saldobalansen etter korrigeringar"))
    rows.append(I("tilskot", "Produksjonstilskot fordelt på produksjonen", "kr", inn_vals.get("tilskot", 0),
                  src=TILSKOT_SRC, kind="tilskot", acc=TILSKOT_ACC, sign=-1))
    for iid in inn_ids:
        rows.append(inn(iid, inn_vals.get("mjolk" if iid == "geitmjolk" else iid, 0)))
    rows += info_rows
    rows.append(H("Variable kostnader – eigen grovfôrproduksjon"))
    rows += husdyr_cost_rows(cost_vals, GF_POSTS)
    rows.append(H("Variable kostnader – anna"))
    rows += husdyr_cost_rows(cost_vals, GF_COST_POSTS)
    rows.append(H("3. INPUT – Fordelte faste kostnader (lenka frå arket «Faste kostnader»)"))
    rows.append(C("maskin", "Maskinkostnader fordelt", "faste.maskin_%s" % key))
    rows.append(C("innmek", "Innandørsmekanisering fordelt", "faste.innmek_%s" % key))

    inn_sum = "+".join(["mjolk" if i == "geitmjolk" else i for i in inn_ids])
    rows.append(H("4. UTREKNING – Totalt for produksjonen"))
    rows += [
        C("inn_total", "Produksjonsinntekter i alt", inn_sum),
        C("gf_sum", "Eigen grovfôr – sum variable kostnader",
          "gf_safro+gf_plantevern+gf_plast+gf_gjodsel+gf_kalk+gf_konserv+gf_lager+gf_endring"),
        C("vk_gf", "Variabel grovfôrkostnad (utan plast og lagerendring eigenprodusert)",
          "gf_safro+gf_plantevern+gf_gjodsel+gf_kalk+gf_konserv+gf_lager",
          note="Brukt i «Variabel grovfôrkostnad pr dekar» – same utval som eksempelanalysen"),
        C("sum_eigen_gf", "Sum eigen grovfôrproduksjon (utan lagerendring)", "gf_sum-gf_endring",
          note="Brukt til å fordele grovfôrarealet mellom produksjonane"),
        C("gf_total", "Grovfôr – kostnad i alt", "gf_sum+kjop_grovfor+lager_innkjopt"),
        C("kost_total", "Variable kostnader i alt", "gf_total+beiteleige+kraftfor+vet+innkjop_dyr+andre_kost"),
        C("db_u", "Dekningsbidrag (DB) utan tilskot", "inn_total-kost_total", kpi=True),
        C("db_m", "Dekningsbidrag (DB) med tilskot", "db_u+tilskot", kpi=True),
        C("db_etter", "DB m/tilskot etter maskin- og innandørskostnad", "db_m-maskin-innmek"),
    ]
    if kueiningar:
        rows.append(C("kueiningar", "«Kueiningar»", kueiningar, "stk", "n2",
                      note="Vekting av dyretal (snitt 1.1. og 31.12.) – sjå formel"))
    rows += list(extra_calc)

    # per eining
    pu = [("Produksjonsinntekter %s" % per_label, None)]
    for iid in inn_ids:
        rid = "mjolk" if iid == "geitmjolk" else iid
        pu.append((INN_POSTS[iid][0], rid))
    pu += [("Produksjonsinntekter i alt", "inn_total"),
           ("Variable kostnader %s" % per_label, None),
           ("Eigen grovfôrproduksjon", "sum_eigen_gf"),
           ("Kjøpt grovfôr", "kjop_grovfor"),
           ("Varelager endring grovfôr", "gf_endring+lager_innkjopt"),
           ("Beiteleige", "beiteleige"),
           ("Kraftfôr mm.", "kraftfor"),
           ("Veterinær, medisin, semin mm.", "vet"),
           ("Innkjøp dyr", "innkjop_dyr"),
           ("Andre variable kostnader", "andre_kost"),
           ("Sum variable kostnader", "kost_total"),
           ("Resultat %s" % per_label, None),
           ("DB utan tilskot", "db_u"),
           ("Tilskot", "tilskot"),
           ("DB med tilskot", "db_m"),
           ("Maskinkostnader", "maskin"),
           ("Innandørsmekanisering", "innmek"),
           ("DB m/tilskot etter maskin og innandørsmek.", "db_etter")]
    rows.append(H("5. OUTPUT – Per eining (%s)" % per_label))
    rows += per_unit(pu)
    rows.append(H("6. OUTPUT – Nøkkeltal"))
    rows += kpis
    return {"key": key, "sheet": sheet, "title": title, "dg": dg, "group": "grovfor",
            "unit_label": unit_label, "per_label": per_label, "desc": desc, "rows": rows}


def grovfor_kpis_common(key):
    """Nøkkeltal som alle grovfôrproduksjonar har (frå arket Grovfôr)."""
    return [
        C("k_db_da_u", "DB pr dekar grovfôr u/tilskot",
          "DIV(db_u,grovfor.da_vk*DIV(sum_eigen_gf,grovfor.sum_eigen_alle))", "kr/eining", kpi=True,
          note="Grovfôrarealet blir fordelt etter produksjonen sin del av eigen grovfôrkostnad"),
        C("k_db_da_m", "DB pr dekar grovfôr m/tilskot",
          "DIV(db_m,grovfor.da_vk*DIV(sum_eigen_gf,grovfor.sum_eigen_alle))", "kr/eining", kpi=True),
        C("k_varkost_da", "Variabel grovfôrkostnad pr dekar", "grovfor.varkost_pr_da", "kr/eining", kpi=True),
        C("k_maskin_da", "Maskinkostnader pr dekar grovfôr", "grovfor.maskin_pr_da", "kr/eining", kpi=True),
        C("k_korr_da", "Korrigerte dekar pr «kueining»", "grovfor.korr_da_pr_ku", "da", "n2", kpi=True),
    ]


# ---------------------------------------------------------------- kraftfôr-husdyr

def kraftfor_husdyr(key, sheet, title, dg, unit_label, per_label, prod_rows, inn_rows, inn_ids, cost_vals,
                    tilskot, kpis, extra_calc=(), desc="", extra_cost=()):
    rows = [H("1. INPUT – Produksjonsdata (frå produksjonstabell / teljingar)")]
    rows += prod_rows
    rows.append(H("2. INPUT – Rekneskapstal (kr) frå saldobalansen etter korrigeringar"))
    rows.append(I("tilskot", "Produksjonstilskot fordelt på produksjonen", "kr", tilskot,
                  src=TILSKOT_SRC, kind="tilskot", acc=TILSKOT_ACC, sign=-1))
    rows += inn_rows
    rows.append(H("Variable kostnader"))
    rows += husdyr_cost_rows(cost_vals, KF_COST_POSTS + list(extra_cost))
    rows.append(H("3. INPUT – Fordelte faste kostnader (lenka frå arket «Faste kostnader»)"))
    rows.append(C("maskin", "Maskinkostnader fordelt", "faste.maskin_%s" % key))
    rows.append(C("innmek", "Innandørsmekanisering fordelt", "faste.innmek_%s" % key))
    rows.append(H("4. UTREKNING – Totalt for produksjonen"))
    rows += [
        C("inn_total", "Produksjonsinntekter i alt", "+".join(inn_ids)),
        C("kost_total", "Variable kostnader i alt", "+".join(["kraftfor", "vet", "innkjop_dyr", "andre_kost"] + [c[0] for c in extra_cost])),
        C("db_u", "Dekningsbidrag (DB) utan tilskot", "inn_total-kost_total", kpi=True),
        C("db_m", "Dekningsbidrag (DB) med tilskot", "db_u+tilskot", kpi=True),
        C("db_etter", "DB m/tilskot etter maskin- og innandørskostnad", "db_m-maskin-innmek"),
    ]
    rows += list(extra_calc)
    labels = {r["id"]: r["label"] for r in inn_rows if r["t"] == "input"}
    pu = [("Produksjonsinntekter %s" % per_label, None)]
    pu += [(labels[i], i) for i in inn_ids]
    pu += [("Produksjonsinntekter i alt", "inn_total"),
           ("Variable kostnader %s" % per_label, None),
           ("Kraftfôr mm.", "kraftfor"), ("Veterinær, medisin, inseminering", "vet"),
           ("Innkjøp dyr", "innkjop_dyr"), ("Andre variable kostnader", "andre_kost")]
    pu += [(c[1], c[0]) for c in extra_cost]
    pu += [("Sum variable kostnader", "kost_total"),
           ("Resultat %s" % per_label, None),
           ("DB utan tilskot", "db_u"), ("Tilskot", "tilskot"), ("DB med tilskot", "db_m"),
           ("Maskinkostnader", "maskin"), ("Innandørsmekanisering", "innmek"),
           ("DB m/tilskot etter maskin og innandørsmek.", "db_etter")]
    rows.append(H("5. OUTPUT – Per eining (%s)" % per_label))
    rows += per_unit(pu)
    rows.append(H("6. OUTPUT – Nøkkeltal"))
    rows += kpis
    return {"key": key, "sheet": sheet, "title": title, "dg": dg, "group": "kraftfor",
            "unit_label": unit_label, "per_label": per_label, "desc": desc, "rows": rows}


# ---------------------------------------------------------------- planteproduksjon

def plante(key, sheet, title, dg, unit, sal_label, sal_src, sal_acc, vals, cost_posts, extra_inn=(),
           kpis_extra=(), prod_extra=(), per_label=None, desc=""):
    unitname = "dekar" if unit == "da" else "m²"
    per_label = per_label or ("pr dekar" if unit == "da" else "pr m²")
    rows = [H("1. INPUT – Produksjonsdata")]
    rows.append(I("n", "Tal %s" % unitname, unit, vals["n"], src="Produksjonstabell / søknad om produksjonstilskot"))
    rows.append(I("avling", "Avling totalt", "kg", vals.get("avling", 0), src="Produksjonstabell / avlingsregistrering"))
    rows.append(I("sal_kg", "Seld mengd", "kg", vals.get("sal_kg", 0), src="Mengde 1/2 på salskonto (%s)" % sal_src))
    rows += list(prod_extra)
    rows.append(H("2. INPUT – Rekneskapstal (kr) frå saldobalansen etter korrigeringar"))
    rows.append(I("tilskot", "Produksjonstilskot fordelt på produksjonen", "kr", vals.get("tilskot", 0),
                  src=TILSKOT_SRC, kind="tilskot", acc=TILSKOT_ACC, sign=-1))
    rows.append(I("sal", sal_label, "kr", vals.get("sal", 0), src=sal_src, kind="inn", acc=sal_acc, sign=-1))
    inn_ids = ["sal"]
    for (iid, lab, src, acc) in extra_inn:
        rows.append(I(iid, lab, "kr", vals.get(iid, 0), src=src, kind="inn", acc=acc, sign=-1))
        inn_ids.append(iid)
    rows.append(I("beh_endring", "Beholdningsendring ferdigvarer (+ = auke)", "kr", vals.get("beh_endring", 0),
                  src="4290-4298 (forteikn snudd)", kind="inn",
                  acc="4290-4296, 4298" if key not in ("silofor", "hoy") else "4297", sign=-1))
    rows.append(I("andre_inn", "Andre inntekter", "kr", vals.get("andre_inn", 0),
                  src="andre 3xxx kopla til produksjonen", kind="inn", acc="", sign=-1))
    inn_ids += ["beh_endring", "andre_inn"]
    rows.append(H("Variable kostnader"))
    cost_ids = []
    for (cid, lab, src, acc) in cost_posts:
        rows.append(I(cid, lab, "kr", vals.get(cid, 0), src=src, kind="kost", acc=acc, sign=1))
        cost_ids.append(cid)
    rows.append(H("3. INPUT – Fordelte faste kostnader (lenka frå arket «Faste kostnader»)"))
    rows.append(C("maskin", "Maskinkostnader fordelt", "faste.maskin_%s" % key))
    rows.append(H("4. UTREKNING – Totalt for produksjonen"))
    rows += [
        C("inn_total", "Produksjonsinntekter i alt", "+".join(inn_ids)),
        C("kost_total", "Variable kostnader i alt", "+".join(cost_ids)),
        C("db_u", "Dekningsbidrag (DB) utan tilskot", "inn_total-kost_total", kpi=True),
        C("db_m", "Dekningsbidrag (DB) med tilskot", "db_u+tilskot", kpi=True),
        C("db_etter", "DB m/tilskot etter maskinkostnader", "db_m-maskin"),
    ]
    labels = {r["id"]: r["label"] for r in rows if r["t"] == "input"}
    pu = [("Produksjonsinntekter %s" % per_label, None)]
    pu += [(labels[i], i) for i in inn_ids]
    pu += [("Produksjonsinntekter i alt", "inn_total"), ("Variable kostnader %s" % per_label, None)]
    pu += [(labels[c], c) for c in cost_ids]
    pu += [("Sum variable kostnader", "kost_total"), ("Resultat %s" % per_label, None),
           ("DB utan tilskot", "db_u"), ("Tilskot", "tilskot"), ("DB med tilskot", "db_m"),
           ("Maskinkostnader", "maskin"), ("DB m/tilskot etter maskinkostnader", "db_etter")]
    rows.append(H("5. OUTPUT – Per eining (%s)" % per_label))
    rows += per_unit(pu)
    rows.append(H("6. OUTPUT – Nøkkeltal"))
    rows += [
        C("k_db_u", "DB %s u/tilskot" % per_label, "DIV(db_u,n)", "kr/eining", kpi=True),
        C("k_db_m", "DB %s m/tilskot" % per_label, "DIV(db_m,n)", "kr/eining", kpi=True),
        C("k_pris", "Oppnådd pris pr kg", "DIV(sal,sal_kg)", "kr/eining", kpi=True),
        C("k_avling", "Avling %s (kg)" % per_label, "DIV(avling,n)", "kg", "n1", kpi=True),
        C("k_maskin", "Maskinkostnader %s" % per_label, "DIV(maskin,n)", "kr/eining", kpi=True),
        C("k_db_etter", "DB etter maskinkostnader %s" % per_label, "DIV(db_etter,n)", "kr/eining", kpi=True),
    ]
    rows += list(kpis_extra)
    return {"key": key, "sheet": sheet, "title": title, "dg": dg, "group": "plante",
            "unit_label": unitname, "per_label": per_label, "desc": desc, "rows": rows}


P_GJODSEL = ("gjodsel_kalk", "Gjødsel og kalk", "4050, 4057, 4095", "4050, 4057, 4095")
P_PLANTEVERN = ("plantevern", "Plantevernmiddel", "4110, 4111, 4191", "4110, 4111, 4191")
P_EMBALLASJE = ("emballasje", "Emballasje", "4140-4142, 4194", "4140-4142, 4194")
P_ENERGI = ("energi", "Brensel, energi (veksthus)", "6200-6249 når kopla til produksjonen", "6200-6249")
P_ANDRE = ("andre_kost", "Andre variable kostnader", "4540-4549, 4160 + andre 4xxx", "4540-4549, 4160, 4170")

# =================================================================== PRODUKSJONAR
PRODS = []

# ---- KORN
PRODS.append(plante(
    "korn", "Korn", "Kornproduksjon", "1100-1190", "da", "Sal av korn", "3000-3007", "3000-3007",
    dict(n=60, avling=24000, sal_kg=20000, tilskot=42000, sal=72000, halm=6000, beh_endring=3000,
         gjodsel_kalk=27000, plantevern=9000, sakorn=15000, andre_kost=8000),
    [P_GJODSEL, P_PLANTEVERN, ("sakorn", "Såkorn", "4000", "4000"), P_ANDRE],
    extra_inn=[("halm", "Sal av halm", "3096", "3096")],
    desc="Korn til modning (driftsgrein 1100–1190)."))

# ---- SILOFOR (sal)
PRODS.append(plante(
    "silofor", "Silofor", "Sal av silofôr / rundballar", "1522-1523", "da", "Sal av silofôr", "3095", "3095",
    dict(n=50, avling=0, sal_kg=0, tilskot=28980.38, sal=206400, beh_endring=-15000,
         gjodsel_kalk=21666.4, plantevern=1223.77, grasfro=4021, andre_kost=15758.5),
    [P_GJODSEL, P_PLANTEVERN, ("grasfro", "Grasfrø", "4001", "4001"), P_ANDRE],
    prod_extra=[I("ballar", "Tal siloballar seld", "stk", 253, src="Mengde 1 på konto 3095"),
                I("lagerauke_kg", "Auke i varelager (kg)", "kg", 0, src="Mengde på konto 4297/1447 (.472)")],
    kpis_extra=[C("k_pris_balle", "Oppnådd pris pr siloballe", "DIV(sal,ballar)", "kr/eining", kpi=True),
                C("k_ballar_da", "Tal siloballar pr dekar", "DIV(ballar+lagerauke_kg/750,n)", "stk", "n2", kpi=True,
                  note="Lagerauke i kg er rekna om til ballar med 750 kg pr balle")],
    desc="Areal der graset blir selt som silofôr/rundballar (driftsgrein 1522–1523). Eksempeltal frå vedlagt analyse 2025."))

# ---- HØY (sal)
PRODS.append(plante(
    "hoy", "Høy", "Sal av høy", "1510-1513", "da", "Sal av høy", "3094", "3094",
    dict(n=20, avling=12000, sal_kg=10000, tilskot=8000, sal=30000, beh_endring=2000,
         gjodsel_kalk=8000, plantevern=0, grasfro=1000, andre_kost=3000),
    [P_GJODSEL, P_PLANTEVERN, ("grasfro", "Grasfrø", "4001", "4001"), P_ANDRE],
    desc="Areal der graset blir selt som høy (driftsgrein 1510–1513)."))

# ---- POTET
PRODS.append(plante(
    "potet", "Potet", "Potetproduksjon", "2100-2130", "da", "Sal av potet", "3010-3011", "3010-3011",
    dict(n=30, avling=75000, sal_kg=60000, tilskot=25000, sal=312000, beh_endring=15000,
         gjodsel_kalk=21000, plantevern=24000, settepotet=45000, emballasje=30000, andre_kost=20000),
    [P_GJODSEL, P_PLANTEVERN, ("settepotet", "Settepotet", "4010, 4091", "4010, 4091"), P_EMBALLASJE, P_ANDRE],
    desc="Matpotet og settepotet (driftsgrein 2100–2130)."))

# ---- AGURK
PRODS.append(plante(
    "agurk", "Agurk", "Agurk i veksthus", "3300-3330", "m2", "Sal av agurk", "3019", "3019",
    dict(n=2000, avling=100000, sal_kg=98000, tilskot=0, sal=1274000, gjodsel_kalk=60000, plantevern=15000,
         planter=90000, emballasje=70000, energi=380000, andre_kost=120000),
    [P_GJODSEL, P_PLANTEVERN, ("planter", "Frø / planter", "4019", "4019"), P_EMBALLASJE, P_ENERGI, P_ANDRE],
    desc="Veksthusagurk (driftsgrein 3300–3330). Nøkkeltal pr m²."))

# ---- TOMAT
PRODS.append(plante(
    "tomat", "Tomat", "Tomat i veksthus", "3400-3490", "m2", "Sal av tomat", "3018", "3018",
    dict(n=2500, avling=110000, sal_kg=108000, tilskot=0, sal=1512000, gjodsel_kalk=70000, plantevern=20000,
         planter=110000, emballasje=80000, energi=520000, andre_kost=150000),
    [P_GJODSEL, P_PLANTEVERN, ("planter", "Frø / planter", "4018", "4018"), P_EMBALLASJE, P_ENERGI, P_ANDRE],
    desc="Veksthustomat (driftsgrein 3400–3490). Nøkkeltal pr m²."))

# ---- FRILANDSGRØNSAK
PRODS.append(plante(
    "fgront", "Frilandsgrønsak", "Frilandsgrønsaker", "2200-2490, 2600-3290", "da", "Sal av grønsaker",
    "3013, 3015, 3020, 3039", "3013, 3015, 3020, 3039",
    dict(n=40, avling=120000, sal_kg=110000, tilskot=60000, sal=660000, gjodsel_kalk=30000, plantevern=20000,
         planter=80000, emballasje=50000, andre_kost=60000),
    [P_GJODSEL, P_PLANTEVERN, ("planter", "Frø / planter", "4012", "4012"), P_EMBALLASJE, P_ANDRE],
    desc="Frilandsgrønsaker utanom gulrot og kålrot."))

# ---- GULROT
PRODS.append(plante(
    "gulrot", "Gulrot", "Gulrot", "2500-2590", "da", "Sal av gulrot", "3014", "3014",
    dict(n=25, avling=125000, sal_kg=115000, tilskot=35000, sal=575000, beh_endring=10000, gjodsel_kalk=17000,
         plantevern=14000, fro=30000, emballasje=45000, andre_kost=40000),
    [P_GJODSEL, P_PLANTEVERN, ("fro", "Frø", "4012", "4012"), P_EMBALLASJE, P_ANDRE],
    desc="Gulrot (driftsgrein 2500–2590)."))

# ---- KÅLROT
PRODS.append(plante(
    "kalrot", "Kålrot", "Kålrot", "2280-2282", "da", "Sal av kålrot", "3013, 3015 (kopla til kålrot)", "3013, 3015",
    dict(n=15, avling=60000, sal_kg=55000, tilskot=20000, sal=357500, gjodsel_kalk=10000, plantevern=7000,
         fro=9000, emballasje=20000, andre_kost=15000),
    [P_GJODSEL, P_PLANTEVERN, ("fro", "Frø", "4012", "4012"), P_EMBALLASJE, P_ANDRE],
    desc="Kålrot (driftsgrein 2280–2282)."))

# ---- MJØLK (storfe)
PRODS.append(grovfor_husdyr(
    "mjolk", "Mjølk", "Mjølkeproduksjon storfe", "6110-6113", "årskyr", "pr årsku",
    [I("n", "Tal årskyr", "stk", 25, src="Kukontrollen / produksjonstabell (EnhetAntall)"),
     I("liter", "Levert mjølk", "liter", 190000, src="Mengde 1 på konto 3071"),
     I("kvote", "Disponibel mjølkekvote", "liter", 200000, src="Kvoteoversikt (Landbruksdirektoratet)"),
     I("kg_kjott", "Kg slakt storfe (mjølkeproduksjonen)", "kg", 5200, src="Mengde 2 på konto 3041"),
     I("kyr_ib", "Kyr pr 1.1.", "stk", 26, src="Mengde 1 på 1451.20.511/512 (IB)"),
     I("kyr_ub", "Kyr pr 31.12.", "stk", 25, src="Mengde 1 på 1451.20.511/512 (UB)"),
     I("ung_o_ib", "Kviger/oksar over 12 md pr 1.1.", "stk", 14, src="1451.20.514/516 (IB)"),
     I("ung_o_ub", "Kviger/oksar over 12 md pr 31.12.", "stk", 15, src="1451.20.514/516 (UB)"),
     I("ung_u_ib", "Kalvar/ungdyr under 12 md pr 1.1.", "stk", 16, src="1451.20.515/517/518 (IB)"),
     I("ung_u_ub", "Kalvar/ungdyr under 12 md pr 31.12.", "stk", 17, src="1451.20.515/517/518 (UB)")],
    ["mjolk", "slakt", "livdyr", "buskap", "andre_inn"],
    dict(tilskot=520000, mjolk=1330000, slakt=210000, livdyr=45000, buskap=15000, andre_inn=8000),
    [I("distrikt_mjolk", "Herav distriktstilskot mjølk", "kr", 22800, src="Del av mjølkeinntekta (avrekning)", kind="info", acc=""),
     I("distrikt_kjott", "Herav distriktstilskot kjøt", "kr", 27300, src="3047 Slakt, distriktstilskot (del av slakteinntekta)", kind="info", acc="3047")],
    dict(gf_safro=18000, gf_plantevern=6000, gf_plast=32000, gf_gjodsel=95000, gf_kalk=12000, gf_konserv=14000,
         gf_lager=0, gf_endring=-10000, kjop_grovfor=20000, beiteleige=0, lager_innkjopt=0, kraftfor=380000,
         vet=65000, innkjop_dyr=0, andre_kost=95000),
    [C("k_db_liter", "DB pr liter mjølk u/tilskot", "DIV(db_u,liter)", "kr/eining", kpi=True),
     C("k_db_ku_u", "DB pr årsku u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_ku_m", "DB pr årsku m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True)]
    + grovfor_kpis_common("mjolk") +
    [C("k_mjolkepris", "Oppnådd mjølkepris pr liter (u/distriktstilskot)", "DIV(mjolk-distrikt_mjolk,liter)", "kr/eining", kpi=True),
     C("k_kjottpris", "Oppnådd slaktepris pr kg (u/distriktstilskot)", "DIV(slakt-distrikt_kjott,kg_kjott)", "kr/eining", kpi=True),
     C("k_maskin_liter", "Maskinkostnader pr liter", "DIV(maskin,liter)", "kr/eining", kpi=True),
     C("k_innmek_liter", "Innandørsmekanisering pr liter", "DIV(innmek,liter)", "kr/eining", kpi=True),
     C("k_innmek_ku", "Innandørsmekanisering pr årsku", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_lev_ku", "Levert mjølk pr årsku (liter)", "DIV(liter,n)", "liter", "n", kpi=True),
     C("k_pasett", "Påsettprosent", "pasett", "%", kpi=True),
     C("k_kvote", "Kvoteoppfylling", "DIV(liter,kvote)", "%", kpi=True),
     C("k_gruppe", "Analysegruppe (påsett)", 'IF(pasett<=0.8,"Påsett < 80 %",IF(pasett<1.2,"Påsett 80–120 %","Påsett > 120 %"))', "tekst", kpi=True,
       note="Tilsvarar arka «Mjølk <80%», «Mjølk 80-120%» og «Mjølk >120%» i eksempelanalysen")],
    extra_calc=[C("pasett", "Påsettprosent", "DIV((ung_o_ib+ung_o_ub+ung_u_ib+ung_u_ub)/2,kyr_ib+kyr_ub)", "%",
                  note="Snitt ungdyr / 2 delt på (kyr 1.1. + kyr 31.12.) – same oppsett som eksempelanalysen")],
    kueiningar="((kyr_ib+kyr_ub)+(ung_o_ib+ung_o_ub)/4+(ung_u_ib+ung_u_ub)/8)/2",
    desc="Mjølkeku med påsett (driftsgrein 6110–6113). Analysegruppe etter påsettprosent blir rekna ut automatisk."))

# ---- OKSAR
PRODS.append(grovfor_husdyr(
    "oksar", "Oksar", "Okseproduksjon", "6120-6123", "oksar", "pr okse",
    [I("n", "Tal produserte oksar", "stk", 12, src="Produksjonstabell (EnhetAntall)"),
     I("kg_slakt", "Kg slakt", "kg", 3960, src="Mengde 2 på konto 3041 (.516/.517/.518)"),
     I("tal_slakta", "Tal slakta oksar", "stk", 12, src="Mengde 1 på konto 3041"),
     I("tal_livs", "Tal storfe selt til liv", "stk", 0, src="Mengde 1 på konto 3051"),
     I("tal_kjopt", "Tal kjøpte kalvar", "stk", 10, src="Mengde 1 på konto 4610"),
     I("o12_ib", "Oksar over 12 md pr 1.1.", "stk", 6, src="1451.20.516 (IB)"),
     I("o12_ub", "Oksar over 12 md pr 31.12.", "stk", 7, src="1451.20.516 (UB)"),
     I("u12_ib", "Oksar/kalvar under 12 md pr 1.1.", "stk", 6, src="1451.20.517/518 (IB)"),
     I("u12_ub", "Oksar/kalvar under 12 md pr 31.12.", "stk", 6, src="1451.20.517/518 (UB)")],
    ["slakt", "livdyr", "buskap", "andre_inn"],
    dict(tilskot=95000, slakt=290000, livdyr=0, buskap=12000, andre_inn=0),
    [I("distrikt_kjott", "Herav distriktstilskot kjøt", "kr", 20790, src="3047 Slakt, distriktstilskot (del av slakteinntekta)", kind="info", acc="3047")],
    dict(gf_safro=3000, gf_plantevern=1000, gf_plast=6000, gf_gjodsel=15000, gf_kalk=2000, gf_konserv=2500,
         gf_lager=0, gf_endring=-2000, kjop_grovfor=0, beiteleige=0, lager_innkjopt=0, kraftfor=60000, vet=8000,
         innkjop_dyr=70000, andre_kost=12000),
    [C("k_db_u", "DB pr okse u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr okse m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True)]
    + grovfor_kpis_common("oksar") +
    [C("k_slaktepris", "Snitt slaktepris pr okse", "DIV(slakt,tal_slakta)", "kr/eining", kpi=True),
     C("k_kjopspris", "Snitt kjøpspris pr kalv", "DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_diff", "Slaktepris – kjøpspris pr okse", "DIV(slakt,tal_slakta)-DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_kjottpris", "Oppnådd pris pr kg slakt (u/distriktstilskot)", "DIV(slakt-distrikt_kjott,kg_slakt)", "kr/eining", kpi=True),
     C("k_vekt", "Slaktevekt pr okse (kg)", "DIV(kg_slakt,tal_slakta)", "kg", "n1", kpi=True),
     C("k_innmek", "Innandørsmekanisering pr okse", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_tap", "Tapsprosent", "MAX(0,DIV((o12_ib+u12_ib+tal_kjopt)-(o12_ub+u12_ub+tal_slakta+tal_livs),o12_ib+u12_ib+tal_kjopt))", "%", kpi=True,
       note="(Inngåande dyr + kjøpte) – (utgåande dyr + slakta + selde), delt på inngåande + kjøpte")],
    kueiningar="((u12_ib+u12_ub)/8+(o12_ib+o12_ub)/4)/2",
    desc="Oksar/kastratar til slakt (driftsgrein 6120–6123)."))

# ---- MELLOMKALV
PRODS.append(grovfor_husdyr(
    "mellomkalv", "Mellomkalv", "Mellomkalv", "6124", "kalvar", "pr kalv",
    [I("n", "Tal produserte kalvar", "stk", 20, src="Produksjonstabell (EnhetAntall)"),
     I("kg_slakt", "Kg slakt", "kg", 3400, src="Mengde 2 på konto 3041"),
     I("tal_slakta", "Tal slakta kalvar", "stk", 20, src="Mengde 1 på konto 3041"),
     I("tal_livs", "Tal selt til liv", "stk", 0, src="Mengde 1 på konto 3051"),
     I("tal_kjopt", "Tal kjøpte kalvar", "stk", 20, src="Mengde 1 på konto 4610"),
     I("o12_ib", "Ungdyr over 12 md pr 1.1.", "stk", 0, src="1451.20 (IB)"),
     I("o12_ub", "Ungdyr over 12 md pr 31.12.", "stk", 0, src="1451.20 (UB)"),
     I("u12_ib", "Kalvar under 12 md pr 1.1.", "stk", 8, src="1451.20 (IB)"),
     I("u12_ub", "Kalvar under 12 md pr 31.12.", "stk", 8, src="1451.20 (UB)")],
    ["slakt", "livdyr", "buskap", "andre_inn"],
    dict(tilskot=30000, slakt=180000, livdyr=0, buskap=0, andre_inn=0),
    [I("distrikt_kjott", "Herav distriktstilskot kjøt", "kr", 8000, src="3047 Slakt, distriktstilskot (del av slakteinntekta)", kind="info", acc="3047")],
    dict(gf_safro=500, gf_plantevern=0, gf_plast=1500, gf_gjodsel=3000, gf_kalk=500, gf_konserv=500,
         gf_lager=0, gf_endring=0, kjop_grovfor=0, beiteleige=0, lager_innkjopt=0, kraftfor=30000, vet=6000,
         innkjop_dyr=90000, andre_kost=6000),
    [C("k_db_u", "DB pr kalv u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr kalv m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True)]
    + grovfor_kpis_common("mellomkalv") +
    [C("k_slaktepris", "Snitt slaktepris pr kalv (u/distriktstilskot)", "DIV(slakt-distrikt_kjott,tal_slakta)", "kr/eining", kpi=True),
     C("k_kjopspris", "Snitt kjøpspris pr kalv", "DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_diff", "Slaktepris – kjøpspris pr kalv", "DIV(slakt,tal_slakta)-DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_kjottpris", "Oppnådd pris pr kg slakt (u/distriktstilskot)", "DIV(slakt-distrikt_kjott,kg_slakt)", "kr/eining", kpi=True),
     C("k_vekt", "Slaktevekt pr kalv (kg)", "DIV(kg_slakt,tal_slakta)", "kg", "n1", kpi=True),
     C("k_innmek", "Innandørsmekanisering pr kalv", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_tap", "Tapsprosent", "MAX(0,DIV((o12_ib+u12_ib+tal_kjopt)-(o12_ub+u12_ub+tal_slakta+tal_livs),o12_ib+u12_ib+tal_kjopt))", "%", kpi=True)],
    kueiningar="((u12_ib+u12_ub)/8+(o12_ib+o12_ub)/4)/2",
    desc="Mellomkalv (driftsgrein 6124)."))

# ---- AMMEKYR  (eksempeltal 2025 frå vedlagt analyse)
PRODS.append(grovfor_husdyr(
    "ammekyr", "Ammekyr", "Ammekuproduksjon", "6130-6132", "ammekyr", "pr ammeku",
    [I("n", "Tal ammekyr", "stk", 16, src="Produksjonstabell (EnhetAntall)"),
     I("kyr_ib", "Ammekyr pr 1.1.", "stk", 11, src="1452.20.521/522 (IB)"),
     I("kyr_ub", "Ammekyr pr 31.12.", "stk", 17, src="1452.20.521/522 (UB)"),
     I("okse_ib", "Avlsokse pr 1.1.", "stk", 0, src="1452.20.526 (IB)"),
     I("okse_ub", "Avlsokse pr 31.12.", "stk", 0, src="1452.20.526 (UB)"),
     I("o12_ib", "Ungdyr over 12 md pr 1.1.", "stk", 9, src="1451.20.514/516 (IB)"),
     I("o12_ub", "Ungdyr over 12 md pr 31.12.", "stk", 5, src="1451.20.514/516 (UB)"),
     I("u12_ib", "Kalvar/ungdyr under 12 md pr 1.1.", "stk", 10, src="1451.20.515/517 (IB)"),
     I("u12_ub", "Kalvar/ungdyr under 12 md pr 31.12.", "stk", 12, src="1451.20.515/517 (UB)"),
     I("kg_slakt_kyr", "Kg slakt kyr", "kg", 1123.3, src="Mengde 2 på 3042 (.521/.522)"),
     I("tal_slakt_kyr", "Tal slakta kyr", "stk", 3, src="Mengde 1 på 3042 (.521/.522)"),
     I("kg_slakt", "Kg slakt totalt", "kg", 1744.4, src="Mengde 2 på 3041/3042"),
     I("tal_ut", "Tal dyr slakta / selt til liv", "stk", 10, src="Mengde 1 på 3041/3042/3051/3052"),
     I("tal_kjopt", "Tal kjøpte dyr", "stk", 2, src="Mengde 1 på 4610/4620")],
    ["livdyr", "slakt", "buskap", "andre_inn"],
    dict(tilskot=287735.22, livdyr=142310.5, slakt=94047.04, buskap=99100, andre_inn=0),
    [I("distrikt_kjott", "Herav distriktstilskot kjøt", "kr", 0, src="3047 Slakt, distriktstilskot (del av slakteinntekta)", kind="info", acc="3047")],
    dict(gf_safro=5099, gf_plantevern=2736.62, gf_plast=14760, gf_gjodsel=42759.04, gf_kalk=5691.36,
         gf_konserv=4969.2, gf_lager=0, gf_endring=19021, kjop_grovfor=643.2, beiteleige=0, lager_innkjopt=0,
         kraftfor=55861.22, vet=31948.8, innkjop_dyr=21960, andre_kost=26689.6),
    [C("k_db_u", "DB pr ammeku u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr ammeku m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True)]
    + grovfor_kpis_common("ammekyr") +
    [C("k_innmek", "Innandørsmekanisering pr ammeku", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_pasett", "Påsettprosent", "pasett", "%", kpi=True),
     C("k_slaktevekt", "Snitt slaktevekt kyr (kg)", "DIV(kg_slakt_kyr,tal_slakt_kyr)", "kg", "n1", kpi=True),
     C("k_kalv", "Nyttbar kalv pr ku", "DIV((kyr_ub+okse_ub+o12_ub+u12_ub)+tal_ut-tal_kjopt-(kyr_ib+okse_ib+o12_ib+u12_ib),n)", "stk", "n2", kpi=True,
       note="(Dyr 31.12. + slakta/selde – kjøpte – dyr 1.1.) delt på tal ammekyr"),
     C("k_kjottpris", "Oppnådd pris pr kg slakt (u/distriktstilskot)", "DIV(slakt-distrikt_kjott,kg_slakt)", "kr/eining", kpi=True),
     C("k_gruppe", "Analysegruppe (påsett)", 'IF(pasett>0.5,"Påsett > 50 %","Påsett < 50 %")', "tekst", kpi=True,
       note="Tilsvarar arka «Ammekyr <50%» og «Ammekyr >50%» i eksempelanalysen")],
    extra_calc=[C("pasett", "Påsettprosent", "DIV((okse_ib+okse_ub+o12_ib+o12_ub+u12_ib+u12_ub)/2,kyr_ib+kyr_ub)", "%",
                  note="Snitt av anna storfe / 2 delt på (ammekyr 1.1. + ammekyr 31.12.)")],
    kueiningar="((kyr_ib+kyr_ub+okse_ib+okse_ub)+(o12_ib+o12_ub)/3+(u12_ib+u12_ub)/4)/2",
    desc="Ammeku med kalv (driftsgrein 6130–6132). Eksempeltal frå vedlagt analyse 2025."))

# ---- SMÅGRIS / KOMBINERT
PRODS.append(kraftfor_husdyr(
    "smagris", "Smågris", "Smågris / kombinert svineproduksjon", "6200-6242 (utan 6220)", "årspurker", "pr årspurke",
    [I("n", "Tal årspurker", "stk", 40, src="Produksjonstabell (EnhetAntall)"),
     I("kull_pr_purke", "Tal kull pr purke", "stk", 2.2, src="Ingris / produksjonstabell", fmt="n2"),
     I("m2", "Bingeareal", "m²", 400, src="Mengde 1 på 1040-1044"),
     I("gris_ib", "Tal gris pr 1.1.", "stk", 450, src="1453.20 (IB)"),
     I("gris_ub", "Tal gris pr 31.12.", "stk", 470, src="1453.20 (UB)"),
     I("tal_slaktegris", "Tal slaktegris levert", "stk", 700, src="Mengde 1 på 3043 (.535–.538)"),
     I("tal_smagris", "Tal smågris selt", "stk", 120, src="Mengde 1 på 3053 (.534)"),
     I("tal_kjopt", "Tal kjøpte smågris", "stk", 0, src="Mengde 1 på 4630"),
     I("kg_slakt", "Kg slakt", "kg", 57400, src="Mengde 2 på 3043")],
    [I("smagris_inn", "Sal av smågris", "kr", 156000, src="3053 (.534)", kind="inn", acc="3050-3059", sign=-1),
     I("slaktegris_inn", "Sal av slaktegris", "kr", 1722000, src="3043 (.535–.538), 3162, 3172", kind="inn", acc="3040-3049, 3162, 3172", sign=-1),
     I("purkeslakt", "Purkeslakt", "kr", 30000, src="3043 (.531/.532)", kind="inn", acc="", sign=-1),
     I("buskap", "Buskap, beholdningsendring (+ = auke)", "kr", 10000, src="4693 (forteikn snudd)", kind="inn", acc="4690-4698", sign=-1),
     I("andre_inn", "Andre inntekter", "kr", 0, src="andre 3xxx", kind="inn", acc="", sign=-1)],
    ["smagris_inn", "slaktegris_inn", "purkeslakt", "buskap", "andre_inn"],
    dict(kraftfor=1250000, vet=90000, innkjop_dyr=40000, andre_kost=150000), 120000,
    [C("k_db_u", "DB pr årspurke u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr årspurke m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True),
     C("k_db_kull", "DB pr kull u/tilskot", "DIV(db_u,kull)", "kr/eining", kpi=True),
     C("k_maskin_purke", "Maskinkostnader pr purke", "DIV(maskin,n)", "kr/eining", kpi=True),
     C("k_maskin_kull", "Maskinkostnader pr kull", "DIV(maskin,kull)", "kr/eining", kpi=True),
     C("k_innmek_purke", "Innandørsmekanisering pr purke", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_innmek_kull", "Innandørsmekanisering pr kull", "DIV(innmek,kull)", "kr/eining", kpi=True),
     C("k_kjottpris", "Oppnådd kjøtpris pr kg slakt", "DIV(slaktegris_inn+purkeslakt,kg_slakt)", "kr/eining", kpi=True),
     C("k_pris_smagris", "Oppnådd pris pr smågris", "DIV(smagris_inn,tal_smagris)", "kr/eining", kpi=True),
     C("k_gris_purke", "Produsert tal gris pr purke", "DIV(produsert,n)", "stk", "n1", kpi=True),
     C("k_gris_kull", "Produsert tal gris pr kull", "DIV(produsert,kull)", "stk", "n1", kpi=True),
     C("k_framforing", "Framfôringsprosent", "DIV(tal_slaktegris,tal_slaktegris+tal_smagris-tal_kjopt)", "%", kpi=True),
     C("k_db_m2", "DB pr m² bingeareal", "DIV(db_u,m2)", "kr/eining", kpi=True),
     C("k_gruppe", "Analysegruppe (oppdrett)", 'IF(DIV(tal_slaktegris,tal_slaktegris+tal_smagris-tal_kjopt)>=0.5,"> 50 % oppdrett (kombinert)","< 50 % oppdrett (smågris)")', "tekst", kpi=True,
       note="Tilsvarar arka «Gris kombinert» og «Smågris» i eksempelanalysen")],
    extra_calc=[C("kull", "Tal kull", "n*kull_pr_purke", "stk", "n1"),
                C("produsert", "Produsert tal gris", "tal_slaktegris+tal_smagris-tal_kjopt+gris_ub-gris_ib", "stk", "n")],
    desc="Smågris- og kombinert produksjon (driftsgrein 6200–6242, utan 6220)."))

# ---- PURKERING
PRODS.append(kraftfor_husdyr(
    "purkering", "Purkering", "Purkering (nav/satellitt)", "6290", "årspurker", "pr årspurke",
    [I("n", "Tal årspurker", "stk", 120, src="Produksjonstabell (EnhetAntall)"),
     I("kull_pr_purke", "Tal kull pr purke", "stk", 2.3, src="Ingris / produksjonstabell", fmt="n2"),
     I("m2", "Husdyrrom", "m²", 900, src="Mengde 1 på 1045-1049"),
     I("gris_ib", "Tal gris pr 1.1.", "stk", 900, src="1453.20 (IB)"),
     I("gris_ub", "Tal gris pr 31.12.", "stk", 920, src="1453.20 (UB)"),
     I("tal_slaktegris", "Tal slaktegris levert", "stk", 0, src="Mengde 1 på 3043"),
     I("tal_smagris", "Tal smågris selt", "stk", 3000, src="Mengde 1 på 3053"),
     I("tal_kjopt", "Tal kjøpte gris", "stk", 0, src="Mengde 1 på 4630"),
     I("kg_slakt", "Kg slakt", "kg", 6000, src="Mengde 2 på 3043")],
    [I("smagris_inn", "Sal av smågris", "kr", 3150000, src="3053", kind="inn", acc="3050-3059", sign=-1),
     I("slaktegris_inn", "Sal av slaktegris", "kr", 0, src="3043, 3162, 3172", kind="inn", acc="3040-3049, 3162, 3172", sign=-1),
     I("purkeslakt", "Purkeslakt", "kr", 120000, src="3043 (.531/.532)", kind="inn", acc="", sign=-1),
     I("buskap", "Buskap, beholdningsendring (+ = auke)", "kr", 8000, src="4693 (forteikn snudd)", kind="inn", acc="4690-4698", sign=-1),
     I("andre_inn", "Andre inntekter", "kr", 0, src="andre 3xxx", kind="inn", acc="", sign=-1)],
    ["smagris_inn", "slaktegris_inn", "purkeslakt", "buskap", "andre_inn"],
    dict(kraftfor=1700000, vet=250000, innkjop_dyr=150000, andre_kost=400000, purkeleige=300000), 200000,
    [C("k_db_u", "DB pr årspurke u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr årspurke m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True),
     C("k_db_kull", "DB pr kull u/tilskot", "DIV(db_u,kull)", "kr/eining", kpi=True),
     C("k_db_kull_m", "DB pr kull m/tilskot", "DIV(db_m,kull)", "kr/eining", kpi=True),
     C("k_maskin_kull", "Maskinkostnader pr kull", "DIV(maskin,kull)", "kr/eining", kpi=True),
     C("k_innmek_kull", "Innandørsmekanisering pr kull", "DIV(innmek,kull)", "kr/eining", kpi=True),
     C("k_maskin_purke", "Maskinkostnader pr purke", "DIV(maskin,n)", "kr/eining", kpi=True),
     C("k_innmek_purke", "Innandørsmekanisering pr purke", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_db_m2", "DB pr m² husdyrrom", "DIV(db_u,m2)", "kr/eining", kpi=True),
     C("k_kjottpris", "Oppnådd kjøtpris pr kg slakt", "DIV(slaktegris_inn+purkeslakt,kg_slakt)", "kr/eining", kpi=True),
     C("k_pris_smagris", "Oppnådd pris pr smågris", "DIV(smagris_inn,tal_smagris)", "kr/eining", kpi=True),
     C("k_gris_purke", "Produsert tal gris pr purke", "DIV(produsert,n)", "stk", "n1", kpi=True),
     C("k_gris_kull", "Produsert tal gris pr kull", "DIV(produsert,kull)", "stk", "n1", kpi=True),
     C("k_framforing", "Framfôringsprosent", "DIV(tal_slaktegris,tal_slaktegris+tal_smagris+tal_kjopt)", "%", kpi=True),
     C("k_gruppe", "Analysegruppe (oppdrett)", 'IF(DIV(tal_slaktegris,tal_slaktegris+tal_smagris+tal_kjopt)>=0.5,"> 50 % oppdrett","< 50 % oppdrett")', "tekst", kpi=True)],
    extra_calc=[C("kull", "Tal kull", "n*kull_pr_purke", "stk", "n1"),
                C("produsert", "Produsert tal gris", "tal_slaktegris+tal_smagris-tal_kjopt+gris_ub-gris_ib", "stk", "n")],
    desc="Purkering (driftsgrein 6290). Purkeleige (6480) er ein variabel kostnad i purkeringen, som i eksempelanalysen.",
    extra_cost=[("purkeleige", "Purkeleige", "6480", "6480")]))

# ---- SLAKTEGRIS
PRODS.append(kraftfor_husdyr(
    "slaktegris", "Slaktegris", "Slaktegrisproduksjon", "6220", "slaktegris", "pr slaktegris",
    [I("n", "Tal produserte slaktegris", "stk", 700, src="Produksjonstabell (EnhetAntall)"),
     I("kg_slakt", "Kg slakt", "kg", 57400, src="Mengde 2 på 3043"),
     I("tal_slakta", "Tal slakta gris", "stk", 690, src="Mengde 1 på 3043"),
     I("gris_ib", "Tal gris pr 1.1.", "stk", 180, src="1453.20 (IB)"),
     I("gris_ub", "Tal gris pr 31.12.", "stk", 190, src="1453.20 (UB)"),
     I("tal_livs", "Tal gris selt til liv", "stk", 0, src="Mengde 1 på 3053"),
     I("tal_kjopt", "Tal kjøpte smågris", "stk", 705, src="Mengde 1 på 4630"),
     I("m2", "Husdyrrom", "m²", 600, src="Mengde 1 på 1050-1054")],
    [I("slakt", "Slakt", "kr", 1693300, src="3043, 3053, 3162, 3172", kind="inn", acc="3040-3059, 3162, 3172", sign=-1),
     I("buskap", "Buskap, beholdningsendring (+ = auke)", "kr", 4000, src="4693 (forteikn snudd)", kind="inn", acc="4690-4698", sign=-1),
     I("andre_inn", "Andre inntekter", "kr", 0, src="andre 3xxx", kind="inn", acc="", sign=-1)],
    ["slakt", "buskap", "andre_inn"],
    dict(kraftfor=1000000, vet=15000, innkjop_dyr=810750, andre_kost=60000), 60000,
    [C("k_db_u", "DB pr slaktegris u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr slaktegris m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True),
     C("k_db_m2", "DB pr m² husdyrrom", "DIV(db_u,m2)", "kr/eining", kpi=True),
     C("k_kjottpris", "Oppnådd kjøtpris pr kg slakt", "DIV(slakt,kg_slakt)", "kr/eining", kpi=True),
     C("k_slaktepris", "Snitt slaktepris pr gris", "DIV(slakt,tal_slakta)", "kr/eining", kpi=True),
     C("k_kjopspris", "Snitt kjøpspris pr smågris", "DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_diff", "Slaktepris – kjøpspris pr gris", "DIV(slakt,tal_slakta)-DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_maskin", "Maskinkostnad pr slaktegris", "DIV(maskin,n)", "kr/eining", kpi=True),
     C("k_innmek", "Innandørsmekanisering pr gris", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_vekt", "Slaktevekt pr gris (kg)", "DIV(kg_slakt,tal_slakta)", "kg", "n1", kpi=True),
     C("k_tap", "Tapsprosent", "MAX(0,DIV((gris_ib+tal_kjopt)-(tal_slakta+gris_ub+tal_livs),n))", "%", kpi=True)],
    desc="Slaktegris (driftsgrein 6220)."))

# ---- SAU (eksempeltal 2025 frå vedlagt analyse)
PRODS.append(grovfor_husdyr(
    "sau", "Sau", "Sauehald", "6300-6390", "vf. sau", "pr vf. sau",
    [I("n", "Tal vinterfôra sau", "stk", 45, src="Søknad om produksjonstilskot / produksjonstabell"),
     I("kg_slakt", "Kg slakt (sau og lam)", "kg", 1586.9, src="Mengde 2 på 3044 (.541–.543)"),
     I("kg_pr_vf", "Kg slakt pr vf. sau frå produksjonstabellen (0 = rekna ut)", "kg", 38.3244, src="Produksjonstabell (AvkastningPrEnhet)", fmt="n2"),
     I("kjopte_lam", "Tal kjøpte lam (mai–des)", "stk", 0, src="Mengde 1 på 4640 (.543)"),
     I("selde_arslam", "Tal slakta/selde årslam (mai–des)", "stk", 62, src="Mengde 1 på 3044/3054 (.543)"),
     I("selde_jul_des", "Tal slakta/selde sau og lam (juli–des)", "stk", 70, src="Mengde 1 på 3044/3054"),
     I("s1_ib", "Sauer frå inngåande balanse pr 1.1. (stk)", "stk", 22, src="1454.20.541 (IB mengd)"),
     I("s2_ib", "Sauer tilført i året pr 1.1. (stk)", "stk", 14, src="1454.20.542 (IB mengd)"),
     I("s3_ib", "Sau/lam under 12 md pr 1.1. (stk)", "stk", 11, src="1454.20.543 (IB mengd)"),
     I("s1_ub", "Sauer frå inngåande balanse pr 31.12. (stk)", "stk", 29, src="1454.20.541 (UB mengd)"),
     I("s2_ub", "Sauer tilført i året pr 31.12. (stk)", "stk", 10, src="1454.20.542 (UB mengd)"),
     I("s3_ub", "Sau/lam under 12 md pr 31.12. (stk)", "stk", 12, src="1454.20.543 (UB mengd)"),
     I("s1_ib_kr", "Verdi sauer IB-gruppe pr 1.1.", "kr", 19360, src="1454.20.541 (IB)", fmt="kr"),
     I("s2_ib_kr", "Verdi sauer tilført pr 1.1.", "kr", 14070, src="1454.20.542 (IB)", fmt="kr"),
     I("s3_ib_kr", "Verdi sau/lam u. 12 md pr 1.1.", "kr", 15400, src="1454.20.543 (IB)", fmt="kr"),
     I("s1_ub_kr", "Verdi sauer IB-gruppe pr 31.12.", "kr", 23200, src="1454.20.541 (UB)", fmt="kr"),
     I("s2_ub_kr", "Verdi sauer tilført pr 31.12.", "kr", 10650, src="1454.20.542 (UB)", fmt="kr"),
     I("s3_ub_kr", "Verdi sau/lam u. 12 md pr 31.12.", "kr", 18780, src="1454.20.543 (UB)", fmt="kr")],
    ["ull", "slakt", "livdyr", "buskap", "andre_inn"],
    dict(tilskot=206447.37, ull=6386.3, slakt=103507.3, livdyr=0, buskap=3800, andre_inn=0),
    [I("distrikt_kjott", "Herav distriktstilskot kjøt", "kr", 0, src="3047 Slakt, distriktstilskot (del av slakteinntekta)", kind="info", acc="3047"),
     I("slakt_kjottpris", "Herav slakteinntekt til kjøtpris (3040–3046, 3048–3049)", "kr", 102881.42,
       src="Slakt sau/lam utan distriktstilskot og utan 3162/3172", kind="info", acc="3040-3046, 3048-3049")],
    dict(gf_safro=1243.5, gf_plantevern=857.6, gf_plast=3600, gf_gjodsel=13399.84, gf_kalk=1783.56, gf_konserv=1212,
         gf_lager=0, gf_endring=4639, kjop_grovfor=156.8, beiteleige=0, lager_innkjopt=0, kraftfor=51071.22,
         vet=14001.21, innkjop_dyr=0, andre_kost=7533.75),
    [C("k_db_u", "DB pr vf. sau u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr vf. sau m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True)]
    + grovfor_kpis_common("sau") +
    [C("k_kjottpris", "Snitt kjøtpris pr kg slakt (u/distriktstilskot)",
       "DIV(IF(slakt_kjottpris>0,slakt_kjottpris,slakt-distrikt_kjott),kg_slakt)", "kr/eining", kpi=True),
     C("k_innmek", "Innandørsmekanisering pr sau", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_slakt_vf", "Kg slakt pr vf. sau", "IF(kg_pr_vf>0,kg_pr_vf,DIV(kg_slakt,n))", "kg", "n1", kpi=True),
     C("k_gagnslam", "Gagnslam pr vf. sau", "DIV(selde_arslam+s3_ub-kjopte_lam,n)", "stk", "n2", kpi=True,
       note="(Selde/slakta årslam + lam på lager 31.12. – kjøpte lam) delt på vf. sau"),
     C("k_utmark", "Del dyr på utmarksbeite", "DIV(grovfor.utm_sau+grovfor.utm_lam,s1_ub+s2_ub+s3_ub+selde_jul_des)", "%", kpi=True),
     C("k_prisendring", "Endra DB pr vf. sau grunna prisendring på dyr", "ROUND(DIV(prisendring,n),0)", "kr/eining", kpi=True,
       note="Positivt tal = auka DB fordi verdien pr dyr har gått opp")],
    extra_calc=[
        C("pris1", "Prisendring pr dyr – sauer IB-gruppe", "IF(s1_ib*s1_ub=0,0,s1_ub_kr/s1_ub-s1_ib_kr/s1_ib)", "kr/eining"),
        C("pris2", "Prisendring pr dyr – sauer tilført", "IF(s2_ib*s2_ub=0,0,s2_ub_kr/s2_ub-s2_ib_kr/s2_ib)", "kr/eining"),
        C("pris3", "Prisendring pr dyr – sau/lam u. 12 md", "IF(s3_ib*s3_ub=0,0,s3_ub_kr/s3_ub-s3_ib_kr/s3_ib)", "kr/eining"),
        C("prisendring", "Endring i buskapsverdi grunna prisendring", "s1_ub*pris1+s2_ub*pris2+s3_ub*pris3")],
    kueiningar="(s1_ib+s2_ib+s3_ib+s1_ub+s2_ub+s3_ub)/2/6",
    desc="Sauehald (driftsgrein 6300–6390). Eksempeltal frå vedlagt analyse 2025."))

# ---- GEIT
PRODS.append(grovfor_husdyr(
    "geit", "Geit", "Mjølkegeit", "6400-6419", "geiter", "pr geit",
    [I("n", "Tal årsgeiter", "stk", 60, src="Produksjonstabell (EnhetAntall)"),
     I("liter", "Levert mjølk", "liter", 48000, src="Mengde 1 på konto 3075"),
     I("kvote", "Disponibel mjølkekvote", "liter", 50000, src="Kvoteoversikt"),
     I("g1_ib", "Geiter frå inngåande balanse pr 1.1.", "stk", 55, src="1455.20.551 (IB)"),
     I("g2_ib", "Geiter tilført i året pr 1.1.", "stk", 8, src="1455.20.552 (IB)"),
     I("g1_ub", "Geiter frå inngåande balanse pr 31.12.", "stk", 54, src="1455.20.551 (UB)"),
     I("g2_ub", "Geiter tilført i året pr 31.12.", "stk", 9, src="1455.20.552 (UB)")],
    ["geitmjolk", "slakt", "livdyr", "buskap", "andre_inn"],
    dict(tilskot=380000, mjolk=420000, slakt=15000, livdyr=10000, buskap=0, andre_inn=0),
    [I("distrikt_mjolk", "Herav distriktstilskot mjølk", "kr", 0, src="Del av mjølkeinntekta (avrekning)", kind="info", acc="")],
    dict(gf_safro=5000, gf_plantevern=1000, gf_plast=10000, gf_gjodsel=25000, gf_kalk=3000, gf_konserv=4000,
         gf_lager=0, gf_endring=0, kjop_grovfor=15000, beiteleige=0, lager_innkjopt=0, kraftfor=110000, vet=18000,
         innkjop_dyr=0, andre_kost=40000),
    [C("k_db_liter", "DB pr liter u/tilskot", "DIV(db_u,liter)", "kr/eining", kpi=True),
     C("k_db_u", "DB pr geit u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr geit m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True)]
    + grovfor_kpis_common("geit") +
    [C("k_mjolkepris", "Oppnådd mjølkepris pr liter (u/distriktstilskot)", "DIV(mjolk-distrikt_mjolk,liter)", "kr/eining", kpi=True),
     C("k_maskin_liter", "Maskinkostnader pr liter", "DIV(maskin,liter)", "kr/eining", kpi=True),
     C("k_innmek_liter", "Innandørsmekanisering pr liter", "DIV(innmek,liter)", "kr/eining", kpi=True),
     C("k_innmek_geit", "Innandørsmekanisering pr geit", "DIV(innmek,n)", "kr/eining", kpi=True),
     C("k_lev", "Levert mjølk pr geit (liter)", "DIV(liter,n)", "liter", "n", kpi=True),
     C("k_kvote", "Kvoteoppfylling", "DIV(liter,kvote)", "%", kpi=True)],
    kueiningar="(g1_ib+g2_ib+g1_ub+g2_ub)/2/7",
    desc="Mjølkegeit (driftsgrein 6400–6419)."))

# ---- EGG
PRODS.append(kraftfor_husdyr(
    "egg", "Egg", "Eggproduksjon", "6530-6532", "høner", "pr høne",
    [I("n", "Tal høner", "stk", 7500, src="Produksjonstabell (EnhetAntall)"),
     I("kg_egg", "Kg egg levert", "kg", 142500, src="Mengde 1 på konto 3076"),
     I("m2", "Husdyrrom", "m²", 1100, src="Mengde 1 på 1055-1059")],
    [I("egg", "Egg", "kr", 2992500, src="3076, 3097, 3163/3173 (.486)", kind="inn", acc="3076, 3097, 3163, 3173", sign=-1),
     I("buskap", "Buskap, beholdningsendring (+ = auke)", "kr", 0, src="4696 (forteikn snudd)", kind="inn", acc="4690-4698", sign=-1),
     I("andre_inn", "Andre inntekter (t.d. slakt av høner)", "kr", 20000, src="3046 + andre 3xxx", kind="inn", acc="", sign=-1)],
    ["egg", "buskap", "andre_inn"],
    dict(kraftfor=1350000, vet=25000, innkjop_dyr=525000, andre_kost=180000), 110000,
    [C("k_db_u", "DB pr høne u/tilskot", "DIV(db_u,n)", "kr/eining", kpi=True),
     C("k_db_m", "DB pr høne m/tilskot", "DIV(db_m,n)", "kr/eining", kpi=True),
     C("k_db_m2", "DB pr m² husdyrrom", "DIV(db_u,m2)", "kr/eining", kpi=True),
     C("k_eggpris", "Oppnådd eggpris pr kg", "DIV(egg,kg_egg)", "kr/eining", kpi=True),
     C("k_kg_hone", "Eggleveranse pr høne (kg)", "DIV(kg_egg,n)", "kg", "n2", kpi=True),
     C("k_forkost_kg", "Fôrkostnad pr kg egg", "DIV(kraftfor,kg_egg)", "kr/eining", kpi=True),
     C("k_maskin", "Maskinkostnad pr høne", "DIV(maskin,n)", "kr/eining", kpi=True),
     C("k_innmek", "Innandørsmekanisering pr høne", "DIV(innmek,n)", "kr/eining", kpi=True)],
    desc="Konsumegg (driftsgrein 6530–6532)."))

# ---- SLAKTEKYLLING
PRODS.append(kraftfor_husdyr(
    "kylling", "Slaktekylling", "Slaktekylling", "6510-6511", "kyllingar", "pr kylling",
    [I("n", "Tal produserte kyllingar", "stk", 150000, src="Produksjonstabell (EnhetAntall)"),
     I("kg_slakt", "Kg slakt", "kg", 240000, src="Mengde 2 på konto 3046 (.564/.565)"),
     I("tal_slakta", "Tal slakta kyllingar", "stk", 148500, src="Mengde 1 på konto 3046"),
     I("ib", "Tal kylling pr 1.1.", "stk", 0, src="1456.20.565 (IB)"),
     I("ub", "Tal kylling pr 31.12.", "stk", 0, src="1456.20.565 (UB)"),
     I("tal_livs", "Tal selt til liv", "stk", 0, src="Mengde 1 på 3056"),
     I("tal_kjopt", "Tal kjøpte daggamle kyllingar", "stk", 150000, src="Mengde 1 på 4660"),
     I("m2", "Husdyrrom", "m²", 2000, src="Mengde 1 på 1060-1064")],
    [I("slakt", "Slakt", "kr", 5040000, src="3046 (.564/.565)", kind="inn", acc="3040-3049", sign=-1),
     I("buskap", "Buskap, beholdningsendring (+ = auke)", "kr", 0, src="4696 (forteikn snudd)", kind="inn", acc="4690-4698", sign=-1),
     I("andre_inn", "Andre inntekter", "kr", 0, src="andre 3xxx", kind="inn", acc="", sign=-1)],
    ["slakt", "buskap", "andre_inn"],
    dict(kraftfor=2900000, vet=40000, innkjop_dyr=630000, andre_kost=600000), 40000,
    [C("k_db_u", "DB pr kylling u/tilskot", "DIV(db_u,n)", "kr/eining", "kr2", kpi=True),
     C("k_db_m", "DB pr kylling m/tilskot", "DIV(db_m,n)", "kr/eining", "kr2", kpi=True),
     C("k_db_m2", "DB pr m² husdyrrom", "DIV(db_u,m2)", "kr/eining", kpi=True),
     C("k_kgpris", "Oppnådd pris pr kg kjøt", "DIV(slakt,kg_slakt)", "kr/eining", kpi=True),
     C("k_pris", "Oppnådd pris pr slaktekylling", "DIV(slakt,tal_slakta)", "kr/eining", kpi=True),
     C("k_kjopspris", "Snitt kjøpspris pr kylling", "DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_diff", "Differanse slakt / kjøp pr kylling", "DIV(slakt,tal_slakta)-DIV(innkjop_dyr,tal_kjopt)", "kr/eining", kpi=True),
     C("k_vekt", "Slaktevekt pr kylling (kg)", "DIV(kg_slakt,tal_slakta)", "kg", "n2", kpi=True),
     C("k_tap", "Tapsprosent", "MAX(0,DIV((ib+tal_kjopt)-(tal_slakta+ub+tal_livs),ib+tal_kjopt))", "%", kpi=True),
     C("k_maskin", "Maskinkostnad pr kylling", "DIV(maskin,n)", "kr/eining", kpi=True),
     C("k_innmek", "Innandørsmekanisering pr kylling", "DIV(innmek,n)", "kr/eining", kpi=True)],
    desc="Slaktekylling (driftsgrein 6510–6511)."))

PROD_KEYS = [p["key"] for p in PRODS]
GROVFOR_KEYS = [p["key"] for p in PRODS if p["group"] == "grovfor"]
HUSDYR_KEYS = [p["key"] for p in PRODS if p["group"] in ("grovfor", "kraftfor")]
# Produksjonar som er med i totalen i eksempelet (same bruk som i vedlagt analyse 2025)
EXAMPLE_ACTIVE = {"silofor", "ammekyr", "sau"}

# =================================================================== FELLESARK
PNAME = {p["key"]: p["sheet"] for p in PRODS}

# ---- Faste kostnader
FASTE_MASKIN = [
    ("m_vedlikehald", "Vedlikehald maskiner og reiskap", "6620-6629", 27820.2),
    ("m_drivstoff", "Drivstoff", "6250-6259", 18004.8),
    ("m_leige", "Maskinleige / leigekøyring / bruk av eigen traktor", "4550-4559, 6400-6409, 6450-6452", 84250.8),
    ("m_leasing", "Leasing maskiner", "6470, 6473", 0),
    ("m_avskr", "Avskrivingar maskiner og reiskap (driftsmessige)", "6010", 98462.99),
    ("m_andre", "Andre maskinkostnader (forsikring, småreiskap)", "6500-6529, 7503", 14251),
]
FASTE_BYGG = [
    ("b_vedlikehald", "Vedlikehald driftsbygningar og installasjonar", "6600-6604, 6640", 94769.39),
    ("b_avskr", "Avskrivingar bygningar og anlegg (driftsmessige)", "6000-6005, 6030, 6070-6092", 147609.12),
    ("b_leasing", "Leasing / leige bygningar", "6300-6301, 6410", 0),
    ("b_andre", "Andre bygningskostnader (vatn, renovasjon, reinhald, forsikring)", "6260-6262, 6320, 6340, 6360, 6390, 6541, 6563, 7501", 36721),
]
FASTE_ANDRE = [
    ("a_leidhjelp", "Leigd hjelp / løn / avløysar", "4570-4580, 5000-5599, 5700-5999, !5800-5899", 78672.8),
    ("a_arbeidsvederlag", "Arbeidsvederlag i selskap", "5600-5699", 0),
    ("a_refusjon", "Refusjon avløysar / sjukdom (-)", "3440-3442, 5800-5899", -62681),
    ("a_leige_jord", "Leige av jord, beite og mjølkekvote", "6471, 6472, 6475", 0),
    ("a_vedlikehald", "Vedlikehald jord, vegar, grøfter og anna", "6650-6699", 5000),
    ("a_energi", "Energi (straum, brensel) som ikkje er fordelt", "6200-6249", 23734.87),
    ("a_admin", "Administrasjon (rekneskap, kontor, telefon, kontingentar, forsikring)", "6700-6999, 7100-7499, 7500, 7502, 7509, 7600-7799", 156309.53),
    ("a_bil", "Bil og transportmiddel", "6011-6014, 7000-7099", 0),
    ("a_avskr", "Andre avskrivingar (driftsmessige)", "6015-6019", 0),
    ("a_annan", "Annan fast kostnad (resten av 5000–7999)", "", 66370.83),
]
INNMEK_ACC = "6030, 6541, 6563, 6640"
INNMEK_EX = {"ammekyr": 72685.58, "sau": 23039.63}
MASKIN_PCT_EX = {"silofor": 78139.31, "ammekyr": 108004.62, "sau": 30916.38}  # kr i eksempelet -> %


def faste_sheet():
    maskin_total_ex = sum(v for *_, v in FASTE_MASKIN)
    rows = [H("1. INPUT – Maskinkostnader (driftsmessige tal)")]
    for id, lab, acc, v in FASTE_MASKIN:
        rows.append(I(id, lab, "kr", v, src=acc, kind="fast", acc=acc, sign=1))
    rows.append(C("maskin_sum", "Sum maskinkostnader", "+".join(i[0] for i in FASTE_MASKIN)))
    rows.append(H("2. INPUT – Fordeling av maskinkostnader på produksjonane (%)"))
    for p in PRODS:
        ex = MASKIN_PCT_EX.get(p["key"], 0) / maskin_total_ex
        rows.append(I("pct_" + p["key"], "Del maskinkostnad – %s" % p["sheet"], "%", ex,
                      src="Skjønn / timeforbruk", kind="pct",
                      note="Eksempel: fordelt som i vedlagt analyse" if ex else ""))
    rows.append(C("pct_anna", "Del maskinkostnad – anna bruk (rest)",
                  "1-(" + "+".join("pct_" + k for k in PROD_KEYS) + ")", "%"))
    rows.append(H("3. UTREKNING – Maskinkostnader fordelt (kr)"))
    for p in PRODS:
        rows.append(C("maskin_" + p["key"], "Maskinkostnader – %s" % p["sheet"], "maskin_sum*pct_%s" % p["key"]))
    rows.append(C("maskin_anna", "Maskinkostnader – anna bruk", "maskin_sum*pct_anna"))
    rows.append(H("4. INPUT – Bygningskostnader"))
    for id, lab, acc, v in FASTE_BYGG:
        rows.append(I(id, lab, "kr", v, src=acc, kind="fast", acc=acc, sign=1))
    rows.append(C("bygg_sum", "Sum bygningskostnader", "+".join(i[0] for i in FASTE_BYGG)))
    rows.append(H("5. INPUT – Herav innandørsmekanisering, fordelt på husdyrproduksjonane (kr)"))
    for k in HUSDYR_KEYS:
        rows.append(I("innmek_" + k, "Innandørsmekanisering – %s" % PNAME[k], "kr", INNMEK_EX.get(k, 0),
                      src=INNMEK_ACC + " (fordelt med nøkkel)", kind="fast", acc=None, sign=1,
                      note="Herav-post: kostnadene ligg alt i bygnings- og andre faste kostnader, her blir dei berre fordelte"))
    rows.append(C("innmek_sum", "Sum innandørsmekanisering (herav, inngår i summane ovanfor)", "+".join("innmek_" + k for k in HUSDYR_KEYS)))
    rows.append(H("6. INPUT – Andre faste kostnader"))
    for id, lab, acc, v in FASTE_ANDRE:
        rows.append(I(id, lab, "kr", v, src=acc, kind="fast", acc=acc, sign=1))
    rows.append(C("andre_sum", "Sum andre faste kostnader", "+".join(i[0] for i in FASTE_ANDRE)))
    rows.append(H("7. OUTPUT – Sum faste kostnader"))
    rows += [
        C("faste_sum", "Sum faste kostnader jordbruk", "maskin_sum+bygg_sum+andre_sum", kpi=True,
          note="Innandørsmekanisering er ein herav-post og blir ikkje lagd til (same som eksempelanalysen)"),
        C("avskr_drift", "– herav driftsmessige avskrivingar", "m_avskr+b_avskr+a_avskr"),
        C("k_maskin_del", "Maskinkostnader i % av faste kostnader", "DIV(maskin_sum,faste_sum)", "%", kpi=True),
        C("k_bygg_del", "Bygningskostnader i % av faste kostnader", "DIV(bygg_sum,faste_sum)", "%", kpi=True),
    ]
    return {"key": "faste", "sheet": "Faste kostnader", "title": "Faste kostnader jordbruk",
            "desc": "Maskin-, bygnings- og andre faste kostnader. Maskinkostnader og innandørsmekanisering blir fordelt på produksjonane.",
            "rows": rows}


def grovfor_sheet():
    rows = [H("1. INPUT – Areal (dekar)")]
    rows += [
        I("overflate", "Overflatedyrka jord til eng og beite", "da", 0, src="Søknad om produksjonstilskot / kontoar 1071 mengd 2"),
        I("fulldyrka", "Fulldyrka jord til eng og beite", "da", 129, src="Søknad om produksjonstilskot / konto 1070 mengd 2"),
        I("kulturbeite", "Innmarksbeite / kulturbeite", "da", 169, src="Søknad om produksjonstilskot / konto 1072 mengd 2"),
        I("kb_utleigd", "Kulturbeite leigd ut", "da", 0, src="Leigeavtalar"),
    ]
    rows.append(H("2. INPUT – Dyr på utmarksbeite (tal dyr sleppt)"))
    rows += [
        I("utm_sau", "Sau på utmarksbeite (1,05 FEm/dag, 95 dagar)", "stk", 36, src="Søknad om tilskot til dyr på beite"),
        I("utm_lam", "Lam på utmarksbeite (1 FEm/dag, 90 dagar)", "stk", 0, src="Søknad om tilskot til dyr på beite"),
        I("utm_kyr", "Kyr på utmarksbeite (7,2 FEm/dag, 92 dagar)", "stk", 16, src="Søknad om tilskot til dyr på beite"),
        I("utm_ungdyr", "Ungdyr på utmarksbeite (4,8 FEm/dag, 92 dagar)", "stk", 21, src="Søknad om tilskot til dyr på beite"),
    ]
    rows.append(H("3. UTREKNING – Grovfôrareal"))
    rows += [
        C("da_sal", "Dekar til sal av silofôr og høy", "silofor.n*ACTIVE(silofor)+hoy.n*ACTIVE(hoy)", "da", "n1"),
        C("fulldyrka_eigen", "Fulldyrka til eigen grovfôrbasert husdyrproduksjon", "fulldyrka-da_sal", "da", "n1"),
        C("kb_eige", "Kulturbeite til eige bruk", "kulturbeite-kb_utleigd", "da", "n1"),
        C("da_mk", "Dekar grovfôr – maskinkostnadsgrunnlag (overflate/2 + fulldyrka + kulturbeite×10 %)",
          "overflate/2+fulldyrka_eigen+kb_eige*0.1", "da", "n1"),
        C("da_vk", "Korrigert grovfôrareal (overflate/1,25 + fulldyrka + kulturbeite/2,5)",
          "overflate/1.25+fulldyrka_eigen+kb_eige/2.5", "da", "n1", kpi=True),
        C("fem_utmark", "Fôreiningar (FEm) henta frå utmark",
          "utm_sau*1.05*95+utm_lam*1*90+utm_kyr*7.2*92+utm_ungdyr*4.8*92", "FEm", "n"),
        C("utm_da", "Utmark omrekna til «fulldyrka» dekar (900 FEm/da)", "fem_utmark/900", "da", "n1"),
    ]
    rows.append(H("4. UTREKNING – «Kueiningar» og eigen grovfôrkostnad pr produksjon (berre produksjonar med i totalen)"))
    for k in GROVFOR_KEYS:
        rows.append(C("ku_" + k, "Kueiningar – %s" % PNAME[k], "%s.kueiningar*ACTIVE(%s)" % (k, k), "stk", "n2"))
    rows.append(C("sum_ku", "Sum «kueiningar»", "+".join("ku_" + k for k in GROVFOR_KEYS), "stk", "n2", kpi=True))
    for k in GROVFOR_KEYS:
        rows.append(C("eg_" + k, "Eigen grovfôrkostnad – %s" % PNAME[k], "%s.sum_eigen_gf*ACTIVE(%s)" % (k, k)))
    rows.append(C("sum_eigen_alle", "Sum eigen grovfôrkostnad (fordelingsnøkkel areal)", "+".join("eg_" + k for k in GROVFOR_KEYS)))
    rows.append(C("sum_vk_alle", "Sum variabel grovfôrkostnad (utan plast)",
                  "+".join("%s.vk_gf*ACTIVE(%s)" % (k, k) for k in GROVFOR_KEYS)))
    rows.append(C("maskin_gf", "Maskinkostnader i grovfôrproduksjonane",
                  "+".join("faste.maskin_%s*ACTIVE(%s)" % (k, k) for k in GROVFOR_KEYS)))
    rows.append(C("db_gp_u", "DB grovfôrbaserte husdyr u/tilskot",
                  "+".join("%s.db_u*ACTIVE(%s)" % (k, k) for k in GROVFOR_KEYS)))
    rows.append(C("db_gp_m", "DB grovfôrbaserte husdyr m/tilskot",
                  "+".join("%s.db_m*ACTIVE(%s)" % (k, k) for k in GROVFOR_KEYS)))
    rows.append(H("5. OUTPUT – Nøkkeltal grovfôr"))
    rows += [
        C("varkost_pr_da", "Variabel grovfôrkostnad pr korrigert dekar", "DIV(sum_vk_alle,da_vk)", "kr/eining", kpi=True),
        C("maskin_pr_da", "Maskinkostnader pr dekar grovfôr", "DIV(maskin_gf,da_mk)", "kr/eining", kpi=True),
        C("korr_da_pr_ku", "Korrigerte dekar pr «kueining»", "DIV(da_vk,sum_ku)", "da", "n2", kpi=True),
        C("fulldyrka_pr_ku", "Dekar fulldyrka pr «kueining»", "DIV(fulldyrka_eigen,sum_ku)", "da", "n2", kpi=True),
        C("kb_pr_ku", "Dekar kulturbeite pr «kueining»", "DIV(kb_eige,sum_ku)", "da", "n2", kpi=True),
        C("fem_pr_ku", "FEm frå utmark pr «kueining»", "DIV(fem_utmark,sum_ku)", "FEm", "n", kpi=True),
        C("db_gp_u_da", "DB grovfôrbaserte husdyr u/tilskot pr korrigert dekar", "DIV(db_gp_u,da_vk)", "kr/eining", kpi=True),
        C("db_gp_m_da", "DB grovfôrbaserte husdyr m/tilskot pr korrigert dekar", "DIV(db_gp_m,da_vk)", "kr/eining", kpi=True),
    ]
    return {"key": "grovfor", "sheet": "Grovfôr", "title": "Grovfôr – areal, kueiningar og grovfôrkostnad",
            "desc": "Felles grovfôrgrunnlag for mjølk, oksar, mellomkalv, ammekyr, sau og geit.",
            "rows": rows}


def total_sheet():
    rows = [H("1. UTREKNING – Dekningsbidrag frå produksjonane (berre produksjonar med i totalen)")]
    rows.append(C("sum_inn", "Sum produksjonsinntekter", "+".join("%s.inn_total*ACTIVE(%s)" % (k, k) for k in PROD_KEYS)))
    rows.append(C("sum_var", "Sum variable kostnader", "+".join("%s.kost_total*ACTIVE(%s)" % (k, k) for k in PROD_KEYS)))
    rows.append(C("sum_db_u", "Sum DB utan tilskot", "+".join("%s.db_u*ACTIVE(%s)" % (k, k) for k in PROD_KEYS)))
    rows.append(C("sum_tilskot", "Sum tilskot fordelt på produksjonane", "+".join("%s.tilskot*ACTIVE(%s)" % (k, k) for k in PROD_KEYS)))
    rows.append(C("sum_db_m", "Sum DB med tilskot", "sum_db_u+sum_tilskot", kpi=True))
    rows.append(H("2. INPUT – Inntekter og kostnader som ikkje er fordelt på produksjon"))
    rows += [
        I("andre_inn", "Andre driftsinntekter, netto (ikkje fordelt på produksjon)", "kr", 228384.97,
          src="3090-3092, 3600-3999 o.l.", kind="fast", acc="3090-3093, 3100-3399, 3500-3999", sign=-1,
          note="Eksempel: DB totalt minus DB i produksjonane i vedlagt analyse (Prod.øk «Andre inntekter»)."),
        I("tilskot_ufordelt", "Tilskot som ikkje er fordelt på produksjon", "kr", 0, src="34xx", kind="fast", acc=None, sign=-1),
        I("andre_var", "Variable kostnader som ikkje er fordelt", "kr", 0, src="4xxx", kind="fast", acc=None, sign=1),
    ]
    rows.append(H("3. UTREKNING – Driftsoverskot jordbruk (driftsmessig)"))
    rows += [
        C("faste", "Sum faste kostnader", "faste.faste_sum"),
        C("do_drift", "Driftsoverskot jordbruk, driftsmessig", "sum_db_m+andre_inn+tilskot_ufordelt-andre_var-faste", kpi=True),
    ]
    rows.append(H("4. INPUT – Skattemessige tal frå rekneskapen (for avstemming)"))
    rows += [
        I("avskr_skatt", "Skattemessige avskrivingar", "kr", 8882.16, src="60xx i saldobalansen (før korrigering)", kind="fast"),
        I("beh_skatt", "Beholdningsendring buskap/varelager, skattemessig (+ = auke)", "kr", -58135,
          src="4090-4099, 4190-4199, 4290-4299, 4690-4699 (før korrigering)", kind="fast"),
        I("beh_drift", "Beholdningsendring buskap/varelager, driftsmessig (+ = auke)", "kr", 3740,
          src="Etter korrigering til driftsmessige verdiar", kind="fast"),
        I("do_skatt_rekneskap", "Driftsoverskot jordbruk i rekneskapen (skatt)", "kr", 403480.85,
          src="Næringsoppgåve / resultatrekneskap", kind="fast"),
    ]
    rows.append(H("5. UTREKNING – Frå skattemessig til driftsmessig resultat"))
    rows += [
        C("avskr_drift", "Driftsmessige avskrivingar", "faste.avskr_drift"),
        C("korr_avskr", "Korrigering avskrivingar (skatt – drift)", "avskr_skatt-avskr_drift"),
        C("korr_beh", "Korrigering beholdningsendring (drift – skatt)", "beh_drift-beh_skatt"),
        C("do_skatt_berekna", "Driftsoverskot skattemessig, berekna", "do_drift-korr_avskr-korr_beh"),
        C("avvik", "Avvik mot rekneskapen (bør vere 0)", "do_skatt_rekneskap-do_skatt_berekna",
          note="Avvik kjem av postar som ikkje er lagt inn i analysen. Eksempelanalysen har same avvik (−329,66)"),
    ]
    rows.append(H("6. INPUT/UTREKNING – Nettoinntekt"))
    rows += [
        I("do_andre", "Driftsoverskot andre næringar (skog, leigekøyring o.l.)", "kr", 117592.22, src="Avdeling 30/70 o.l.", kind="fast"),
        I("renteinn", "Renteinntekter", "kr", 117, src="80xx", kind="fast", acc="8000-8099", sign=-1),
        I("renteut", "Renteutgifter", "kr", 190284.05, src="8100-8179", kind="fast", acc="8100-8179", sign=1),
        C("netto", "Nettoinntekt næring (driftsmessig)", "do_drift+do_andre+renteinn-renteut", kpi=True),
    ]
    rows.append(H("7. OUTPUT – Nøkkeltal totaløkonomi"))
    rows += [
        C("k_db_grad", "DB m/tilskot i % av produksjonsinntekter og tilskot", "DIV(sum_db_m,sum_inn+sum_tilskot)", "%", kpi=True),
        C("k_tilskot_del", "Tilskot i % av DB m/tilskot", "DIV(sum_tilskot,sum_db_m)", "%", kpi=True),
        C("k_faste_del", "Faste kostnader i % av DB m/tilskot", "DIV(faste,sum_db_m)", "%", kpi=True),
        C("k_rente_del", "Renteutgifter i % av driftsoverskot", "DIV(renteut,do_drift)", "%", kpi=True),
    ]
    return {"key": "total", "sheet": "Totaløkonomi", "title": "Totaløkonomi",
            "desc": "Summerer dekningsbidraget frå alle produksjonar som er med, trekkjer frå faste kostnader og avstemmer mot skattemessig resultat.",
            "rows": rows}


COMMON = [faste_sheet(), grovfor_sheet(), total_sheet()]

# =================================================================== Kopling konto -> produksjon (Tripletex)
# Rekkjefølgje: første regel som passar vinn. "prod" kan vere ei liste – første aktive produksjon blir brukt.
# "suffix" = produktkode i Tripletex-kontoen (t.d. 1454.20.541 -> 541).
SUFFIX_PROD = [
    ("511-512", ["mjolk"]), ("514-515", ["mjolk", "ammekyr", "oksar"]), ("516-518", ["oksar", "mellomkalv", "mjolk", "ammekyr"]),
    ("521-526", ["ammekyr"]), ("531-534", ["smagris", "purkering"]), ("535-538", ["slaktegris", "smagris"]),
    ("541-543", ["sau"]), ("551-553", ["geit"]), ("561-562", ["egg"]), ("564-568", ["kylling"]),
    ("411-414", ["korn"]), ("421-423", ["korn"]), ("431", ["potet"]), ("441", ["fgront"]),
    ("471", ["hoy", "POOL"]), ("472", ["silofor", "POOL"]), ("473", ["korn"]), ("482", ["mjolk", "geit"]),
    ("484", ["sau"]), ("486", ["egg"]),
]
ACCOUNT_PROD = [
    ("3000-3008, 3096, 3410, 4000, 4291-4292", ["korn"]),
    ("3010-3011, 3421, 4010, 4091, 4293", ["potet"]),
    ("3013, 3015, 3020, 3039, 3420, 4012, 4294", ["fgront", "gulrot", "kalrot"]),
    ("3014", ["gulrot", "fgront"]),
    ("3018, 4018", ["tomat"]),
    ("3019, 4019", ["agurk"]),
    ("3094", ["hoy"]),
    ("3095", ["silofor"]),
    ("3071-3073, 3450, 3455", ["mjolk"]),
    ("3451", ["oksar", "mellomkalv", "mjolk", "ammekyr"]),
    ("3041, 3051, 4021, 4041, 4521-4522, 4531, 4534, 4610, 4691, 4131", ["mjolk", "oksar", "mellomkalv", "ammekyr"]),
    ("3042, 3052, 3452, 4620, 4692", ["ammekyr"]),
    ("3043, 3053, 3453, 4023, 4043, 4523, 4535, 4630, 4693, 4133", ["smagris", "slaktegris", "purkering"]),
    ("3044, 3054, 3074, 3454, 4024, 4524, 4533, 4134, 4640, 4694", ["sau"]),
    ("3045, 3055, 3075, 3079, 4130, 4650, 4695", ["geit"]),
    ("3076, 3097, 3456, 4026, 4046, 4526, 4536, 4136", ["egg", "kylling"]),
    ("3046, 3056, 4660, 4696", ["egg", "kylling"]),
    # Fellespott (fordelast med nøkkel på grovfôr-/husdyrproduksjonane)
    ("3413, 3430, 3457, 3460, 3470, 3490, 3499, 4001, 4020, 4030-4031, 4040, 4050, 4057, 4080, 4090, 4093-4095, 4098, 4110, 4120, 4140, 4191-4194, 4297, 4520, 4530, 4532, 4539, 4560-4569", ["POOL"]),
    ("4550-4559", ["FASTE"]),
    ("3440-3442", ["FASTE"]),
    # Andre tilskot utan eigen kopling: fellespott (fordelast med nøkkel)
    ("3400-3439, 3443-3499", ["POOL"]),
    ("6030, 6541, 6563, 6640", ["INNMEK"]),
    ("6480", ["purkering", "FASTE"]),
    ("4570-4580, 5000-7999", ["FASTE"]),
    ("3090-3093, 3100-3399, 3500-3999", ["TOTAL"]),
    ("8000-8179", ["TOTAL"]),
]


def spec_json():
    return {"prods": PRODS, "common": COMMON, "prod_keys": PROD_KEYS, "grovfor_keys": GROVFOR_KEYS,
            "husdyr_keys": HUSDYR_KEYS, "suffix_prod": SUFFIX_PROD, "account_prod": ACCOUNT_PROD,
            "faste_maskin": FASTE_MASKIN, "faste_bygg": FASTE_BYGG, "faste_andre": FASTE_ANDRE,
            "innmek_acc": INNMEK_ACC}
