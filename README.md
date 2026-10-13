# Nøkkeltalsanalyse jordbruk

Produksjonsvis nøkkeltalsanalyse for jordbruket, bygd på eksempelanalysen «Analyse 2025» (Duett-import) og kontoplanen frå Tripletex.

## Innhald

| Fil | Kva |
|---|---|
| `excel/Analyse_jordbruk_produksjonar.xlsx` | Arbeidsbok med éi fane per produksjon (korn, silofôr, høy, potet, agurk, tomat, frilandsgrønsak, gulrot, kålrot, mjølk, oksar, mellomkalv, ammekyr, smågris, purkering, slaktegris, sau, geit, egg, slaktekylling) + Grovfôr, Faste kostnader, Totaløkonomi, Oversikt og Kontokopling. Kvar fane: input → utrekning → per eining → nøkkeltal, med formelen forklart i ord (kolonne E) og som Excel-formel (kolonne F). |
| `app/analyse.html` | Nettside (éi fil, bokmål/nynorsk, lys/mørk modus) med seks steg: 1) import av saldobalanse (Tripletex – førebels fiktive tal – eller CSV/Excel), 2) driftsmessige verdiar på varelager/buskap, 3) driftsmessige avskrivingar, 4) korrigeringar i fem obligatoriske understeg (produksjonar, produksjonsdata, fordelingsnøklar, flytt beløp, kontroll), 5) tekstinnhald (automatisk tekst per seksjon, kan redigerast), 6) PDF-rapport (vektor-PDF med pdfmake, førehandsvisning og nedlasting). «Innstillingar» i toppen held koplinga konto → produksjon; ho kan setjast før import og endrast direkte frå importtabellen. |
| `tools/spec.py` | Felles definisjon av alle postar, kontokoplingar og formlar. Både Excel-fila og nettsida blir genererte frå denne, så formlane er like. |
| `tools/build_excel.py`, `tools/build_html.py` | Byggjeskript. |

## Bygg på nytt

```bash
pip install openpyxl
python tools/build_excel.py   # -> excel/Analyse_jordbruk_produksjonar.xlsx
python tools/build_html.py    # -> app/analyse.html (frå app/src/template.html)
```

## Merknader

* Eksempeltala for silofôr, ammekyr, sau, grovfôrareal og faste kostnader i Excel-fila er henta frå vedlagt analyse 2025, og nøkkeltala stemmer med den (t.d. DB pr vf. sau u/tilskot 315 kr, DB pr dekar grovfôr 430 kr, DB pr ammeku 6 457 kr). Dei andre produksjonane har fiktive eksempeltal og står som «N» (ikkje med i totalen) på fana Oversikt.
* Nettsida treng internett for PDF (pdfmake) og opplasting av Excel-filer (SheetJS).
* Arbeidet blir lagra automatisk i nettlesaren og kan lagrast/opnast som prosjektfil (JSON). Kontokoplingane er lagra for seg og gjeld for alle prosjekt.
* Bokmålsetikettane blir laga frå spesifikasjonen med `tools/bokmal.py`.
* Marginalanalysane i eksempelfila er ikkje bygde om.
