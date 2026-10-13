# Nøkkeltalsanalyse jordbruk

Produksjonsvis nøkkeltalsanalyse for jordbruket, bygd på eksempelanalysen «Analyse 2025» (Duett-import) og kontoplanen frå Tripletex.

## Innhald

| Fil | Kva |
|---|---|
| `excel/Analyse_jordbruk_produksjonar.xlsx` | Arbeidsbok med éi fane per produksjon (korn, silofôr, høy, potet, agurk, tomat, frilandsgrønsak, gulrot, kålrot, mjølk, oksar, mellomkalv, ammekyr, smågris, purkering, slaktegris, sau, geit, egg, slaktekylling) + Grovfôr, Faste kostnader, Totaløkonomi, Oversikt og Kontokopling. Kvar fane: input → utrekning → per eining → nøkkeltal, med formelen forklart i ord (kolonne E) og som Excel-formel (kolonne F). |
| `app/analyse.html` | Nettside (éi fil, bokmål/nynorsk, lys/mørk modus) med seks steg: 1) import av saldobalanse og/eller hovudbok (Tripletex – førebels fiktive tal – eller CSV/Excel; hovudboka gir mengder per konto, tal per månad, avstemming mot saldobalansen og visning av posteringar), 2) driftsmessige verdiar på varelager/buskap, 3) driftsmessige avskrivingar, 4) korrigeringar i fem obligatoriske understeg (produksjonar, produksjonsdata, fordelingsnøklar, flytt beløp, kontroll), 5) tekstinnhald (automatisk tekst per seksjon, kan redigerast), 6) PDF-rapport (vektor-PDF med pdfmake, førehandsvisning og nedlasting). «Innstillingar» har to faner: **Kontoplan** (heile kontoplanen med avdelingar; avdeling og kopling til produksjon kan overstyrast per konto, også før import og direkte frå importtabellen) og **Konteringsoppsett** (kva kontoar som inngår i kvar post, sortert etter formel – kontoar kan fjernast med × og leggjast til med +, formlar kan endrast i eige vindauge, og alt kan tilbakestillast til standard). |
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
* Endra formlar og kontoar i Konteringsoppsett gjeld berre nettsida, ikkje Excel-arbeidsboka.
* Bokmålsetikettane blir laga frå spesifikasjonen med `tools/bokmal.py`.
* Marginalanalysane i eksempelfila er ikkje bygde om.
