# Nøkkeltalsanalyse jordbruk

Produksjonsvis nøkkeltalsanalyse for jordbruket, bygd på eksempelanalysen «Analyse 2025» (Duett-import) og kontoplanen frå Tripletex.

## Innhald

| Fil | Kva |
|---|---|
| `excel/Analyse_jordbruk_produksjonar.xlsx` | Arbeidsbok med éi fane per produksjon (korn, silofôr, høy, potet, agurk, tomat, frilandsgrønsak, gulrot, kålrot, mjølk, oksar, mellomkalv, ammekyr, smågris, purkering, slaktegris, sau, geit, egg, slaktekylling) + Grovfôr, Faste kostnader, Totaløkonomi, Oversikt og Kontokopling. Kvar fane: input → utrekning → per eining → nøkkeltal, med formelen forklart i ord (kolonne E) og som Excel-formel (kolonne F). |
| `app/analyse.html` | Nettside (éi fil, bokmål/nynorsk, lys/mørk modus) med seks steg: 1) import av saldobalanse og/eller hovudbok (Tripletex – førebels fiktive tal – eller CSV/Excel; hovudboka gir mengder per konto, tal per månad, avstemming mot saldobalansen og visning av posteringar), 2) driftsmessige verdiar på varelager/buskap, 3) driftsmessige avskrivingar, 4) korrigeringar i fem obligatoriske understeg (produksjonar, produksjonsdata, fordelingsnøklar, flytt beløp, kontroll), 5) tekstinnhald (automatisk tekst per seksjon, kan redigerast), 6) PDF-rapport (vektor-PDF med pdfmake, førehandsvisning og nedlasting). «Innstillingar» har to faner: **Kontoplan** (heile kontoplanen med avdelingar; avdeling og kopling til produksjon kan overstyrast per konto, også før import og direkte frå importtabellen) og **Konteringsoppsett** (kva kontoar som inngår i kvar post, sortert etter formel – kontoar kan fjernast med × og leggjast til med +, formlar kan endrast i eige vindauge, og alt kan tilbakestillast til standard). |
| `tools/spec.py` | Felles definisjon av alle postar, kontokoplingar og formlar. Både Excel-fila og nettsida blir genererte frå denne, så formlane er like. |
| `tools/build_excel.py`, `tools/build_html.py` | Byggjeskript. |

## Nettside (Vercel)

Appen er ei statisk nettside i mappa `public/` (`index.html` + Excel-arbeidsboka til nedlasting). `vercel.json` peikar Vercel dit, så det trengst ingen byggesteg på Vercel.

1. Gå til vercel.com → **Add New… → Project** og vel GitHub-repoet `analyse-jordbruk`.
2. La **Framework Preset** stå på «Other». Root Directory er repoet (`./`). Vercel les `vercel.json` og publiserer `public/`.
3. Trykk **Deploy**. Kvar push til produksjonsgreina (vanlegvis `main`) gir ny versjon automatisk.

Etter endringar i `tools/` eller `app/src/`: køyr byggeskripta under og commit `public/` saman med resten.

Arbeid og innstillingar blir lagra i nettlesaren til kvar brukar (localStorage). Bruk «Lagre prosjekt» og **Innstillingar → Eksporter/Importer innstillingar** for å dele mellom maskiner og kollegaer.

## Bygg på nytt

```bash
pip install openpyxl
python tools/build_excel.py   # -> excel/Analyse_jordbruk_produksjonar.xlsx
python tools/build_html.py    # -> app/analyse.html (frå app/src/template.html)
```

## Merknader

* Eksempeltala for silofôr, ammekyr, sau, grovfôrareal og faste kostnader i Excel-fila er henta frå vedlagt analyse 2025, og tala stemmer med den: DB pr vf. sau u/tilskot 315 kr, DB pr dekar grovfôr 430 kr, DB pr ammeku 6 457 kr, sum faste kostnader 789 296 kr, driftsoverskot jordbruk 228 496 kr og kontrollavvik −330 kr (same avvik som eksempelet). Dei andre produksjonane har fiktive eksempeltal og står som «N» (ikkje med i totalen) på fana Oversikt.
* Fordelingsnøklar i nettsida (steg 4.4) har seks grupper med automatisk grunnlag: areal (gjødsel, kalk, plantevern, såfrø, arealtilskot – husdyr får arealbehov etter fôrnorm), hausta grovfôr (vinterfôr), felles husdyrkostnader og tilskot utan kopling (etter det som er ført direkte), maskin (maskindekar, med «anna bruk» som eiga linje) og innandørsmekanisering (fôr innandørs). Fôrnormene er henta frå marginalarka i eksempelet og kan endrast under Produksjonsdata. Alle nøklar kan overstyrast.
* Innandørsmekanisering er ein herav-post av faste kostnader (som i eksempelet), ikkje eit tillegg. Purkeleige er variabel kostnad i purkeringen.
* «Variabel grovfôrkostnad pr dekar» er utan rundballeplast, som i eksempelet (544 kr mot 542 kr i eksempelet; skilnaden er ein inkonsistens i eksempelfila).
* Ikkje bygd om frå eksempelfila: marginalanalysane, «Fri»-arka, fleirårs- og samanlikningsstatistikk, normbasert grovfôr-/kraftfôrbehov og jordbrukskonto/skog/samdrift i totaløkonomien.
* Nettsida treng internett for PDF (pdfmake) og opplasting av Excel-filer (SheetJS).
* Arbeidet blir lagra automatisk i nettlesaren og kan lagrast/opnast som prosjektfil (JSON). Kontokoplingane er lagra for seg og gjeld for alle prosjekt.
* Endra formlar og kontoar i Konteringsoppsett gjeld berre nettsida, ikkje Excel-arbeidsboka.
* Bokmålsetikettane blir laga frå spesifikasjonen med `tools/bokmal.py`.

## Bokføringsoppdrag i Teams

Mappen [`spfx-oppdrag/`](spfx-oppdrag/) inneholder en SharePoint-app (SPFx) med en felles oppdragsliste som vises i Teams. Se [spfx-oppdrag/README.md](spfx-oppdrag/README.md) for installasjon.
