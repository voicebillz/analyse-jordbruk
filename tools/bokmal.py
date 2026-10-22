# -*- coding: utf-8 -*-
"""Omset nynorske etikettar i spec.py til bokmål (brukt av build_html.py)."""
import re

L = "A-Za-zÆØÅæøåÔô"

# Fraser først, deretter einskildord. Rekkjefølgja tel.
PHRASES = [
    ("anna bruk", "annen bruk"), ("driftsmessige tal)", "driftsmessige tall)"), ("vinterfôra sau", "vinterfôrede sauer"), ("Skattemessige tal", "Skattemessige tall"), ("Rekneskapstal", "Regnskapstall"),
    ("Nøkkeltal", "Nøkkeltall"), ("Sal av", "Salg av"), ("sal av", "salg av"), ("tal dyr sleppt", "antall dyr sluppet"),
    ("Kulturbeite til eige bruk", "Kulturbeite til eget bruk"), ("som ikkje er fordelt", "som ikke er fordelt"),
    ("Kviger/oksar", "Kviger/okser"), ("Oksar/kalvar", "Okser/kalver"), ("Kalvar/ungdyr", "Kalver/ungdyr"),
    ("Bil og transportmiddel", "Bil og transportmidler"), ("bør vere 0", "bør være 0"),
]
WORDS = {
    "Tal": "Antall", "tal": "antall", "eining": "enhet", "einingar": "enheter", "Kueiningar": "Kuenheter",
    "kueiningar": "kuenheter", "kueining": "kuenhet", "Fôreiningar": "Fôrenheter",
    "tilskot": "tilskudd", "Tilskot": "Tilskudd", "TILSKOT": "TILSKUDD", "Produksjonstilskot": "Produksjonstilskudd",
    "distriktstilskot": "distriktstilskudd", "utan": "uten", "Utan": "Uten",
    "eigen": "egen", "Eigen": "Egen", "eigenprodusert": "egenprodusert", "eige": "eget",
    "Mjølk": "Melk", "mjølk": "melk", "Mjølkeproduksjon": "Melkeproduksjon", "mjølkeproduksjonen": "melkeproduksjonen",
    "mjølkepris": "melkepris", "Mjølkegeit": "Melkegeit", "mjølkekvote": "melkekvote",
    "Sauehald": "Sauehold", "Livdyrsal": "Livdyrsalg", "Sal": "Salg", "sal": "salg", "Seld": "Solgt", "seld": "solgt",
    "selde": "solgte", "selt": "solgt", "slakta": "slaktede", "driftsbygningar": "driftsbygninger", "produksjonar": "produksjoner", "produksjonane": "produksjonene", "berre": "bare", "trekkjer": "trekker", "utanom": "utenom", "Eksempeltal": "Eksempeltall", "ligg": "ligger", "kastratar": "kastrater", "mengd": "mengde",
    "kjøt": "kjøtt", "kjøtpris": "kjøttpris", "frå": "fra", "Frå": "Fra", "inngåande": "inngående",
    "Endra": "Endret", "grunna": "grunnet", "Del": "Andel", "anna": "annet",
    "leige": "leie", "Leige": "Leie", "leigd": "leid", "Leigd": "Leid", "Beiteleige": "Beiteleie",
    "Maskinleige": "Maskinleie", "leigekøyring": "leiekjøring", "Purkeleige": "Purkeleie", "leigeinntekter": "leieinntekter",
    "Vedlikehald": "Vedlikehold", "Avskrivingar": "Avskrivninger", "avskrivingar": "avskrivninger",
    "reiskap": "redskap", "småreiskap": "småredskap", "bygningar": "bygninger", "vegar": "veier",
    "løn": "lønn", "avløysar": "avløser", "sjukdom": "sykdom", "straum": "strøm", "rekneskap": "regnskap",
    "rekneskapen": "regnskapet", "kontingentar": "kontingenter", "Driftsoverskot": "Driftsoverskudd",
    "driftsoverskot": "driftsoverskudd", "næringar": "næringer", "ikkje": "ikke", "berekna": "beregnet",
    "Overflatedyrka": "Overflatedyrket", "Fulldyrka": "Fulldyrket", "fulldyrka": "fulldyrket", "henta": "hentet",
    "omrekna": "omregnet", "vinterfôra": "vinterfôrede", "Kalvar": "Kalver", "kalvar": "kalver", "Oksar": "Okser",
    "oksar": "okser", "kyllingar": "kyllinger", "Plantevernmiddel": "Plantevernmidler",
    "Konserveringsmiddel": "Konserveringsmidler", "Kvoteoppfylling": "Kvoteoppfyllelse", "rundballar": "rundballer",
    "siloballar": "siloballer", "Frilandsgrønsak": "Frilandsgrønnsak", "Frilandsgrønsaker": "Frilandsgrønnsaker",
    "grønsaker": "grønnsaker", "auke": "økning", "tenester": "tjenester", "teljingar": "tellinger",
    "korrigeringar": "korrigeringer", "lenka": "lenket", "UTREKNING": "BEREGNING", "arket": "arket",
    "Geiter": "Geiter", "Sauer": "Sauer", "kjøpte": "kjøpte",
}


def nb(text):
    if not text:
        return text
    for a, b in PHRASES:
        text = text.replace(a, b)
    pat = re.compile(r"(?<![%s])(%s)(?![%s])" % (L, "|".join(sorted(map(re.escape, WORDS), key=len, reverse=True)), L))
    return pat.sub(lambda m: WORDS[m.group(1)], text)
