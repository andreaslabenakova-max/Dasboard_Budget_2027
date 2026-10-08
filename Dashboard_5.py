# -*- coding: utf-8 -*-

"""
DASHBOARD STÁTNÍ ROZPOČET 2027 – WEBOVÝ PROHLÍŽEČ
=================================================

Po spuštění:

    python Dashboard_verejne_finance_2027_web.py

se:

1. načtou všechny Excel soubory (výdaje, příjmy, deficit,
   obnovitelné zdroje, výdaje na obranu),
2. vytvoří interaktivní HTML dashboard,
3. HTML se uloží,
4. dashboard se automaticky otevře v prohlížeči.

DONUTY (stejně jako v PNG pro Canvu)
  Příjmy:
    - položka 13 jako jedna výseč „Daně a poplatky z vybraných
      činností a služeb“ (ROZPAD_13 = True -> rozpad na 132–138,
      podpoložky se berou z bloku POD řádkem 13),
    - navíc řádek 46 (nedaňové a kapitálové příjmy, transfery),
    - položky pod SLOUCIT_POD % sloučené do „Ostatní daně a poplatky“,
    - názvy bez čísel kódů.
  Výdaje:
    - malé položky skupiny 33–39 sloučené do
      „Ostatní služby pro obyvatelstvo“, zbylé mají různé odstíny.
  Obojí: procenta za názvem v legendě.

OBRANA (stejně jako ve skriptu pro Canvu)
  - Česko a USA – cesta k cíli NATO 5 % HDP do roku 2035.

Instalace:

    pip install pandas numpy plotly openpyxl
"""

from pathlib import Path
import json
import re
import webbrowser
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go


# ============================================================
# SOUBORY
# ============================================================

SLOZKA = Path(
    r"C:\Users\SlabenakovaA\OneDrive - KPS CR\Plocha\Python"
)

SOUBOR_VYDAJE = SLOZKA / "Rozpocet_Odvetvi.xlsx"

SOUBOR_PRIJMY = SLOZKA / "Dane_Extrakt.xlsx"

SOUBOR_DEFICIT = SLOZKA / "Deficit.xlsx"

SOUBOR_OZE = SLOZKA / "Obnovitelne_Zdroje.xlsx"

SOUBOR_OBRANA = SLOZKA / "Vydaje_Obrana_.xlsx"

SHEET_VYDAJE = "Výdaje 2027-odvětvové"

SHEET_PRIJMY = "Tab.1 - příjmy"

SHEET_DEFICIT = "List1"

SHEET_OZE = "Sheet 1"

SHEET_OBRANA = "Vojenské výdaje"


# ============================================================
# VÝSTUPNÍ HTML
# ============================================================

VYSTUP = (
    SOUBOR_PRIJMY.parent
    / "Dashboard_verejne_finance_2027.html"
)


# ============================================================
# BARVY
# ============================================================

MODRA = "#0055A0"
CERVENA = "#E32219"
CERVENA_2 = "#D13D41"     # ostatní sloupce v grafu obrany
SVETLE_MODRA = "#98C3D0"
ZELENA = "#B3C0A7"
TMAVE_SEDA = "#535C55"
SEDA = "#2E808C"
ZLATA = "#84754E"
SVETLE_SEDA = "#8A8D8F"
TMAVE_MODRA = "#142B53"
SVETLE_ZLATA = "#C9B98F"  # nedaňové příjmy (řádek 46)

# Odstíny pro položky skupiny 3 (služby pro obyvatelstvo) ve výdajích
ODSTINY_SKUPINA3 = ["#98C3D0", "#5E97AB", "#C7DEE6", "#3E7487"]

# Světlejší odstíny pro výhled (graf deficitu)
SVETLE_CERVENA = "#F19A94"
SVETLE_MODRA_ODHAD = "#7FA9D0"

# Zelená pro šipky u "dobrého" vývoje
ZELENA_TMAVA = "#2E7D4F"

TEXT = "#222222"
TEXT_LIGHT = "#777777"
LINE = "#D0D0D0"
BACKGROUND = "#FFFFFF"


# ============================================================
# DONUTY – OBSAH (stejně jako v PNG pro Canvu)
# ============================================================

# Příjmy: položka 13 rozpadlá na podpoložky 132–138 (True),
# nebo jako jedna výseč (False)
ROZPAD_13 = False

# Příjmy: nová položka – číslo řádku tak, jak ho ukazuje Excel
RADEK_NOVA_POLOZKA = 46

# Příjmy: položky menší než tolik % z celku se sloučí do jedné
# výseče (jinak jsou v donutu neviditelné). 0 = nic neslučovat.
SLOUCIT_POD = 1.0
NAZEV_SLOUCENE = "Ostatní daně a poplatky"

# Výdaje: položky skupiny 3 (33–39) menší než tolik % z celku
# se sloučí do jedné výseče. 0 = nic neslučovat.
VYDAJE_SLOUCIT_POD = 1.0
VYDAJE_NAZEV_SLOUCENE = "Ostatní služby pro obyvatelstvo"

# Procenta ve výsečích – u menších výsečí se nepíší
# (jsou v legendě za názvem)
DONUT_MIN_PODIL_TEXT = 2.0


# ============================================================
# ROZMĚRY DONUTŮ (v pixelech, stejné pro oba grafy)
# ============================================================

DONUT_PLOCHA = 330      # výška plochy, ve které se kreslí donut = jeho průměr
TITULEK_VYSKA = 50      # horní okraj pro nadpis grafu
RADEK_LEGENDY = 24      # výška jednoho řádku legendy (pro písmo 14)
STRED_ODSTUP = 30       # rozestup řádků textu uprostřed donutu (px)

# Nadpisy všech grafů: první řádek vždy ve stejné výšce
# od horního okraje panelu (bez ohledu na to, jak velký
# horní okraj graf má nebo kolik řádků nadpis má)
TITULEK_POZICE = dict(
    y=1,
    yref="container",
    yanchor="top",
    pad=dict(t=14)
)

# Výška grafu deficitu
DEFICIT_VYSKA = 470

# Kolik posledních let zobrazit v grafu deficitu
DEFICIT_POCET_LET = 10


# ============================================================
# OBRANA
# ============================================================

OBRANA_CIL = 5.0            # cíl NATO (% HDP)
OBRANA_CIL_ROK = 2035
OBRANA_VYCHOZI_ROK = 2025   # sloupec v Excelu se stavem
OBRANA_ZVYRAZNENY_ROK = 2027
OBRANA_VYSKA = 470          # výška grafu obrany (px)


# ============================================================
# MOBILNÍ ZOBRAZENÍ
#
# Pod touto šířkou okna (px) se grafy přepnou do mobilní
# úpravy: menší písmo, zkrácené legendy, méně popisků
# a vypnuté přibližování tahem prstu.
# ============================================================

MOBIL_SIRKA = 650

# Maximální délka názvu v legendě donutu na mobilu
# (plný název zůstává v bublině po klepnutí)
MOBIL_LEGENDA_ZNAKU = 38

# Výška grafů na mobilu (px) – aby nebyly příliš protáhlé
MOBIL_VYSKA_BK = 420        # běžné a kapitálové výdaje
MOBIL_VYSKA_DEFICIT = 380   # vývoj salda
MOBIL_VYSKA_OBRANA = 380    # graf obrany


# ============================================================
# UKAZATELE NAHOŘE (KARTY SE ŠIPKAMI)
#
# Směr šipky se určí podle znaménka hodnoty.
# "růst je dobrý?" určuje barvu šipky:
#   True  -> růst = zelená, pokles = červená
#   False -> růst = červená, pokles = zelená (např. inflace)
# "porovnat s inflací?" přidá rozdíl proti inflaci v p. b.
# ============================================================

INFLACE_2027 = 2.7

UKAZATELE = [
    # (nadpis, hodnota v %, popisek, růst je dobrý?, porovnat s inflací?, ikona)
    ("INFLACE 2027", INFLACE_2027, "průměrná míra inflace", False, False, "inflace"),
    ("VYSOKOŠKOLSKÉ VZDĚLÁVÁNÍ", -3.6, "meziroční změna 2027", True, False, "vs"),
    ("VÝZKUM A VÝVOJ (R&D)", -3.0, "meziroční změna 2027", True, False, "vyzkum"),
]

# Ikony karet (SVG, kreslené čarou)
IKONY = {
    "inflace": (
        '<line x1="19" y1="5" x2="5" y2="19"/>'
        '<circle cx="6.5" cy="6.5" r="2.5"/>'
        '<circle cx="17.5" cy="17.5" r="2.5"/>'
    ),
    "vs": (
        '<path d="M22 10 12 5 2 10l10 5 10-5z"/>'
        '<path d="M6 12v5c3 3 9 3 12 0v-5"/>'
        '<path d="M22 10v6"/>'
    ),
    "vyzkum": (
        '<path d="M9 3h6"/>'
        '<path d="M10 3v6L4.5 19a1.5 1.5 0 0 0 1.3 2h12.4a1.5 1.5 0 0 0 1.3-2L14 9V3"/>'
        '<path d="M7 15h10"/>'
    ),
}


# ============================================================
# STÁTNÍ ROZPOČET 2027 – HLAVNÍ ČÍSLA (v Kč)
#
# OBSLUHA_DLUHU = None -> karta zůstane prázdná ("doplní se").
# Až budou čísla, stačí zadat částku v Kč.
# ============================================================

PRIJMY_SR_CELKEM = 2_198_649_791_455

VYDAJE_SR_CELKEM = 2_584_649_791_455

SALDO_SR = PRIJMY_SR_CELKEM - VYDAJE_SR_CELKEM

OBSLUHA_DLUHU = 129_966_880_286

# Od kolika % podílu obsluhy dluhu na výdajích SR se karta
# obarví červeně (pod prahem zůstává zlatá)
PRAH_DLUH_CERVENA = 5.0


# ============================================================
# OBNOVITELNÉ ZDROJE (EUROSTAT nrg_cb_pem)
#
# Podíl OZE = roční součet "Renewables and biofuels" (GWh)
#             / roční součet "Total" (GWh) za 12 měsíců.
# OZE_ROK = None -> vezme se poslední rok, za který má EU-27
# všech 12 měsíců. Lze zadat i ručně, např. OZE_ROK = 2025.
# ============================================================

OZE_ROK = 2026

# Které měsíce roku použít:
#   None        -> všechny měsíce, které mají všechny země
#   [1, 2, 3]   -> 1. čtvrtletí, [4, 5, 6] -> 2. čtvrtletí atd.
OZE_MESICE = None

EU_AGREGAT = "European Union - 27 countries (from 2020)"

ZEME_EU = {
    "Belgium": "Belgie",
    "Bulgaria": "Bulharsko",
    "Czechia": "Česko",
    "Denmark": "Dánsko",
    "Germany": "Německo",
    "Estonia": "Estonsko",
    "Ireland": "Irsko",
    "Greece": "Řecko",
    "Spain": "Španělsko",
    "France": "Francie",
    "Croatia": "Chorvatsko",
    "Italy": "Itálie",
    "Cyprus": "Kypr",
    "Latvia": "Lotyšsko",
    "Lithuania": "Litva",
    "Luxembourg": "Lucembursko",
    "Hungary": "Maďarsko",
    "Malta": "Malta",
    "Netherlands": "Nizozemsko",
    "Austria": "Rakousko",
    "Poland": "Polsko",
    "Portugal": "Portugalsko",
    "Romania": "Rumunsko",
    "Slovenia": "Slovinsko",
    "Slovakia": "Slovensko",
    "Finland": "Finsko",
    "Sweden": "Švédsko",
}

CESKO = "Česko"


# ============================================================
# BĚŽNÉ A KAPITÁLOVÉ VÝDAJE (v mld. Kč)
#
# Hodnoty jsou v miliardách Kč – zadávají se přímo tak,
# jak jsou (např. 1798.9 = 1 798,9 mld. Kč).
# ============================================================

YEARS = [2022, 2023, 2024, 2025, 2026, 2027]

BEZNE = np.array([
    1798.9,
    1992.2,
    2024.3,
    2111.8,
    2164.6,
    2293.8
])

KAPITALOVE = np.array([
    186.0,
    210.5,
    212.5,
    260.0,
    263.2,
    290.8
])



# ============================================================
# FORMÁT ČÍSLA
# ============================================================

def format_cislo(value, desetinna_mista=1):

    text = f"{value:,.{desetinna_mista}f}"

    return (
        text
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", " ")
    )


# ============================================================
# FORMÁT ČÍSLA SE ZNAMÉNKEM (typografické minus)
# ============================================================

def format_cislo_znamenko(value, desetinna_mista=1):

    text = format_cislo(abs(value), desetinna_mista)

    if value < 0:
        return "−" + text

    return text


# ============================================================
# ZKRÁCENÍ NÁZVŮ (legenda donutu na mobilu)
#
# Dlouhé názvy se zkrátí a doplní "…". Názvy musí zůstat
# různé (Plotly by stejné názvy sloučil do jedné výseče),
# proto se případné duplicity očíslují.
# ============================================================

def zkrat_nazvy(nazvy, max_znaku):

    vysledek = []

    for nazev in nazvy:

        if len(nazev) <= max_znaku:
            kratky = nazev
        else:
            kratky = nazev[:max_znaku - 1].rstrip(" ,;–-") + "…"

        zaklad = kratky
        n = 2

        while kratky in vysledek:
            kratky = f"{zaklad} ({n})"
            n += 1

        vysledek.append(kratky)

    return vysledek


# ============================================================
# FORMÁT PROCENTA
# ============================================================

def format_procento(value, desetinna_mista=1):

    return (
        f"{value:.{desetinna_mista}f}"
        .replace(".", ",")
        + " %"
    )


# ============================================================
# PŘEVOD ČÍSLA Z EXCELU
#
# V Excelu s deficitem jsou některé hodnoty uložené jako text
# s českou desetinnou čárkou a mezerou mezi tisíci
# (např. "1 523,23" nebo "-28,52"). Tato funkce je převede
# na float.
# ============================================================

def prevod_cisla(x):

    if pd.isna(x):
        return np.nan

    if isinstance(x, (int, float, np.number)):
        return float(x)

    text = (
        str(x)
        .replace("\xa0", "")
        .replace("\u202f", "")
        .replace(" ", "")
        .replace(",", ".")
        .strip()
    )

    try:
        return float(text)
    except ValueError:
        return np.nan


# ============================================================
# ODSTRANĚNÍ KÓDU ZE ZAČÁTKU NÁZVU
#   '113 Daně ...'                    -> 'Daně ...'
#   '122, 123, 124 Zvláštní daně ...' -> 'Zvláštní daně ...'
# ============================================================

def bez_kodu(text):

    return re.sub(
        r"^\s*\d+(?:\s*(?:,|a|-|–)\s*\d+)*\s*[.:\-–]?\s*",
        "",
        text
    )


# ============================================================
# SLOUČENÍ MALÝCH POLOŽEK DO JEDNÉ VÝSEČE
#
# indexy  – které položky sloučit
# Sloučená výseč se dá na místo poslední z nich.
# ============================================================

def sluc_polozky(values, labels, colors, codes, indexy, nazev, barva, kod):

    soucet = sum(values[i] for i in indexy)
    posledni = indexy[-1]

    nove = []

    for i, polozka in enumerate(zip(values, labels, colors, codes)):

        if i == posledni:
            nove.append((soucet, nazev, barva, kod))
        elif i not in indexy:
            nove.append(polozka)

    values, labels, colors, codes = map(list, zip(*nove))

    return values, labels, colors, codes


# ============================================================
# NAČTENÍ VÝDAJŮ
# ============================================================

def load_vydaje():

    d = pd.read_excel(
        SOUBOR_VYDAJE,
        sheet_name=SHEET_VYDAJE,
        header=None
    )

    codes = d.iloc[1]

    def get_value(code):

        matches = codes[codes == code].index

        if len(matches) == 0:
            raise KeyError(
                f"Kód výdajů {code!r} "
                "nebyl v Excelu nalezen."
            )

        col = matches[0]

        value = d.iloc[8, col]

        if pd.isna(value):
            return 0.0

        return float(value) / 1_000_000_000

    # ========================================================
    # KÓDY, NÁZVY, BARVY
    # ========================================================

    polozky = [
        (1, "Zemědělství, lesní hospodářství a rybářství", MODRA),
        (2, "Průmyslová a ostatní odvětví hospodářství", CERVENA),
        ("31 a 32", "Vzdělávání a školské služby", SEDA),
        (33, "Kultura, církve a sdělovací prostředky", SVETLE_MODRA),
        (34, "Sport a zájmová činnost", SVETLE_MODRA),
        (35, "Zdravotnictví", SVETLE_MODRA),
        (36, "Bydlení, komunální služby a územní rozvoj", SVETLE_MODRA),
        (37, "Ochrana životního prostředí", SVETLE_MODRA),
        (38, "Ostatní výzkum a vývoj", SVETLE_MODRA),
        (39, "Ostatní činnosti související se službami pro fyzické osoby", SVETLE_MODRA),
        (4, "Sociální věci a politika zaměstnanosti", ZELENA),
        (5, "Bezpečnost státu a právní ochrana", TMAVE_SEDA),
        (6, "Všeobecná veřejná správa a služby", SVETLE_SEDA),
    ]

    # ========================================================
    # HODNOTY (bez nul)
    # ========================================================

    values, labels, colors, plot_codes = [], [], [], []

    for kod, nazev, barva in polozky:

        v = get_value(kod)

        if v > 0:
            values.append(v)
            labels.append(nazev)
            colors.append(barva)
            plot_codes.append(kod)

    # ========================================================
    # SLOUČENÍ MALÝCH POLOŽEK SKUPINY 3 (33–39)
    # ========================================================

    skupina3 = (33, 34, 35, 36, 37, 38, 39)

    celkem = sum(values)

    male = [
        i for i, k in enumerate(plot_codes)
        if k in skupina3 and values[i] / celkem * 100 < VYDAJE_SLOUCIT_POD
    ]

    if VYDAJE_SLOUCIT_POD > 0 and len(male) > 1:

        print(f"  Výdaje – sloučeno do „{VYDAJE_NAZEV_SLOUCENE}“:")
        for i in male:
            print(f"    {format_cislo(values[i], 3):>8} mld. Kč  {labels[i]}")

        values, labels, colors, plot_codes = sluc_polozky(
            values, labels, colors, plot_codes, male,
            VYDAJE_NAZEV_SLOUCENE, ZLATA, "33–39 ostatní"
        )

    # Zbylé položky skupiny 3 dostanou různé odstíny modré
    odstiny = iter(ODSTINY_SKUPINA3 * 3)

    for i, k in enumerate(plot_codes):
        if k in skupina3:
            colors[i] = next(odstiny)

    # ========================================================
    # CELKEM
    # ========================================================

    values = np.array(values)

    total = values.sum()

    if total > 0:
        shares = values / total * 100
    else:
        shares = np.zeros_like(values)

    return {
        "values": values,
        "labels": labels,
        "codes": plot_codes,
        "shares": shares,
        "colors": colors,
        "total": total
    }


# ============================================================
# NAČTENÍ PŘÍJMŮ
# ============================================================

def load_prijmy():

    df = pd.read_excel(
        SOUBOR_PRIJMY,
        sheet_name=SHEET_PRIJMY,
        header=None
    )

    # ========================================================
    # NALEZENÍ ŘÁDKU PODLE KÓDU (od řádku 'od' dál)
    # ========================================================

    def najdi_radek(kod, od=0):

        kod = str(kod).strip()

        vzor = re.compile(
            r"^\s*"
            + re.escape(kod)
            + r"(?=\s|$)"
        )

        for i in range(od, len(df)):

            bunka = str(df.iloc[i, 1])

            if vzor.search(bunka):
                return i

        return None

    # ========================================================
    # HODNOTA 2027 (sloupec H)
    # ========================================================

    def hodnota_radku(radek):

        hodnota = df.iloc[radek, 7]

        if pd.isna(hodnota):
            return 0.0

        return float(hodnota) / 1_000_000_000

    def hodnota_2027(kod):

        radek = najdi_radek(kod)

        if radek is None:
            print(
                f"VAROVÁNÍ – kód {kod} "
                "nebyl nalezen."
            )
            return 0.0

        return hodnota_radku(radek)

    # ========================================================
    # NÁZEV ZE SLOUPCE B (bez čísla kódu)
    # ========================================================

    def nazev_radku(radek, zalozni):

        if radek is None:
            return zalozni

        text = df.iloc[radek, 1]

        if pd.isna(text):
            return zalozni

        return bez_kodu(" ".join(str(text).split()))

    def nazev_z_excelu(kod, zalozni):

        return nazev_radku(najdi_radek(kod), zalozni)

    # ========================================================
    # HLAVNÍ SKUPINY
    #   (kód, záložní název, barva, ruční název nebo None)
    #   None -> název ze sloupce B (bez kódu)
    # ========================================================

    hlavni = [
        ("111", "DPFO", MODRA, "DPFO"),
        ("112", "DPPO", CERVENA, "DPPO"),
        ("113", "Ostatní přímé daně", TMAVE_SEDA, None),
        ("1211", "DPH", SVETLE_MODRA, "DPH"),
        ("122, 123, 124", "Zvláštní daně, poplatky a obdobná peněžitá plnění", ZELENA,
         "Zvláštní daně, poplatky a obdobná peněžitá plnění"),   # zkrácený název
        ("13", "Daně a poplatky z vybraných činností a služeb", SEDA, None),
        ("14", "Daně a cla ze zahraničí", ZLATA, None),
        ("15", "Majetkové daně", SVETLE_SEDA, None),
        ("17", "Ostatní daňové příjmy", TMAVE_SEDA, None),
        ("16", "Příjem z povinného pojistného", TMAVE_MODRA, None),
    ]

    # ========================================================
    # DETAIL SKUPINY 13
    #
    # Podpoložky 132–138 se berou z bloku POD řádkem 13
    # (v Excelu řádky 29–35) – jen ten dává přesně součet
    # položky 13 (v bloku nad ním má 137 jinou hodnotu).
    # ========================================================

    detail13 = [
        ("132", "Provoz motorových vozidel"),
        ("133", "Životní prostředí"),
        ("134", "Místní poplatky"),
        ("135", "Ostatní odvody"),
        ("136", "Správní a soudní poplatky"),
        ("137", "Poplatky na činnost správních úřadů"),
        ("138", "Hazardní hry"),
    ]

    radek13 = najdi_radek("13")

    detail13_values = []
    detail13_labels_final = []
    detail13_kody_final = []

    for kod, zalozni in detail13:

        radek = (
            najdi_radek(kod, od=radek13 + 1)
            if radek13 is not None else None
        )

        if radek is None:
            print(f"VAROVÁNÍ – kód {kod} pod řádkem 13 nebyl nalezen.")
            continue

        v = hodnota_radku(radek)

        if v > 0:
            detail13_values.append(v)
            detail13_labels_final.append(nazev_radku(radek, zalozni))
            detail13_kody_final.append(kod)

    detail13_values = np.array(detail13_values)

    skupina_13 = hodnota_2027("13")

    skupina_16 = hodnota_2027("16")

    detail13_total = detail13_values.sum()

    rozdil13 = skupina_13 - detail13_total

    if skupina_13 > 0:
        detail13_shares = detail13_values / skupina_13 * 100
    else:
        detail13_shares = np.zeros(len(detail13_values))

    # ========================================================
    # POLOŽKY DONUTU
    # ========================================================

    values, labels, colors, codes = [], [], [], []

    for kod, zalozni, barva, rucni in hlavni:

        v = hodnota_2027(kod)

        if v <= 0:
            continue

        if kod == "13" and ROZPAD_13:

            for d_v, d_l, d_k in zip(
                detail13_values,
                detail13_labels_final,
                detail13_kody_final
            ):
                values.append(d_v)
                labels.append(d_l)
                colors.append(SEDA)
                codes.append(d_k)

            continue

        values.append(v)
        labels.append(rucni or nazev_z_excelu(kod, zalozni))
        colors.append(barva)
        codes.append(kod)

    # ========================================================
    # NOVÁ POLOŽKA Z ŘÁDKU 46 (číslo řádku v Excelu)
    # ========================================================

    radek = RADEK_NOVA_POLOZKA - 1      # Excel čísluje od 1, pandas od 0

    if radek >= len(df):

        print(f"VAROVÁNÍ – řádek {RADEK_NOVA_POLOZKA} v listu neexistuje.")

    else:

        v = hodnota_radku(radek)

        popisek = nazev_radku(
            radek, f"Nová položka (řádek {RADEK_NOVA_POLOZKA})"
        )

        # Celé velkými písmeny -> jen první písmeno velké
        if popisek.isupper():
            popisek = popisek.capitalize()

        # „… celkem“ na konci je v legendě zbytečné
        popisek = re.sub(r"\s+celkem$", "", popisek)

        if v > 0:
            values.append(v)
            labels.append(popisek)
            colors.append(SVETLE_ZLATA)
            codes.append(f"řádek {RADEK_NOVA_POLOZKA}")
        else:
            print(
                f"VAROVÁNÍ – řádek {RADEK_NOVA_POLOZKA} "
                f"({popisek}) nemá kladnou hodnotu ve sloupci H."
            )

    # ========================================================
    # SLOUČENÍ MALÝCH POLOŽEK
    # ========================================================

    if SLOUCIT_POD > 0 and values:

        celkem = sum(values)

        male = [
            i for i, v in enumerate(values)
            if v / celkem * 100 < SLOUCIT_POD
        ]

        if len(male) > 1:

            print(f"  Příjmy – sloučeno do „{NAZEV_SLOUCENE}“:")
            for i in male:
                print(f"    {format_cislo(values[i], 3):>8} mld. Kč  {labels[i]}")

            values, labels, colors, codes = sluc_polozky(
                values, labels, colors, codes, male,
                NAZEV_SLOUCENE, ZLATA, "ostatní"
            )

    # ========================================================
    # CELKEM
    # ========================================================

    values = np.array(values)

    total = values.sum()

    if total > 0:
        shares = values / total * 100
    else:
        shares = np.zeros_like(values)

    return {
        "values": values,
        "labels": labels,
        "codes": codes,
        "shares": shares,
        "colors": colors,
        "total": total,
        "skupina_13": skupina_13,
        "skupina_16": skupina_16,
        "detail13_values": detail13_values,
        "detail13_labels": detail13_labels_final,
        "detail13_codes": detail13_kody_final,
        "detail13_shares": detail13_shares,
        "detail13_total": detail13_total,
        "rozdil13": rozdil13
    }


# ============================================================
# NAČTENÍ DEFICITU (SALDA)
#
# Z tabulky se berou jen roční hodnoty:
#   - řádky s měsícem 12 = skutečnost za celý rok,
#   - řádky bez měsíce (2026, 2027) = výhled.
# Hodnoty jsou už v mld. Kč.
# ============================================================

def load_deficit():

    df = pd.read_excel(
        SOUBOR_DEFICIT,
        sheet_name=SHEET_DEFICIT
    )

    df.columns = [
        str(c).strip().lower()
        for c in df.columns
    ]

    for sloupec in ["rok", "měsíc", "saldo"]:
        if sloupec not in df.columns:
            raise KeyError(
                f"Sloupec {sloupec!r} nebyl "
                "v souboru s deficitem nalezen."
            )

    df["rok"] = df["rok"].map(prevod_cisla)
    df["měsíc"] = df["měsíc"].map(prevod_cisla)
    df["saldo"] = df["saldo"].map(prevod_cisla)

    # ========================================================
    # JEN ROČNÍ HODNOTY
    # ========================================================

    rocni = df[
        df["měsíc"].isna() | (df["měsíc"] == 12)
    ].dropna(subset=["rok", "saldo"]).copy()

    rocni["odhad"] = rocni["měsíc"].isna()

    # Kdyby byl některý rok v tabulce dvakrát,
    # přednost má skutečnost (měsíc 12)
    rocni = (
        rocni
        .sort_values(["rok", "odhad"])
        .drop_duplicates(subset="rok", keep="first")
        .sort_values("rok")
    )

    # ========================================================
    # JEN POSLEDNÍCH N LET
    # ========================================================

    rocni = rocni.tail(DEFICIT_POCET_LET)

    roky = rocni["rok"].astype(int).to_numpy()
    saldo = rocni["saldo"].to_numpy()
    odhad = rocni["odhad"].to_numpy()

    # ========================================================
    # BARVY
    #   deficit = červená, přebytek = modrá,
    #   výhled = světlejší odstín
    # ========================================================

    colors = []

    for hodnota, je_odhad in zip(saldo, odhad):

        if hodnota < 0:
            colors.append(SVETLE_CERVENA if je_odhad else CERVENA)
        else:
            colors.append(SVETLE_MODRA_ODHAD if je_odhad else MODRA)

    return {
        "roky": roky,
        "saldo": saldo,
        "odhad": odhad,
        "colors": colors
    }




# ============================================================
# NAČTENÍ PODÍLU OBNOVITELNÝCH ZDROJŮ
# ============================================================

def load_oze():

    df = pd.read_excel(
        SOUBOR_OZE,
        sheet_name=SHEET_OZE,
        header=None
    )

    prvni = df[0].astype(str).str.strip()

    radek_cas = prvni[prvni == "TIME"].index[0]
    radek_jednotka = prvni[prvni.str.startswith("UNIT")].index[0]

    # ========================================================
    # SLOUPCE S GWh PODLE ROKU
    # ========================================================

    sloupce_podle_roku = {}

    for c in range(2, df.shape[1]):

        cas = str(df.iloc[radek_cas, c]).strip()
        jednotka = str(df.iloc[radek_jednotka, c]).strip()

        if jednotka == "Gigawatt-hour" and re.match(r"^\d{4}-\d{2}$", cas):
            rok = int(cas[:4])
            sloupce_podle_roku.setdefault(rok, []).append(c)

    # ========================================================
    # ŘADA HODNOT PRO ZEMI A UKAZATEL
    # ========================================================

    def rada(zeme, ukazatel, sloupce):

        radky = df[
            (prvni == zeme)
            & (df[1].astype(str).str.strip() == ukazatel)
        ]

        if radky.empty:
            return None

        return np.array([
            prevod_cisla(v)
            for v in radky.iloc[0, sloupce]
        ])

    def podil(zeme, sloupce):

        celkem = rada(zeme, "Total", sloupce)
        oze = rada(zeme, "Renewables and biofuels", sloupce)

        if celkem is None or oze is None or len(sloupce) == 0:
            return None

        if np.isnan(celkem).any() or np.isnan(oze).any():
            return None

        if celkem.sum() <= 0:
            return None

        return oze.sum() / celkem.sum() * 100

    def mesic_ma_data(c):

        for zeme in [EU_AGREGAT] + list(ZEME_EU):
            for ukazatel in ["Total", "Renewables and biofuels"]:
                hodnota = rada(zeme, ukazatel, [c])
                if hodnota is None or np.isnan(hodnota).any():
                    return False
        return True

    # ========================================================
    # ROK
    #   OZE_ROK = None -> poslední úplný rok (12 měsíců u EU-27)
    #   OZE_ROK = 2026 -> i neúplný rok
    # ========================================================

    if OZE_ROK is not None:
        rok = OZE_ROK
    else:
        rok = None
        for r in sorted(sloupce_podle_roku, reverse=True):
            if (
                len(sloupce_podle_roku[r]) == 12
                and podil(EU_AGREGAT, sloupce_podle_roku[r]) is not None
            ):
                rok = r
                break

    if rok is None or rok not in sloupce_podle_roku:
        raise ValueError(
            "V souboru s OZE nebyl nalezen požadovaný rok s daty."
        )

    # ========================================================
    # SPOLEČNÉ MĚSÍCE
    #
    # Aby byly země srovnatelné, berou se jen měsíce, za které
    # mají data všechny země EU i EU-27 (u neúplného roku
    # např. jen leden–červen).
    # ========================================================

    def cislo_mesice(c):
        return int(str(df.iloc[radek_cas, c]).strip()[5:7])

    sloupce = [
        c for c in sloupce_podle_roku[rok]
        if (OZE_MESICE is None or cislo_mesice(c) in OZE_MESICE)
        and mesic_ma_data(c)
    ]

    if OZE_MESICE is not None and len(sloupce) < len(OZE_MESICE):
        print(
            "VAROVÁNÍ – OZE: některé z vybraných měsíců nemají "
            "všechny země, použijí se jen dostupné."
        )

    if not sloupce:
        raise ValueError(
            f"Za rok {rok} nemají všechny země EU data za žádný měsíc."
        )

    MESICE = [
        "leden", "únor", "březen", "duben", "květen", "červen",
        "červenec", "srpen", "září", "říjen", "listopad", "prosinec"
    ]

    cisla_mesicu = [
        int(str(df.iloc[radek_cas, c]).strip()[5:7])
        for c in sloupce
    ]

    CTVRTLETI = {
        (1, 2, 3): 1, (4, 5, 6): 2, (7, 8, 9): 3, (10, 11, 12): 4
    }

    if len(sloupce) == 12:
        obdobi = f"rok {rok}"
    elif tuple(cisla_mesicu) in CTVRTLETI:
        obdobi = f"{CTVRTLETI[tuple(cisla_mesicu)]}. čtvrtletí {rok}"
    elif len(sloupce) == 1:
        obdobi = f"{MESICE[cisla_mesicu[0] - 1]} {rok}"
    else:
        obdobi = (
            f"{MESICE[cisla_mesicu[0] - 1]}–"
            f"{MESICE[cisla_mesicu[-1] - 1]} {rok}"
        )

    # ========================================================
    # PODÍLY ZEMÍ EU
    # ========================================================

    zeme = []
    hodnoty = []

    for anglicky, cesky in ZEME_EU.items():

        h = podil(anglicky, sloupce)

        if h is None:
            print(
                f"VAROVÁNÍ – OZE: {cesky} nemá za {obdobi} "
                "úplná data, v grafu chybí."
            )
            continue

        zeme.append(cesky)
        hodnoty.append(h)

    # Seřazení vzestupně -> ve vodorovném grafu je největší nahoře
    poradi = np.argsort(hodnoty)

    zeme = [zeme[i] for i in poradi]
    hodnoty = np.array(hodnoty)[poradi]

    return {
        "rok": rok,
        "obdobi": obdobi,
        "zeme": zeme,
        "hodnoty": hodnoty,
        "eu": podil(EU_AGREGAT, sloupce)
    }


# ============================================================
# NAČTENÍ VÝDAJŮ NA OBRANU (% HDP)
#
# Stejně jako ve skriptu pro Canvu: list "Vojenské výdaje",
# sloupec "Země" a sloupec s rokem 2025.
# ============================================================

def load_obrana():

    df = pd.read_excel(
        SOUBOR_OBRANA,
        sheet_name=SHEET_OBRANA
    )

    # Sloupec s rokem může být v Excelu číslo i text ("2025")
    sloupec_roku = None

    for c in df.columns:
        if str(c).strip().split(".")[0] == str(OBRANA_VYCHOZI_ROK):
            sloupec_roku = c
            break

    if "Země" not in df.columns or sloupec_roku is None:
        raise KeyError(
            f"V souboru s obranou chybí sloupec 'Země' "
            f"nebo {OBRANA_VYCHOZI_ROK}."
        )

    df["Země"] = df["Země"].astype(str).str.strip()
    df["hodnota"] = df[sloupec_roku].map(prevod_cisla)

    def hodnota_zeme(nazvy):

        radky = df[df["Země"].isin(nazvy)].dropna(subset=["hodnota"])

        if radky.empty:
            raise KeyError(
                f"V souboru s obranou chybí země {nazvy[0]!r}."
            )

        return float(radky["hodnota"].iloc[0])

    cesko = hodnota_zeme([CESKO])
    usa = hodnota_zeme(["USA", "Spojené státy"])

    # ========================================================
    # ROVNOMĚRNÁ CESTA K CÍLI
    # ========================================================

    pocet_let = OBRANA_CIL_ROK - OBRANA_VYCHOZI_ROK

    roky = list(range(OBRANA_VYCHOZI_ROK, OBRANA_CIL_ROK + 1))

    krok_cr = (OBRANA_CIL - cesko) / pocet_let
    krok_usa = (OBRANA_CIL - usa) / pocet_let

    cesta_cr = [cesko + krok_cr * (r - OBRANA_VYCHOZI_ROK) for r in roky]
    cesta_usa = [usa + krok_usa * (r - OBRANA_VYCHOZI_ROK) for r in roky]

    cesko_zvyrazneny = cesko + krok_cr * (
        OBRANA_ZVYRAZNENY_ROK - OBRANA_VYCHOZI_ROK
    )

    return {
        "cesko": cesko,
        "usa": usa,
        "rozdil": usa - cesko,
        "roky": roky,
        "cesta_cr": cesta_cr,
        "cesta_usa": cesta_usa,
        "cesko_zvyrazneny": cesko_zvyrazneny
    }


# ============================================================
# CSS
# ============================================================

def html_css():

    return r"""

<style>

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    background: #f4f6f7;
    color: #222222;
    font-family: "Segoe UI", Arial, sans-serif;
}

.dashboard {
    max-width: 1800px;
    margin: 0 auto;
    padding: 28px 34px 35px 34px;
}

.header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 22px;
}

.title {
    font-size: 30px;
    font-weight: 700;
    letter-spacing: -0.5px;
}

.subtitle {
    margin-top: 7px;
    font-size: 14px;
    color: #777777;
}

.year {
    color: #0055A0;
    font-size: 14px;
    font-weight: 700;
    padding-top: 5px;
}

/* Druhý řádek: 5 KPI karet */
.cards {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 14px;
    margin-bottom: 18px;
}

.card {
    background: white;
    border: 1px solid #d9dddf;
    border-radius: 10px;
    padding: 17px 19px 15px 19px;
    min-height: 105px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.035);
}

.card-title {
    font-size: 11px;
    font-weight: 700;
    color: #777777;
}

.card-value {
    font-size: 25px;
    font-weight: 700;
    margin-top: 9px;
    color: #222222;
}

.card-subtitle {
    margin-top: 5px;
    font-size: 10px;
    color: #777777;
}


/* ============================================================
   KARTY UKAZATELŮ SE ŠIPKAMI
   ============================================================ */

.cards-ukazatele {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
    margin-bottom: 14px;
}

.card-value-row {
    display: flex;
    align-items: baseline;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 9px;
}

.card-value-row .card-value {
    margin-top: 0;
}

.card-arrow {
    font-size: 20px;
    font-weight: 700;
}

.card-badge {
    font-size: 11px;
    font-weight: 700;
}

/* Karta ukazatele s ikonou */
.card-ukazatel {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 14px;
    border-top: 4px solid #0055A0;
}

.card-ukazatel-text {
    min-width: 0;
    flex: 1;
}

.card-ikona {
    flex: 0 0 auto;
    width: 52px;
    height: 52px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
}

.card-ikona svg {
    width: 26px;
    height: 26px;
    fill: none;
    stroke-width: 2;
    stroke-linecap: round;
    stroke-linejoin: round;
}

/* Prázdná karta – místo pro doplnění */
.card-volna {
    border: 2px dashed #d0d4d6;
    box-shadow: none;
    background: #fafbfb;
    padding: 16px 18px 14px 18px;
}

.card-volna .card-value {
    color: #aaaaaa;
    font-weight: 400;
    font-size: 18px;
}


/* ============================================================
   KARTA S PODÍLEM OBSLUHY DLUHU
   ============================================================ */

/* Silnější barevný okraj vlevo (5 px místo 1 px) by posunul
   text doprava – vnitřní okraj se proto o 4 px zmenší,
   aby texty začínaly stejně jako u ostatních karet */
.card-nej {
    border-left: 5px solid #0055A0;
    padding-left: 15px;
}

/* Ukazatel podílu (vodorovný proužek) */
.podil-pruh {
    margin-top: 12px;
    height: 8px;
    background: #eef0f1;
    border-radius: 4px;
    overflow: hidden;
}

.podil-pruh-vypln {
    height: 100%;
    border-radius: 4px;
}


/* ============================================================
   HLAVNÍ GRID
   ============================================================ */

.main-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 18px;
    align-items: start;
}

/* Graf obrany pod hlavní částí */
.obrana-panel {
    margin-top: 18px;
}

/* Sloupec s grafy pod sebou */
.column {
    display: flex;
    flex-direction: column;
    gap: 18px;
    min-width: 0;
}

.panel {
    background: white;
    border: 1px solid #d9dddf;
    border-radius: 10px;
    padding: 15px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.035);
    min-width: 0;
}

.panel-title {
    font-size: 16px;
    font-weight: 700;
    margin-bottom: 2px;
}

.panel-subtitle {
    font-size: 10px;
    color: #777777;
    margin-bottom: 5px;
}


/* ============================================================
   DETAIL SKUPINY 13
   ============================================================ */

.detail-row {
    display: grid;
    grid-template-columns: 8px minmax(0, 1fr) 82px 62px;
    gap: 7px;
    align-items: center;
    margin: 11px 0;
}

.square {
    width: 8px;
    height: 8px;
    background: #2E808C;
    border-radius: 1px;
}

.detail-name {
    font-size: 12px;
    font-weight: 400;
    line-height: 1.15;
}

.detail-value {
    text-align: right;
    font-size: 11px;
    font-weight: 700;
    white-space: nowrap;
}

.detail-share {
    text-align: right;
    font-size: 10px;
    color: #777777;
    white-space: nowrap;
}

.control-row {
    display: flex;
    justify-content: space-between;
    padding: 9px 0;
    border-bottom: 1px solid #eeeeee;
    font-size: 11px;
}

.control-label {
    color: #777777;
}

.control-value {
    font-weight: 700;
}


/* ============================================================
   GRAF DEFICITU
   ============================================================ */

.deficit-panel {
    margin-top: 18px;
}

.deficit-legend {
    display: flex;
    flex-wrap: wrap;
    gap: 18px;
    font-size: 11px;
    color: #555555;
    margin: 4px 0 2px 0;
}

.deficit-legend span {
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

.legend-box {
    width: 11px;
    height: 11px;
    border-radius: 2px;
    display: inline-block;
}


.footer {
    border-top: 1px solid #d0d0d0;
    margin-top: 22px;
    padding-top: 12px;
    display: flex;
    justify-content: space-between;
    color: #888888;
    font-size: 9px;
}


/* ============================================================
   RESPONSIVITA
   ============================================================ */

@media (max-width: 1100px) {

    .cards {
        grid-template-columns: repeat(3, 1fr);
    }

    .main-grid {
        grid-template-columns: 1fr;
    }

}

@media (max-width: 650px) {

    .dashboard {
        padding: 12px;
    }

    .title {
        font-size: 24px;
    }

    .panel {
        padding: 8px;
    }

    .card-value {
        font-size: 22px;
    }

    /* Lišta nástrojů Plotly na mobilu jen překáží */
    .modebar-container {
        display: none !important;
    }

    .footer {
        display: block;
    }

    .cards {
        grid-template-columns: 1fr;
    }

    .cards-ukazatele {
        grid-template-columns: 1fr;
    }

    .header {
        display: block;
    }

    .year {
        margin-top: 10px;
    }

}

</style>

"""


# ============================================================
# DONUT
#
# Oba donuty mají stejnou velikost, protože:
#   - plocha pro donut má pevnou výšku DONUT_PLOCHA,
#   - legenda je ve spodním okraji, který má u obou grafů
#     stejnou výšku (podle delší legendy),
#   - automatické rozšiřování okrajů je vypnuté,
#     takže legenda donut nezmenší.
#
# V legendě je za názvem podíl, např. „DPH (20,4 %)“ – tak jsou
# vidět i podíly malých výsečí, do kterých se číslo nevejde.
# ============================================================

def create_donut(
    values,
    labels,
    colors,
    shares,
    total,
    title,
    pocet_radku_legendy
):

    # Obyčejné seznamy místo numpy polí (spolehlivě se přenesou do HTML)
    values = [float(v) for v in values]
    shares = [float(v) for v in shares]
    colors = list(colors)
    labels = list(labels)

    # Legenda: název + podíl (nezlomitelná mezera před %)
    def s_podilem(nazvy):
        return [
            f"{n} ({format_procento(p)})".replace(" %", "\u00a0%")
            for n, p in zip(nazvy, shares)
        ]

    labels_desktop = s_podilem(labels)

    # Mobil: zkrátí se jen název, podíl zůstane celý
    labels_mobil = s_podilem(zkrat_nazvy(labels, MOBIL_LEGENDA_ZNAKU))

    # Text bubliny po najetí myší – hotový řetězec pro každou výseč
    # (plný název je přímo v textu bubliny, aby zůstal celý
    #  i na mobilu, kde je název v legendě zkrácený)
    hover_texty = [
        f"<b>{l}</b><br>"
        f"Částka: {format_cislo(v, 2)} mld. Kč<br>"
        f"Podíl: {format_procento(p)}"
        for l, v, p in zip(labels, values, shares)
    ]

    # Procenta do výsečí – u malých výsečí se nic nepíše
    # (podíl je v legendě)
    text_vyseci = [
        format_procento(p) if p >= DONUT_MIN_PODIL_TEXT else ""
        for p in shares
    ]

    # Spodní okraj pro legendu – stejný u obou grafů
    spodni_okraj = pocet_radku_legendy * RADEK_LEGENDY + 30

    vyska = TITULEK_VYSKA + DONUT_PLOCHA + spodni_okraj

    fig = go.Figure()

    fig.add_trace(

        go.Pie(

            values=values,
            labels=labels_desktop,

            hole=0.60,
            sort=False,
            direction="clockwise",
            rotation=90,

            # Donut vyplní celou plochu grafu
            domain=dict(x=[0, 1], y=[0, 1]),

            marker=dict(
                colors=colors,
                line=dict(color="white", width=2)
            ),

            # Procenta uvnitř výsečí
            text=text_vyseci,
            textinfo="text",
            textposition="inside",
            insidetextorientation="horizontal",
            textfont=dict(
                size=14,
                color="white",
                family="Arial Black"
            ),

            hovertext=hover_texty,

            hovertemplate=(
                "%{hovertext}"
                "<extra></extra>"
            )

        )

    )

    fig.update_layout(

        title=dict(
            text=title,
            x=0,
            xanchor="left",
            **TITULEK_POZICE,
            font=dict(size=17, color=TEXT)
        ),

        # Legenda pod donutem (ve spodním okraji)
        showlegend=True,
        legend=dict(
            orientation="v",
            x=0.04,
            y=-0.04,
            xanchor="left",
            yanchor="top",
            font=dict(size=14, color=TEXT, family="Segoe UI"),
            itemsizing="constant"
        ),

        # Pevné okraje – legenda je nesmí zvětšovat
        margin=dict(
            l=0,
            r=0,
            t=TITULEK_VYSKA,
            b=spodni_okraj,
            autoexpand=False
        ),

        height=vyska,
        autosize=True,

        # Česká desetinná čárka, mezera jako oddělovač tisíců
        separators=", ",

        paper_bgcolor="white",
        plot_bgcolor="white",

        uniformtext=dict(minsize=8, mode="hide")

    )

    # ========================================================
    # TEXT UPROSTŘED DONUTU
    # ========================================================

    radky_stredu = [
        ("Celkem", 15, False, STRED_ODSTUP),
        (format_cislo(total, 1), 26, True, 0),
        ("mld. Kč", 15, False, -STRED_ODSTUP),
    ]

    for text, velikost, tucne, posun in radky_stredu:

        fig.add_annotation(
            text=f"<b>{text}</b>" if tucne else text,
            xref="paper",
            yref="paper",
            x=0.5,
            y=0.5,
            xanchor="center",
            yanchor="middle",
            yshift=posun,
            showarrow=False,
            font=dict(size=velikost, color=TEXT, family="Segoe UI")
        )

    # ========================================================
    # PŘEPNUTÍ DESKTOP / MOBIL
    # ========================================================

    prepinani = {
        "desktop": {
            "data": {"labels": [labels_desktop], "textfont.size": 14},
            "layout": {"legend.font.size": 14, "title.font.size": 17}
        },
        "mobil": {
            "data": {
                "labels": [labels_mobil],
                "textfont.size": 11
            },
            "layout": {"legend.font.size": 11, "title.font.size": 15}
        }
    }

    return fig, prepinani


# ============================================================
# GRAF VÝVOJE DEFICITU (SLOUPCOVÝ – JEN SALDO)
# ============================================================

def create_deficit_chart(deficit):

    roky_text = [str(r) for r in deficit["roky"]]

    popisky = [
        format_cislo(v, 1)
        for v in deficit["saldo"]
    ]

    typ = [
        "výhled" if o else "skutečnost"
        for o in deficit["odhad"]
    ]

    fig = go.Figure()

    fig.add_trace(

        go.Bar(

            x=roky_text,
            y=[float(v) for v in deficit["saldo"]],

            marker=dict(
                color=list(deficit["colors"]),
                line=dict(width=0)
            ),

            text=popisky,
            textposition="outside",
            cliponaxis=False,
            textfont=dict(size=11, color=TEXT, family="Segoe UI"),

            hovertext=[
                f"Saldo: {p} mld. Kč<br>{t}"
                for p, t in zip(popisky, typ)
            ],

            hovertemplate=(
                "<b>%{x}</b><br>"
                "%{hovertext}"
                "<extra></extra>"
            )

        )

    )

    # Rezerva na osách, aby se popisky vešly nad / pod sloupce
    minimum = min(0, float(np.min(deficit["saldo"])))
    maximum = max(0, float(np.max(deficit["saldo"])))
    rozpeti = maximum - minimum if maximum > minimum else 1
    rezerva = rozpeti * 0.10

    fig.update_layout(

        title=dict(
            text="Vývoj salda státního rozpočtu",
            x=0,
            xanchor="left",
            **TITULEK_POZICE,
            font=dict(size=17, color=TEXT)
        ),

        showlegend=False,

        height=DEFICIT_VYSKA,
        autosize=True,

        margin=dict(l=60, r=20, t=TITULEK_VYSKA, b=45),

        separators=", ",

        paper_bgcolor="white",
        plot_bgcolor="white",

        bargap=0.25,

        xaxis=dict(
            type="category",
            tickfont=dict(size=11, color=TEXT),
            showgrid=False,
            linecolor=LINE
        ),

        yaxis=dict(
            title=dict(text="mld. Kč", font=dict(size=11, color=TEXT_LIGHT)),
            tickfont=dict(size=11, color=TEXT_LIGHT),
            gridcolor="#EEEEEE",
            zeroline=True,
            zerolinecolor=TMAVE_SEDA,
            zerolinewidth=1.5,
            range=[minimum - rezerva, maximum + rezerva],
            tickformat=",.0f"
        )

    )

    # ========================================================
    # PŘEPNUTÍ DESKTOP / MOBIL
    #   mobil: celá čísla, svislé popisky, menší písmo,
    #          roky natočené, větší rezerva na ose
    # ========================================================

    rezerva_mobil = rozpeti * 0.20

    prepinani = {
        "desktop": {
            "data": {
                "text": [popisky],
                "textangle": "auto",
                "textfont.size": 11
            },
            "layout": {
                "height": DEFICIT_VYSKA,
                "title.font.size": 17,
                "margin.l": 60,
                "margin.r": 20,
                "margin.b": 45,
                "xaxis.tickangle": "auto",
                "xaxis.tickfont.size": 11,
                "yaxis.tickfont.size": 11,
                "yaxis.range": [minimum - rezerva, maximum + rezerva]
            }
        },
        "mobil": {
            "data": {
                "text": [[format_cislo(v, 0) for v in deficit["saldo"]]],
                "textangle": -90,
                "textfont.size": 9
            },
            "layout": {
                "height": MOBIL_VYSKA_DEFICIT,
                "title.font.size": 15,
                "margin.l": 45,
                "margin.r": 5,
                "margin.b": 50,
                "xaxis.tickangle": -90,
                "xaxis.tickfont.size": 9,
                "yaxis.tickfont.size": 9,
                "yaxis.range": [
                    minimum - rezerva_mobil,
                    maximum + rezerva_mobil
                ]
            }
        }
    }

    return fig, prepinani


# ============================================================
# GRAF 3 – PODÍL OBNOVITELNÝCH ZDROJŮ V ZEMÍCH EU
# (Česko červeně, ostatní modře, svislá čára = průměr EU-27)
# ============================================================

def create_oze_chart(oze, vyska):

    barvy = [
        CERVENA if z == CESKO else MODRA
        for z in oze["zeme"]
    ]

    popisky_os = [
        f"<b>{z}</b>" if z == CESKO else z
        for z in oze["zeme"]
    ]

    popisky = [
        format_procento(v)
        for v in oze["hodnoty"]
    ]

    # Nadpis je vždy jen jednořádkový (víceřádkový nadpis Plotly
    # usadí jinak než ostatní). Podtitulek se zdrojem je pod
    # grafem v HTML – viz create_dashboard_html.
    nadpis = "Podíl obnovitelných zdrojů na výrobě elektřiny"

    nadpis_mobil = "Podíl OZE na výrobě elektřiny"

    fig = go.Figure()

    fig.add_trace(

        go.Bar(

            x=[float(v) for v in oze["hodnoty"]],
            y=popisky_os,
            orientation="h",

            marker=dict(color=barvy, line=dict(width=0)),

            text=popisky,
            textposition="outside",
            cliponaxis=False,
            textfont=dict(size=11, color=TEXT, family="Segoe UI"),

            hovertext=[
                f"<b>{z}</b><br>Podíl OZE: {p}"
                for z, p in zip(oze["zeme"], popisky)
            ],

            hovertemplate=(
                "%{hovertext}"
                "<extra></extra>"
            )

        )

    )

    # Průměr EU-27
    if oze["eu"] is not None:

        fig.add_vline(
            x=float(oze["eu"]),
            line=dict(color=TMAVE_SEDA, width=1.5, dash="dash")
        )

        fig.add_annotation(
            x=oze["eu"],
            y=1,
            xref="x",
            yref="paper",
            yanchor="bottom",
            text=f"EU-27: {format_procento(oze['eu'])}",
            showarrow=False,
            font=dict(size=11, color=TMAVE_SEDA, family="Segoe UI")
        )

    fig.update_layout(

        title=dict(
            text=nadpis,
            x=0,
            xanchor="left",
            **TITULEK_POZICE,
            font=dict(size=17, color=TEXT)
        ),

        showlegend=False,

        height=vyska,
        autosize=True,

        margin=dict(l=10, r=55, t=TITULEK_VYSKA + 20, b=35),

        separators=", ",

        paper_bgcolor="white",
        plot_bgcolor="white",

        bargap=0.25,

        xaxis=dict(
            range=[0, 105],
            ticksuffix=" %",
            tickfont=dict(size=11, color=TEXT_LIGHT),
            gridcolor="#EEEEEE",
            zeroline=False
        ),

        yaxis=dict(
            type="category",
            tickfont=dict(size=12, color=TEXT),
            automargin=True
        )

    )

    # ========================================================
    # PŘEPNUTÍ DESKTOP / MOBIL
    #   mobil: kratší nadpis, menší písmo
    # ========================================================

    prepinani = {
        "desktop": {
            "data": {"textfont.size": 11},
            "layout": {
                "title.text": nadpis,
                "title.font.size": 17,
                "margin.t": TITULEK_VYSKA + 20,
                "margin.r": 55,
                "xaxis.tickfont.size": 11,
                "yaxis.tickfont.size": 12
            }
        },
        "mobil": {
            "data": {"textfont.size": 9},
            "layout": {
                "title.text": nadpis_mobil,
                "title.font.size": 15,
                "margin.t": TITULEK_VYSKA + 20,
                "margin.r": 40,
                "xaxis.tickfont.size": 9,
                "yaxis.tickfont.size": 10
            }
        }
    }

    return fig, prepinani


# ============================================================
# GRAF 4 – BĚŽNÉ A KAPITÁLOVÉ VÝDAJE (SKLÁDANÝ SLOUPCOVÝ)
# ============================================================

def create_bezne_kapitalove_chart(vyska):

    if not (len(YEARS) == len(BEZNE) == len(KAPITALOVE)):
        raise ValueError(
            "YEARS, BEZNE a KAPITALOVE musí mít stejný počet hodnot."
        )

    # Převod na obyčejné seznamy (spolehlivější serializace do HTML)
    roky = [str(r) for r in YEARS]
    bezne = [float(v) for v in BEZNE]
    kapitalove = [float(v) for v in KAPITALOVE]
    celkem = [b + k for b, k in zip(bezne, kapitalove)]
    podil_kap = [k / c * 100 for k, c in zip(kapitalove, celkem)]

    PISMO = "Arial, Segoe UI, sans-serif"

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            name="Běžné výdaje",
            x=roky,
            y=bezne,
            marker=dict(color=MODRA, line=dict(width=0)),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Běžné výdaje: %{y:,.1f} mld. Kč"
                "<extra></extra>"
            )
        )
    )

    fig.add_trace(
        go.Bar(
            name="Kapitálové výdaje",
            x=roky,
            y=kapitalove,
            marker=dict(color=CERVENA, line=dict(width=0)),
            customdata=[format_procento(p) for p in podil_kap],
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Kapitálové výdaje: %{y:,.1f} mld. Kč<br>"
                "Podíl na výdajích: %{customdata}"
                "<extra></extra>"
            )
        )
    )

    # ========================================================
    # HODNOTY VE SLOUPCÍCH (tučně) A SOUČET NAD SLOUPCEM
    #
    # POZOR: popisky se umisťují podle POŘADÍ sloupce (0, 1, 2 …),
    # ne podle textu "2022".
    # ========================================================

    for i, (b, k, c) in enumerate(zip(bezne, kapitalove, celkem)):

        fig.add_annotation(
            x=i,
            y=b / 2,
            text=f"<b>{format_cislo(b, 1)}</b>",
            showarrow=False,
            font=dict(size=14, color="white", family=PISMO)
        )

        fig.add_annotation(
            x=i,
            y=b + k / 2,
            text=f"<b>{format_cislo(k, 1)}</b>",
            showarrow=False,
            font=dict(size=14, color="white", family=PISMO)
        )

        fig.add_annotation(
            x=i,
            y=c,
            text=f"<b>{format_cislo(c, 1)}</b>",
            showarrow=False,
            yanchor="bottom",
            yshift=5,
            font=dict(size=14, color=TEXT, family=PISMO)
        )

    # ========================================================
    # OSA Y PO 400, ZAOKROUHLENÁ NAHORU
    # ========================================================

    krok = 400

    horni = int(np.ceil(max(celkem) * 1.06 / krok) * krok)

    hodnoty_osy = list(range(0, horni + 1, krok))

    popisky_osy = [
        f"<b>{format_cislo(v, 0)}</b>"
        for v in hodnoty_osy
    ]

    fig.update_layout(

        title=dict(
            text="Běžné a kapitálové výdaje",
            x=0,
            xanchor="left",
            **TITULEK_POZICE,
            font=dict(size=17, color=TEXT)
        ),

        barmode="stack",

        # Legenda nahoře vlevo, vodorovně
        showlegend=True,
        legend=dict(
            orientation="h",
            x=0,
            y=1.0,
            xanchor="left",
            yanchor="bottom",
            font=dict(size=15, color=TEXT, family=PISMO),
            itemsizing="constant",
            traceorder="normal"
        ),

        height=vyska,
        autosize=True,

        margin=dict(l=75, r=15, t=TITULEK_VYSKA + 45, b=45),

        separators=", ",

        paper_bgcolor="white",
        plot_bgcolor="white",

        bargap=0.4,

        xaxis=dict(
            type="category",
            range=[-0.5, len(roky) - 0.5],
            tickvals=list(range(len(roky))),
            ticktext=[f"<b>{r}</b>" for r in roky],
            tickfont=dict(size=14, color=TEXT, family=PISMO),
            showgrid=False,
            showline=True,
            linecolor=TEXT,
            linewidth=1,
            ticks="outside",
            tickcolor=TEXT,
            ticklen=5
        ),

        yaxis=dict(
            title=dict(
                text="<b>mld. Kč</b>",
                font=dict(size=14, color=TEXT, family=PISMO)
            ),
            range=[0, horni],
            tickvals=hodnoty_osy,
            ticktext=popisky_osy,
            tickfont=dict(size=14, color=TEXT, family=PISMO),
            showgrid=False,
            zeroline=False,
            showline=True,
            linecolor=TEXT,
            linewidth=1,
            ticks="outside",
            tickcolor=TEXT,
            ticklen=5
        )

    )

    # ========================================================
    # PŘEPNUTÍ DESKTOP / MOBIL
    #   mobil: hodnoty uvnitř sloupců se skryjí (jsou v bublině
    #          po klepnutí), nad sloupcem zůstane jen součet
    #          zaokrouhlený na celé mld. Kč, menší písmo
    # ========================================================

    anotace_desktop = [
        a.to_plotly_json()
        for a in fig.layout.annotations
    ]

    anotace_mobil = [
        dict(
            x=i,
            y=c,
            text=f"<b>{format_cislo(c, 0)}</b>",
            showarrow=False,
            yanchor="bottom",
            yshift=3,
            font=dict(size=10, color=TEXT, family=PISMO)
        )
        for i, c in enumerate(celkem)
    ]

    prepinani = {
        "desktop": {
            "data": {},
            "layout": {
                "annotations": anotace_desktop,
                "height": vyska,
                "title.font.size": 17,
                "legend.font.size": 15,
                "margin.l": 75,
                "margin.t": TITULEK_VYSKA + 45,
                "xaxis.tickfont.size": 14,
                "yaxis.tickfont.size": 14,
                "yaxis.title.font.size": 14
            }
        },
        "mobil": {
            "data": {},
            "layout": {
                "annotations": anotace_mobil,
                "height": MOBIL_VYSKA_BK,
                "title.font.size": 15,
                "legend.font.size": 12,
                "margin.l": 50,
                "margin.t": TITULEK_VYSKA + 30,
                "xaxis.tickfont.size": 10,
                "yaxis.tickfont.size": 10,
                "yaxis.title.font.size": 11
            }
        }
    }

    return fig, prepinani


# ============================================================
# GRAF 5 – OBRANA: ČESKO A USA, CESTA K 5 % HDP DO ROKU 2035
#
# Sloupce = rovnoměrná cesta Česka k cíli (výchozí rok a
# zvýrazněný rok sytě červeně), červená čára = trend ČR,
# modrá čára = USA, přerušovaná čára = cíl NATO.
# ============================================================

def create_obrana_nato_chart(obrana):

    roky = obrana["roky"]
    cesta_cr = [float(v) for v in obrana["cesta_cr"]]
    cesta_usa = [float(v) for v in obrana["cesta_usa"]]

    zvyraznit = (OBRANA_VYCHOZI_ROK, OBRANA_ZVYRAZNENY_ROK)

    barvy = [
        CERVENA if r in zvyraznit else CERVENA_2
        for r in roky
    ]

    # Značky jen na začátku a na konci čar
    znacky = [9 if i in (0, len(roky) - 1) else 0 for i in range(len(roky))]

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=roky,
            y=cesta_cr,
            width=0.68,
            marker=dict(color=barvy, line=dict(width=0)),
            hovertext=[
                f"<b>{r}</b><br>Česko: {format_procento(v, 2)} HDP"
                + ("<br>(stav)" if r == OBRANA_VYCHOZI_ROK else "<br>(potřebná úroveň)")
                for r, v in zip(roky, cesta_cr)
            ],
            hovertemplate="%{hovertext}<extra></extra>"
        )
    )

    fig.add_trace(
        go.Scatter(
            x=roky,
            y=cesta_cr,
            mode="lines+markers",
            line=dict(color=CERVENA, width=3),
            marker=dict(size=znacky, color=CERVENA),
            hoverinfo="skip"
        )
    )

    fig.add_trace(
        go.Scatter(
            x=roky,
            y=cesta_usa,
            mode="lines+markers",
            line=dict(color=MODRA, width=3.5),
            marker=dict(size=znacky, color=MODRA),
            hovertext=[
                f"<b>{r}</b><br>USA: {format_procento(v, 2)} HDP"
                for r, v in zip(roky, cesta_usa)
            ],
            hovertemplate="%{hovertext}<extra></extra>"
        )
    )

    # ========================================================
    # CÍL NATO
    # ========================================================

    zacatek = OBRANA_VYCHOZI_ROK - 0.5
    konec = OBRANA_CIL_ROK + 0.5

    fig.add_shape(
        type="line",
        x0=zacatek, x1=konec,
        y0=OBRANA_CIL, y1=OBRANA_CIL,
        line=dict(color=SVETLE_SEDA, width=2, dash="dash"),
        layer="below"
    )

    PISMO = "Arial, Segoe UI, sans-serif"

    anotace = [
        dict(
            x=konec, y=OBRANA_CIL,
            text=f"<b>Cíl NATO {format_cislo(OBRANA_CIL, 0)} % v roce {OBRANA_CIL_ROK}</b>",
            xanchor="right", yanchor="bottom", yshift=4,
            font=dict(size=13, color=SVETLE_SEDA, family=PISMO)
        ),
        # Popisky vlevo – stav ve výchozím roce
        dict(
            x=zacatek, y=obrana["usa"],
            text=f"<b>USA {format_procento(obrana['usa'], 2)}</b>",
            xanchor="right", xshift=-6,
            font=dict(size=13, color=MODRA, family=PISMO)
        ),
        dict(
            x=zacatek, y=obrana["cesko"],
            text=f"<b>Česko {format_procento(obrana['cesko'], 2)}</b>",
            xanchor="right", xshift=-6,
            font=dict(size=13, color=CERVENA, family=PISMO)
        ),
        dict(
            x=zacatek, y=(obrana["usa"] + obrana["cesko"]) / 2,
            text=f"<b>rozdíl {format_cislo(obrana['rozdil'], 2)} p. b.</b>",
            xanchor="right", xshift=-6,
            font=dict(size=11, color=SVETLE_SEDA, family=PISMO)
        ),
        # Hodnota v zvýrazněném roce nad sloupcem
        dict(
            x=OBRANA_ZVYRAZNENY_ROK, y=obrana["cesko_zvyrazneny"],
            text=f"<b>{format_procento(obrana['cesko_zvyrazneny'], 2)}</b>",
            yanchor="bottom", yshift=8,
            bgcolor="white",
            font=dict(size=13, color=CERVENA, family=PISMO)
        ),
    ]

    for a in anotace:
        fig.add_annotation(showarrow=False, **a)

    # Osa X: jen vybrané roky, zvýrazněný rok červeně
    popisky_let = [OBRANA_VYCHOZI_ROK, OBRANA_ZVYRAZNENY_ROK, 2030, OBRANA_CIL_ROK]

    ticktext = [
        f"<b><span style='color:{CERVENA}'>{r}</span></b>"
        if r == OBRANA_ZVYRAZNENY_ROK else f"<b>{r}</b>"
        for r in popisky_let
    ]

    nadpis = f"Výdaje na obranu – cesta k cíli NATO {format_cislo(OBRANA_CIL, 0)} % HDP"

    fig.update_layout(

        title=dict(
            text=nadpis,
            x=0,
            xanchor="left",
            **TITULEK_POZICE,
            font=dict(size=17, color=TEXT)
        ),

        showlegend=False,

        height=OBRANA_VYSKA,
        autosize=True,

        margin=dict(l=10, r=15, t=TITULEK_VYSKA + 10, b=40),

        separators=", ",

        paper_bgcolor="white",
        plot_bgcolor="white",

        xaxis=dict(
            range=[OBRANA_VYCHOZI_ROK - 3.7, konec + 0.1],
            tickvals=popisky_let,
            ticktext=ticktext,
            tickfont=dict(size=14, color=TEXT, family=PISMO),
            showgrid=False,
            zeroline=False,
            showline=False,
            ticks=""
        ),

        yaxis=dict(
            range=[0, OBRANA_CIL * 1.12],
            visible=False
        )

    )

    # Čára osy X jen pod sloupci
    fig.add_shape(
        type="line",
        x0=zacatek, x1=konec, y0=0, y1=0,
        line=dict(color=TEXT, width=1.4)
    )

    # ========================================================
    # PŘEPNUTÍ DESKTOP / MOBIL
    #   mobil: menší písmo popisků, kratší nadpis
    # ========================================================

    anotace_desktop = [a.to_plotly_json() for a in fig.layout.annotations]

    anotace_mobil = []

    for a in anotace_desktop:
        a = dict(a)
        a["font"] = dict(a["font"], size=max(9, a["font"]["size"] - 3))
        anotace_mobil.append(a)

    prepinani = {
        "desktop": {
            "data": {},
            "layout": {
                "annotations": anotace_desktop,
                "height": OBRANA_VYSKA,
                "title.text": nadpis,
                "title.font.size": 17,
                "xaxis.tickfont.size": 14
            }
        },
        "mobil": {
            "data": {},
            "layout": {
                "annotations": anotace_mobil,
                "height": MOBIL_VYSKA_OBRANA,
                "title.text": f"Obrana – cesta k {format_cislo(OBRANA_CIL, 0)} % HDP",
                "title.font.size": 15,
                "xaxis.tickfont.size": 10
            }
        }
    }

    return fig, prepinani


# ============================================================
# HTML KARTA
# ============================================================

def html_card(title, value, subtitle):

    return f"""

    <div class="card">
        <div class="card-title">{title}</div>
        <div class="card-value">{value}</div>
        <div class="card-subtitle">{subtitle}</div>
    </div>

    """


# ============================================================
# HTML KARTA UKAZATELE SE ŠIPKOU
# ============================================================

def html_card_ukazatel(
    title,
    hodnota,
    popisek,
    rust_je_dobry,
    porovnat_s_inflaci,
    ikona=None
):

    # Směr šipky podle znaménka hodnoty
    if hodnota > 0:
        sipka = "▲"
        dobre = rust_je_dobry
    elif hodnota < 0:
        sipka = "▼"
        dobre = not rust_je_dobry
    else:
        sipka = "▶"
        dobre = None

    # Barva šipky
    if dobre is None:
        barva = TEXT_LIGHT
    elif dobre:
        barva = ZELENA_TMAVA
    else:
        barva = CERVENA

    # Volitelné porovnání s inflací (v procentních bodech)
    porovnani = ""

    if porovnat_s_inflaci:

        rozdil = hodnota - INFLACE_2027
        znamenko = "+" if rozdil >= 0 else "−"
        barva_rozdilu = ZELENA_TMAVA if rozdil >= 0 else CERVENA

        porovnani = f"""
            <span class="card-badge" style="color:{barva_rozdilu}">
                {znamenko}{format_cislo(abs(rozdil), 1)} p. b. oproti inflaci
            </span>
        """

    # Ikona v kroužku (barva podle vývoje, světlé pozadí)
    ikona_html = ""

    if ikona in IKONY:

        ikona_html = f"""
        <div class="card-ikona" style="background:{barva}1A">
            <svg viewBox="0 0 24 24" style="stroke:{barva}">
                {IKONY[ikona]}
            </svg>
        </div>
        """

    return f"""

    <div class="card card-ukazatel" style="border-top-color:{barva}">
        <div class="card-ukazatel-text">
            <div class="card-title">{title}</div>
            <div class="card-value-row">
                <span class="card-arrow" style="color:{barva}">{sipka}</span>
                <span class="card-value">{format_procento(hodnota)}</span>
                {porovnani}
            </div>
            <div class="card-subtitle">{popisek}</div>
        </div>
        {ikona_html}
    </div>

    """


# ============================================================
# KARTA: PODÍL OBSLUHY DLUHU NA VÝDAJÍCH SR
# ============================================================

def html_card_dluh_podil():

    if OBSLUHA_DLUHU is None:
        return """

    <div class="card card-volna">
        <div class="card-title">OBSLUHA DLUHU / VÝDAJE SR</div>
        <div class="card-value">doplní se</div>
        <div class="card-subtitle">podíl na výdajích státního rozpočtu</div>
    </div>

        """

    podil = OBSLUHA_DLUHU / VYDAJE_SR_CELKEM * 100

    # Proužek: celá šířka = 100 % výdajů SR
    sirka = min(podil, 100)

    # Od prahu (např. 5 %) červeně, jinak zlatá
    barva = CERVENA if podil >= PRAH_DLUH_CERVENA else ZLATA

    return f"""

    <div class="card card-nej" style="border-left-color:{barva}">
        <div class="card-title">OBSLUHA DLUHU / VÝDAJE SR</div>
        <div class="card-value" style="color:{barva}">{format_procento(podil)}</div>
        <div class="card-subtitle">
            každá {format_cislo(100 / podil, 0)}. koruna výdajů státního rozpočtu
            jde na obsluhu dluhu
        </div>
        <div class="podil-pruh">
            <div class="podil-pruh-vypln" style="width:{sirka:.1f}%; background:{barva}"></div>
        </div>
    </div>

    """


# ============================================================
# HTML DETAIL 13
# ============================================================

def html_detail13(prijmy):

    html = """

    <div class="panel">

        <div class="panel-title">
            Detail položky 13
        </div>

        <div class="panel-subtitle">
            Daně a poplatky z vybraných činností a služeb
        </div>

    """

    for label, value, share in zip(
        prijmy["detail13_labels"],
        prijmy["detail13_values"],
        prijmy["detail13_shares"]
    ):

        html += f"""

        <div class="detail-row">
            <div class="square"></div>
            <div class="detail-name">{label}</div>
            <div class="detail-value">{format_cislo(value, 2)} mld.</div>
            <div class="detail-share">({format_procento(share)})</div>
        </div>

        """

    html += f"""

        <div style="border-top:1px solid #d0d0d0; margin-top:16px; padding-top:12px;">

            <div class="control-row">
                <span class="control-label">Součet detailu</span>
                <span class="control-value">
                    {format_cislo(prijmy["detail13_total"], 3)} mld. Kč
                </span>
            </div>

            <div class="control-row">
                <span class="control-label">Rozdíl proti položce 13</span>
                <span class="control-value">
                    {format_cislo(prijmy["rozdil13"], 6)} mld. Kč
                </span>
            </div>

        </div>

    </div>

    """

    return html


# ============================================================
# HTML PANEL S GRAFEM DEFICITU
# ============================================================

def html_deficit(graf_deficit, deficit):

    prvni_rok = deficit["roky"][0]
    posledni_rok = deficit["roky"][-1]

    roky_odhad = [
        str(r)
        for r, o in zip(deficit["roky"], deficit["odhad"])
        if o
    ]

    text_odhad = ", ".join(roky_odhad) if roky_odhad else "–"

    return f"""

    <div class="panel deficit-panel">

        {graf_deficit}

        <div class="deficit-legend">
            <span><i class="legend-box" style="background:{CERVENA}"></i>deficit – skutečnost</span>
            <span><i class="legend-box" style="background:{MODRA}"></i>přebytek – skutečnost</span>
            <span><i class="legend-box" style="background:{SVETLE_CERVENA}"></i>výhled ({text_odhad})</span>
        </div>

        <div class="panel-subtitle">
            Roční saldo státního rozpočtu {prvni_rok}–{posledni_rok} v mld. Kč
            (skutečnost k 31. 12.)
        </div>

    </div>

    """


# ============================================================
# VYTVOŘENÍ DASHBOARDU
# ============================================================

def create_dashboard_html(vydaje, prijmy, deficit, oze, obrana):

    # ========================================================
    # STEJNÝ POČET ŘÁDKŮ LEGENDY PRO OBA GRAFY
    # -> stejná výška grafů i stejná velikost donutů
    # ========================================================

    pocet_radku = max(
        len(vydaje["labels"]),
        len(prijmy["labels"])
    )

    # ========================================================
    # GRAF VÝDAJŮ
    # ========================================================

    fig_vydaje, prep_vydaje = create_donut(
        vydaje["values"],
        vydaje["labels"],
        vydaje["colors"],
        vydaje["shares"],
        vydaje["total"],
        "Výdaje podle odvětvového třídění",
        pocet_radku
    )

    # ========================================================
    # GRAF PŘÍJMŮ
    # (obsahuje i nedaňové a kapitálové příjmy a transfery)
    # ========================================================

    fig_prijmy, prep_prijmy = create_donut(
        prijmy["values"],
        prijmy["labels"],
        prijmy["colors"],
        prijmy["shares"],
        prijmy["total"],
        "Příjmy státního rozpočtu",
        pocet_radku
    )

    # ========================================================
    # GRAF DEFICITU
    # ========================================================

    fig_deficit, prep_deficit = create_deficit_chart(deficit)

    # ========================================================
    # GRAFY V PRAVÉM SLOUPCI – STEJNÁ VÝŠKA JAKO DONUTY
    # ========================================================

    vyska_donutu = (
        TITULEK_VYSKA
        + DONUT_PLOCHA
        + pocet_radku * RADEK_LEGENDY
        + 30
    )

    # (o 22 px nižší – pod grafem je ještě řádek se zdrojem,
    #  aby panel měl stejnou výšku jako donut vedle)
    fig_oze, prep_oze = create_oze_chart(oze, vyska_donutu - 22)

    fig_bk, prep_bk = create_bezne_kapitalove_chart(vyska_donutu)

    # ========================================================
    # GRAF OBRANY
    # ========================================================

    fig_nato, prep_nato = create_obrana_nato_chart(obrana)

    # ========================================================
    # GRAF → HTML
    # ========================================================

    plotly_config = {
        "responsive": True,
        "displaylogo": False,
        "modeBarButtonsToRemove": [
            "lasso2d",
            "select2d"
        ]
    }

    def do_html(fig, div_id, s_knihovnou=False):
        return fig.to_html(
            full_html=False,
            div_id=div_id,
            include_plotlyjs="cdn" if s_knihovnou else False,
            config=plotly_config
        )

    graf_vydaje = do_html(fig_vydaje, "graf_vydaje", s_knihovnou=True)
    graf_prijmy = do_html(fig_prijmy, "graf_prijmy")
    graf_deficit = do_html(fig_deficit, "graf_deficit")
    graf_oze = do_html(fig_oze, "graf_oze")
    graf_bk = do_html(fig_bk, "graf_bk")
    graf_nato = do_html(fig_nato, "graf_nato")

    # ========================================================
    # SKRIPT PRO PŘEPÍNÁNÍ DESKTOP / MOBIL
    #
    # Po načtení stránky a při změně velikosti okna (např.
    # otočení telefonu) se každému grafu nastaví desktopová
    # nebo mobilní úprava. Na mobilu se navíc vypne
    # přibližování tahem prstu, aby šla stránka posouvat.
    # ========================================================

    prepinani = {
        "graf_vydaje": prep_vydaje,
        "graf_prijmy": prep_prijmy,
        "graf_deficit": prep_deficit,
        "graf_oze": prep_oze,
        "graf_bk": prep_bk,
        "graf_nato": prep_nato
    }

    skript_mobil = (
        "<script>\n"
        "(function () {\n"
        "  var NASTAVENI = " + json.dumps(prepinani, ensure_ascii=False) + ";\n"
        "  var HRANICE = " + str(MOBIL_SIRKA) + ";\n"
        "  var stav = null;\n"
        "  function nastav() {\n"
        "    if (!window.Plotly) return;\n"
        "    var mobil = window.innerWidth < HRANICE;\n"
        "    var novy = mobil ? 'mobil' : 'desktop';\n"
        "    if (novy === stav) return;\n"
        "    stav = novy;\n"
        "    Object.keys(NASTAVENI).forEach(function (id) {\n"
        "      var div = document.getElementById(id);\n"
        "      if (!div) return;\n"
        "      var n = NASTAVENI[id][novy];\n"
        "      var layout = Object.assign({}, n.layout, {dragmode: mobil ? false : 'zoom'});\n"
        "      Plotly.update(div, n.data, layout);\n"
        "    });\n"
        "  }\n"
        "  window.addEventListener('load', nastav);\n"
        "  var casovac;\n"
        "  window.addEventListener('resize', function () {\n"
        "    clearTimeout(casovac);\n"
        "    casovac = setTimeout(nastav, 200);\n"
        "  });\n"
        "})();\n"
        "</script>\n"
    )

    # ========================================================
    # DATUM
    # ========================================================

    cas = datetime.now().strftime("%d.%m.%Y %H:%M")

    # ========================================================
    # KARTY UKAZATELŮ SE ŠIPKAMI (1. řádek)
    # ========================================================

    karty_ukazatele = "".join(
        html_card_ukazatel(*u)
        for u in UKAZATELE
    )

    # ========================================================
    # KPI KARTY (2. řádek – 5 karet)
    # ========================================================

    mld = 1_000_000_000

    if OBSLUHA_DLUHU is None:
        karta_dluh = """

    <div class="card card-volna">
        <div class="card-title">NÁKLADY NA OBSLUHU DLUHU</div>
        <div class="card-value">doplní se</div>
        <div class="card-subtitle">mld. Kč</div>
    </div>

        """
    else:
        karta_dluh = html_card(
            "NÁKLADY NA OBSLUHU DLUHU",
            format_cislo(OBSLUHA_DLUHU / mld, 1) + " mld. Kč",
            "úroky a poplatky z dluhu, návrh 2027"
        )

    barva_salda = CERVENA if SALDO_SR < 0 else MODRA

    karty = (
        html_card(
            "PŘÍJMY STÁTNÍHO ROZPOČTU CELKEM",
            format_cislo(PRIJMY_SR_CELKEM / mld, 1) + " mld. Kč",
            "návrh 2027"
        )
        + html_card(
            "VÝDAJE STÁTNÍHO ROZPOČTU CELKEM",
            format_cislo(VYDAJE_SR_CELKEM / mld, 1) + " mld. Kč",
            "návrh 2027"
        )
        + html_card(
            "SALDO STÁTNÍHO ROZPOČTU",
            f'<span style="color:{barva_salda}">'
            + format_cislo_znamenko(SALDO_SR / mld, 1)
            + " mld. Kč</span>",
            "deficit (příjmy − výdaje)" if SALDO_SR < 0 else "přebytek (příjmy − výdaje)"
        )
        + karta_dluh
        + html_card_dluh_podil()
    )

    # ========================================================
    # HTML
    # ========================================================

    html = f"""<!DOCTYPE html>

<html lang="cs">

<head>

    <meta charset="UTF-8">

    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <title>Státní rozpočet 2027</title>

    {html_css()}

</head>

<body>

<div class="dashboard">

    <!-- ==================================================
         HLAVIČKA
         ================================================== -->

    <div class="header">

        <div>

            <div class="title">
                Státní rozpočet 2027
            </div>

        </div>

        <div class="year">
            ČR • 2027
        </div>

    </div>


    <!-- ==================================================
         1. ŘÁDEK: UKAZATELE SE ŠIPKAMI
         ================================================== -->

    <div class="cards-ukazatele">

        {karty_ukazatele}

    </div>


    <!-- ==================================================
         2. ŘÁDEK: KPI + PODÍL OBSLUHY DLUHU
         ================================================== -->

    <div class="cards">

        {karty}

    </div>


    <!-- ==================================================
         HLAVNÍ ČÁST
         ================================================== -->

    <div class="main-grid">

        <!-- LEVÝ SLOUPEC: VÝDAJE A PŘÍJMY POD SEBOU -->

        <div class="column">

            <div class="panel">

                {graf_vydaje}

            </div>

            <div class="panel">

                {graf_prijmy}

            </div>

        </div>


        <!-- PRAVÝ SLOUPEC: OZE V EU A BĚŽNÉ / KAPITÁLOVÉ VÝDAJE -->

        <div class="column">

            <div class="panel">

                {graf_oze}

                <div class="panel-subtitle">
                    Země EU, {oze['obdobi']} • zdroj: Eurostat (nrg_cb_pem)
                </div>

            </div>

            <div class="panel">

                {graf_bk}

            </div>

        </div>


    </div>


    <!-- ==================================================
         VÝDAJE NA OBRANU (přes celou šířku)
         ================================================== -->

    <div class="panel obrana-panel">

        {graf_nato}

        <div class="panel-subtitle">
            Výdaje na obranu v % HDP, stav {OBRANA_VYCHOZI_ROK} a rovnoměrný
            nárůst k cíli {format_cislo(OBRANA_CIL, 0)} % v roce {OBRANA_CIL_ROK}
        </div>

    </div>


    <!-- ==================================================
         VÝVOJ DEFICITU (úplně dole)
         ================================================== -->

    {html_deficit(graf_deficit, deficit)}


    <!-- ==================================================
         PATIČKA
         ================================================== -->

    <div class="footer">

        <div>
            Dashboard vytvořen v Pythonu • pandas + numpy + plotly
        </div>

        <div>
            Aktualizováno: {cas} • Částky v mld. Kč
        </div>

    </div>

</div>

{skript_mobil}

</body>

</html>
"""

    return html


# ============================================================
# OTEVŘENÍ V PROHLÍŽEČI
# ============================================================

def open_in_browser(path):

    url = path.resolve().as_uri()

    print()
    print("Otevírám dashboard v internetovém prohlížeči...")
    print()
    print(url)
    print()

    webbrowser.open_new_tab(url)


# ============================================================
# HLAVNÍ PROGRAM
# ============================================================

if __name__ == "__main__":

    print()
    print("=" * 80)
    print("DASHBOARD STÁTNÍ ROZPOČET 2027")
    print("=" * 80)
    print()

    # ========================================================
    # KONTROLA SOUBORŮ
    # ========================================================

    for soubor, popis in [
        (SOUBOR_VYDAJE, "s výdaji"),
        (SOUBOR_PRIJMY, "s příjmy"),
        (SOUBOR_DEFICIT, "s deficitem"),
        (SOUBOR_OZE, "s obnovitelnými zdroji"),
        (SOUBOR_OBRANA, "s výdaji na obranu"),
    ]:

        if not soubor.exists():
            raise FileNotFoundError(
                f"Soubor {popis} nebyl nalezen:\n\n"
                + str(soubor)
            )

    # ========================================================
    # NAČTENÍ DAT
    # ========================================================

    print("Načítám výdaje...")
    vydaje = load_vydaje()

    print("Načítám příjmy...")
    prijmy = load_prijmy()

    print("Načítám deficit...")
    deficit = load_deficit()

    print("Načítám obnovitelné zdroje...")
    oze = load_oze()

    print("Načítám výdaje na obranu...")
    obrana = load_obrana()

    # ========================================================
    # VYTVOŘENÍ DASHBOARDU
    # ========================================================

    print("Vytvářím dashboard...")
    html = create_dashboard_html(vydaje, prijmy, deficit, oze, obrana)

    # ========================================================
    # ULOŽENÍ
    # ========================================================

    VYSTUP.write_text(html, encoding="utf-8")

    # ========================================================
    # VÝPIS VÝSLEDKŮ
    # ========================================================

    print()
    print("VÝSLEDKY")
    print("-" * 80)

    print("Příjmy SR celkem:", format_cislo(PRIJMY_SR_CELKEM / 1e9, 1), "mld. Kč")
    print("Výdaje SR celkem:", format_cislo(VYDAJE_SR_CELKEM / 1e9, 1), "mld. Kč")
    print("Saldo SR:", format_cislo_znamenko(SALDO_SR / 1e9, 1), "mld. Kč")

    if OBSLUHA_DLUHU is not None:
        print(
            "Obsluha dluhu / výdaje SR:",
            format_procento(OBSLUHA_DLUHU / VYDAJE_SR_CELKEM * 100)
        )

    print()
    print("Výdaje (donut):")
    for l, v, p in zip(vydaje["labels"], vydaje["values"], vydaje["shares"]):
        print(f"  {format_cislo(v, 1):>8} mld. Kč  {format_procento(p):>7}  {l}")
    print(f"  {format_cislo(vydaje['total'], 1):>8} mld. Kč  CELKEM")

    print()
    print("Příjmy (donut):")
    for l, v, p in zip(prijmy["labels"], prijmy["values"], prijmy["shares"]):
        print(f"  {format_cislo(v, 1):>8} mld. Kč  {format_procento(p):>7}  {l}")
    print(f"  {format_cislo(prijmy['total'], 1):>8} mld. Kč  CELKEM")

    print()
    print("Položka 13:", format_cislo(prijmy["skupina_13"], 3), "mld. Kč",
          f"(podpoložky 132–138: {format_cislo(prijmy['detail13_total'], 3)})")
    print("Položka 16:", format_cislo(prijmy["skupina_16"], 3), "mld. Kč")

    print()
    print("Ukazatele:")

    for nadpis, hodnota, *_ in UKAZATELE:
        print(f"  {nadpis}: {format_procento(hodnota)}")

    print()

    if CESKO in oze["zeme"]:
        print(f"Podíl OZE ({oze['obdobi']}): Česko "
              f"{format_procento(oze['hodnoty'][oze['zeme'].index(CESKO)])}, "
              f"EU-27 {format_procento(oze['eu'])}")

    print()
    print("Výdaje na obranu (% HDP):")
    print(f"  Česko {OBRANA_VYCHOZI_ROK}: {format_procento(obrana['cesko'], 2)}")
    print(f"  Česko {OBRANA_ZVYRAZNENY_ROK}: "
          f"{format_procento(obrana['cesko_zvyrazneny'], 2)} (potřebná úroveň)")
    print(f"  USA {OBRANA_VYCHOZI_ROK}: {format_procento(obrana['usa'], 2)}")
    print(f"  Rozdíl: {format_cislo(obrana['rozdil'], 2)} p. b.")

    print()
    print("Běžné / kapitálové výdaje (mld. Kč):")

    for rok, b, k in zip(YEARS, BEZNE, KAPITALOVE):
        print(f"  {rok}: běžné {format_cislo(b, 1):>8}, "
              f"kapitálové {format_cislo(k, 1):>6}, "
              f"celkem {format_cislo(b + k, 1):>8}")

    print()
    print("Saldo podle let (mld. Kč):")

    for rok, hodnota, odhad in zip(
        deficit["roky"],
        deficit["saldo"],
        deficit["odhad"]
    ):
        poznamka = "  (výhled)" if odhad else ""
        print(f"  {rok}: {format_cislo(hodnota, 1):>8}{poznamka}")

    print()
    print("Dashboard uložen do:")
    print(VYSTUP)
    print()
    print("=" * 80)

    # ========================================================
    # AUTOMATICKÉ OTEVŘENÍ
    # ========================================================

    open_in_browser(VYSTUP)
