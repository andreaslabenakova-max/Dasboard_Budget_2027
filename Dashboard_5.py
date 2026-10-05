# -*- coding: utf-8 -*-

"""
DASHBOARD STÁTNÍ ROZPOČET 2027 – WEBOVÝ PROHLÍŽEČ
=================================================

Po spuštění:

    python Dashboard_verejne_finance_2027_web.py

se:

1. načtou všechny tři Excel soubory (výdaje, příjmy, deficit),
2. vytvoří interaktivní HTML dashboard,
3. HTML se uloží,
4. dashboard se automaticky otevře v prohlížeči.

Instalace:

    pip install pandas numpy plotly openpyxl
"""

from pathlib import Path
import re
import webbrowser
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go


# ============================================================
# SOUBORY
# ============================================================

SOUBOR_VYDAJE = Path(
    r"C:\Users\SlabenakovaA\OneDrive - KPS CR\Plocha\Python\Rozpocet_Odvetvi.xlsx"
)

SOUBOR_PRIJMY = Path(
    r"C:\Users\SlabenakovaA\OneDrive - KPS CR\Plocha\Python\Dane_Extrakt.xlsx"
)

SOUBOR_DEFICIT = Path(
    r"C:\Users\SlabenakovaA\OneDrive - KPS CR\Plocha\Python\Deficit.xlsx"
)

SOUBOR_OZE = Path(
    r"C:\Users\SlabenakovaA\OneDrive - KPS CR\Plocha\Python\nrg_cb_pem__custom_22895279_spreadsheet.xlsx"
)

SHEET_VYDAJE = "Výdaje 2027-odvětvové"

SHEET_PRIJMY = "Tab.1 - příjmy"

SHEET_DEFICIT = "List1"

SHEET_OZE = "Sheet 1"


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
SVETLE_MODRA = "#98C3D0"
ZELENA = "#B3C0A7"
TMAVE_SEDA = "#535C55"
SEDA = "#2E808C"
ZLATA = "#84754E"
SVETLE_SEDA = "#8A8D8F"
TMAVE_MODRA = "#142B53"

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
# ROZMĚRY DONUTŮ (v pixelech, stejné pro oba grafy)
# ============================================================

DONUT_PLOCHA = 330      # výška plochy, ve které se kreslí donut = jeho průměr
TITULEK_VYSKA = 50      # horní okraj pro nadpis grafu
RADEK_LEGENDY = 24      # výška jednoho řádku legendy (pro písmo 14)
STRED_ODSTUP = 30       # rozestup řádků textu uprostřed donutu (px)

# Výška grafu deficitu
DEFICIT_VYSKA = 470

# Kolik posledních let zobrazit v grafu deficitu
DEFICIT_POCET_LET = 10


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
# FORMÁT PROCENTA
# ============================================================

def format_procento(value):

    return (
        f"{value:.1f}"
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
    # KÓDY
    # ========================================================

    codes_to_plot = [
        1,
        2,
        "31 a 32",
        33,
        34,
        35,
        36,
        37,
        38,
        39,
        4,
        5,
        6
    ]

    # ========================================================
    # NÁZVY
    # ========================================================

    labels = [
        "Zemědělství, lesní hospodářství a rybářství",
        "Průmyslová a ostatní odvětví hospodářství",
        "Vzdělávání a školské služby",
        "Kultura, církve a sdělovací prostředky",
        "Sport a zájmová činnost",
        "Zdravotnictví",
        "Bydlení, komunální služby a územní rozvoj",
        "Ochrana životního prostředí",
        "Ostatní výzkum a vývoj",
        "Ostatní činnosti související se službami pro fyzické osoby",
        "Sociální věci a politika zaměstnanosti",
        "Bezpečnost státu a právní ochrana",
        "Všeobecná veřejná správa a služby"
    ]

    # ========================================================
    # HODNOTY
    # ========================================================

    values_all = np.array([
        get_value(code)
        for code in codes_to_plot
    ])

    # ========================================================
    # ODSTRANĚNÍ NUL
    # ========================================================

    mask = values_all > 0

    values = values_all[mask]

    plot_codes = [
        code
        for code, keep in zip(codes_to_plot, mask)
        if keep
    ]

    plot_labels = [
        label
        for label, keep in zip(labels, mask)
        if keep
    ]

    # ========================================================
    # CELKEM
    # ========================================================

    total = values.sum()

    if total > 0:
        shares = values / total * 100
    else:
        shares = np.zeros_like(values)

    # ========================================================
    # BARVY
    # ========================================================

    color_map = {
        1: MODRA,
        2: CERVENA,
        "31 a 32": SEDA,
        33: SVETLE_MODRA,
        34: SVETLE_MODRA,
        35: SVETLE_MODRA,
        36: SVETLE_MODRA,
        37: SVETLE_MODRA,
        38: SVETLE_MODRA,
        39: SVETLE_MODRA,
        4: ZELENA,
        5: TMAVE_SEDA,
        6: SVETLE_SEDA
    }

    colors = [
        color_map[code]
        for code in plot_codes
    ]

    return {
        "values": values,
        "labels": plot_labels,
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
    # NALEZENÍ ŘÁDKU PODLE KÓDU
    # ========================================================

    def najdi_radek(kod):

        kod = str(kod).strip()

        vzor = re.compile(
            r"^\s*"
            + re.escape(kod)
            + r"(?=\s|$)"
        )

        for i in range(len(df)):

            bunka = str(df.iloc[i, 1])

            if vzor.search(bunka):
                return i

        return None

    # ========================================================
    # HODNOTA 2027
    # ========================================================

    def hodnota_2027(kod):

        radek = najdi_radek(kod)

        if radek is None:
            print(
                f"VAROVÁNÍ – kód {kod} "
                "nebyl nalezen."
            )
            return 0.0

        hodnota = df.iloc[radek, 7]

        if pd.isna(hodnota):
            return 0.0

        return float(hodnota) / 1_000_000_000

    # ========================================================
    # NÁZEV PŘESNĚ PODLE SLOUPCE B
    #
    # Vezme text buňky ve sloupci B na řádku daného kódu
    # (jen odstraní okrajové a zdvojené mezery). Když kód
    # nenajde, použije záložní název ze skriptu.
    # ========================================================

    def nazev_z_excelu(kod, zalozni):

        radek = najdi_radek(kod)

        if radek is None:
            return zalozni

        text = df.iloc[radek, 1]

        if pd.isna(text):
            return zalozni

        return " ".join(str(text).split())

    # ========================================================
    # HLAVNÍ SKUPINY
    # ========================================================

    labels_all = [
        "DPFO",
        "DPPO",
        "Ostatní přímé daně",
        "DPH",
        "Zvláštní daně, poplatky a obdobná peněžitá plnění",
        "Daně a poplatky z vybraných činností a služeb",
        "Daně a cla ze zahraničí",
        "Majetkové daně",
        "Ostatní daňové příjmy",
        "Příjem z povinného pojistného"
    ]

    main_kody = [
        "111",
        "112",
        "113",
        "1211",
        "122, 123, 124",
        "13",
        "14",
        "15",
        "17",
        "16"
    ]

    # Ruční přejmenování vybraných položek (kód -> název v grafu).
    # Ostatní položky mají název přesně podle sloupce B.
    prejmenovani = {
        "111": "DPFO",
        "112": "DPPO",
        "1211": "DPH"
    }

    labels_all = [
        prejmenovani.get(kod, nazev_z_excelu(kod, label))
        for kod, label in zip(main_kody, labels_all)
    ]

    # ========================================================
    # HODNOTY
    # ========================================================

    values_all = np.array([
        hodnota_2027(kod)
        for kod in main_kody
    ])

    # ========================================================
    # ODSTRANĚNÍ NUL
    # ========================================================

    mask = values_all > 0

    values = values_all[mask]

    labels = [
        label
        for label, keep in zip(labels_all, mask)
        if keep
    ]

    codes = [
        code
        for code, keep in zip(main_kody, mask)
        if keep
    ]

    # ========================================================
    # CELKEM
    # ========================================================

    total = values.sum()

    if total > 0:
        shares = values / total * 100
    else:
        shares = np.zeros_like(values)

    # ========================================================
    # BARVY
    # ========================================================

    colors_all = [
        MODRA,
        CERVENA,
        TMAVE_SEDA,
        SVETLE_MODRA,
        ZELENA,
        SEDA,
        ZLATA,
        SVETLE_SEDA,
        TMAVE_SEDA,
        TMAVE_MODRA
    ]

    colors = [
        color
        for color, keep in zip(colors_all, mask)
        if keep
    ]

    # ========================================================
    # DETAIL SKUPINY 13
    # ========================================================

    detail13_kody = [
        "132",
        "133",
        "134",
        "135",
        "136",
        "137",
        "138"
    ]

    detail13_labels = [
        "Provoz motorových vozidel",
        "Životní prostředí",
        "Místní poplatky",
        "Ostatní odvody",
        "Správní a soudní poplatky",
        "Poplatky na činnost správních úřadů",
        "Hazardní hry"
    ]

    detail13_labels = [
        nazev_z_excelu(kod, label)
        for kod, label in zip(detail13_kody, detail13_labels)
    ]

    detail13_values_all = np.array([
        hodnota_2027(kod)
        for kod in detail13_kody
    ])

    mask13 = detail13_values_all > 0

    detail13_values = detail13_values_all[mask13]

    detail13_kody_final = [
        kod
        for kod, keep in zip(detail13_kody, mask13)
        if keep
    ]

    detail13_labels_final = [
        label
        for label, keep in zip(detail13_labels, mask13)
        if keep
    ]

    # ========================================================
    # SKUPINY 13 A 16
    # ========================================================

    skupina_13 = hodnota_2027("13")

    skupina_16 = hodnota_2027("16")

    # ========================================================
    # PODÍLY DETAILU 13
    # ========================================================

    if skupina_13 > 0:
        detail13_shares = detail13_values / skupina_13 * 100
    else:
        detail13_shares = np.zeros(len(detail13_values))

    # ========================================================
    # KONTROLA
    # ========================================================

    detail13_total = detail13_values.sum()

    rozdil13 = skupina_13 - detail13_total

    # ========================================================
    # DONUT PŘÍJMŮ: SKUPINA 13 ROZPADLÁ NA DETAIL
    #
    # Místo jedné výseče "Daně a poplatky z vybraných činností
    # a služeb" se v grafu zobrazí její jednotlivé položky
    # (132–138). Případný zbytek (skupina 13 minus součet
    # detailu) se přidá jako "Ostatní (položka 13)", takže
    # celkový součet se nezmění.
    # ========================================================

    graf_values = []
    graf_labels = []
    graf_codes = []
    graf_colors = []

    for value, label, code, color in zip(values, labels, codes, colors):

        if code != "13":
            graf_values.append(value)
            graf_labels.append(label)
            graf_codes.append(code)
            graf_colors.append(color)
            continue

        for d_value, d_label, d_code in zip(
            detail13_values,
            detail13_labels_final,
            detail13_kody_final
        ):
            graf_values.append(d_value)
            graf_labels.append(d_label)
            graf_codes.append(d_code)
            graf_colors.append(SEDA)

        if rozdil13 > 0.0005:
            graf_values.append(rozdil13)
            graf_labels.append("Ostatní (položka 13)")
            graf_codes.append("13 – zbytek")
            graf_colors.append(SEDA)

    graf_values = np.array(graf_values)

    if total > 0:
        graf_shares = graf_values / total * 100
    else:
        graf_shares = np.zeros_like(graf_values)

    return {
        "values": graf_values,
        "labels": graf_labels,
        "codes": graf_codes,
        "shares": graf_shares,
        "colors": graf_colors,
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
}

.card-volna .card-value {
    color: #aaaaaa;
    font-weight: 400;
    font-size: 18px;
}


/* ============================================================
   KARTA S PODÍLEM OBSLUHY DLUHU
   ============================================================ */

.card-nej {
    border-left: 5px solid #0055A0;
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
        padding: 18px;
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

    # Text bubliny po najetí myší – hotový řetězec pro každou výseč
    hover_texty = [
        f"Částka: {format_cislo(v, 2)} mld. Kč<br>"
        f"Podíl: {format_procento(p)}"
        for v, p in zip(values, shares)
    ]

    # Procenta do výsečí – u malých výsečí (< 2 %) se nic nepíše
    text_vyseci = [
        format_procento(p) if p >= 2 else ""
        for p in shares
    ]

    # Spodní okraj pro legendu – stejný u obou grafů
    spodni_okraj = pocet_radku_legendy * RADEK_LEGENDY + 30

    vyska = TITULEK_VYSKA + DONUT_PLOCHA + spodni_okraj

    fig = go.Figure()

    fig.add_trace(

        go.Pie(

            values=values,
            labels=labels,

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
                "<b>%{label}</b><br>"
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

    return fig


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

    return fig


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
            text=(
                "Podíl obnovitelných zdrojů na výrobě elektřiny"
                f"<br><span style='font-size:11px;color:{TEXT_LIGHT}'>"
                f"Země EU, {oze['obdobi']} • zdroj: Eurostat (nrg_cb_pem)"
                "</span>"
            ),
            x=0,
            xanchor="left",
            font=dict(size=17, color=TEXT)
        ),

        showlegend=False,

        height=vyska,
        autosize=True,

        margin=dict(l=10, r=55, t=TITULEK_VYSKA + 35, b=35),

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

    return fig


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

    return fig


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

def create_dashboard_html(vydaje, prijmy, deficit, oze):

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

    fig_vydaje = create_donut(
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
    # ========================================================

    fig_prijmy = create_donut(
        prijmy["values"],
        prijmy["labels"],
        prijmy["colors"],
        prijmy["shares"],
        prijmy["total"],
        "Příjmy z daní a poplatků",
        pocet_radku
    )

    # ========================================================
    # GRAF DEFICITU
    # ========================================================

    fig_deficit = create_deficit_chart(deficit)

    # ========================================================
    # GRAFY V PRAVÉM SLOUPCI – STEJNÁ VÝŠKA JAKO DONUTY
    # ========================================================

    vyska_donutu = (
        TITULEK_VYSKA
        + DONUT_PLOCHA
        + pocet_radku * RADEK_LEGENDY
        + 30
    )

    fig_oze = create_oze_chart(oze, vyska_donutu)

    fig_bk = create_bezne_kapitalove_chart(vyska_donutu)

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

    graf_vydaje = fig_vydaje.to_html(
        full_html=False,
        include_plotlyjs="cdn",
        config=plotly_config
    )

    graf_prijmy = fig_prijmy.to_html(
        full_html=False,
        include_plotlyjs=False,
        config=plotly_config
    )

    graf_deficit = fig_deficit.to_html(
        full_html=False,
        include_plotlyjs=False,
        config=plotly_config
    )

    graf_oze = fig_oze.to_html(
        full_html=False,
        include_plotlyjs=False,
        config=plotly_config
    )

    graf_bk = fig_bk.to_html(
        full_html=False,
        include_plotlyjs=False,
        config=plotly_config
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

            </div>

            <div class="panel">

                {graf_bk}

            </div>

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

    # ========================================================
    # VYTVOŘENÍ DASHBOARDU
    # ========================================================

    print("Vytvářím dashboard...")
    html = create_dashboard_html(vydaje, prijmy, deficit, oze)

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
    print("Výdaje celkem (odvětvové):", format_cislo(vydaje["total"], 3), "mld. Kč")
    print("Příjmy celkem:", format_cislo(prijmy["total"], 3), "mld. Kč")
    print("Položka 13:", format_cislo(prijmy["skupina_13"], 3), "mld. Kč")
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