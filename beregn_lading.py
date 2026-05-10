#!/usr/bin/env python3
"""
Elbillading - strømkostnad beregning, Køhnkvartalet Sameie

Leser CloudCharge CSV-filer fra mappen CloudCharge/ og strømpris-PDF-er
fra mappen Faktura/, beregner kostnad per ladepunkt per måned og skriver
ut Excel-fil til regnskapsfører.

Bruk:
    python beregn_lading.py

Mappestruktur:
    CloudCharge/   - CSV-filer lastet ned fra CloudCharge-portalen
    Faktura/       - PDF-fakturaer fra strømleverandør (Ustekveikja Energi)

Måneder uten registrert strømpris utelates automatisk fra rapporten.
Strømpriser leses fra PDF-fakturaer. Eldre måneder uten PDF-faktura
hentes fra Excel-fanen Strømpriser som fallback.
"""

import os
import glob
import re
from datetime import datetime

import pandas as pd
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MAPPE = os.path.dirname(os.path.abspath(__file__))
EXCEL_FIL = os.path.join(MAPPE, "Elbillading Strøm+Beboere.xlsx")
CLOUDCHARGE_MAPPE = os.path.join(MAPPE, "CloudCharge")
FAKTURA_MAPPE = os.path.join(MAPPE, "Faktura")

MÅNEDER = {
    1: "Januar", 2: "Februar", 3: "Mars", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

# Påslag på spotpris for å dekke usikkerhet rundt ladetidspunkt på døgnet
PRISPÅSLAG = 1.20


# ── Strømpriser ────────────────────────────────────────────────────────────

def _sf(s):
    """Norsk desimalstreng → float ('2 895,01' → 2895.01)."""
    return float(str(s).replace("\xa0", "").replace(" ", "").replace(",", "."))


FORVENTET_ANLEGGSREFERANSE = "Strøm Fellesareal"


def les_pris_fra_pdf(filsti):
    """
    Leser strømpriskomponenter fra Ustekveikja Energi-faktura (PDF) og
    beregner Totalt pr/kWh etter samme formel som Excel-regnearket:

        spot_m_mva = spotpris × 1.25
        strømstøtte_øre = strømstøtte_kr / forbruk_kwh × 100
        nettleie_m_mva = (energiledd_hverdag + elavgift) × 1.25
        totalt = spot_m_mva − strømstøtte_øre + nettleie_m_mva

    Returnerer (år, måned, totalt_pr_kwh) eller None ved parsefeil eller
    feil anleggsreferanse.
    """
    filnavn = os.path.basename(filsti)
    with pdfplumber.open(filsti) as pdf:
        tekst = "\n".join(p.extract_text() or "" for p in pdf.pages)

    # Sjekk at fakturaen gjelder riktig anlegg (Strøm Fellesareal)
    m_anlegg = re.search(r"Anleggsreferanse\s*[:\-]?\s*(.+)", tekst)
    if not m_anlegg:
        print(f"  ! {filnavn}: Fant ikke Anleggsreferanse — hopper over.")
        return None
    anleggsref = m_anlegg.group(1).strip()
    if FORVENTET_ANLEGGSREFERANSE.lower() not in anleggsref.lower():
        print(f"  ! {filnavn}: Anleggsreferanse er '{anleggsref}' "
              f"(forventet '{FORVENTET_ANLEGGSREFERANSE}') — hopper over.")
        return None

    # Periode, totalt forbruk og spotpris fra Strømpris-linjen
    m = re.search(
        r"Strømpris\s+(\d{2})\.(\d{2})\.(\d{2})-(\d{2})\.(\d{2})\.(\d{2})\s+"
        r"([\d ,]+,\d+)\s+kWh\s+([\d,]+)\s+øre/kWh",
        tekst,
    )
    if not m:
        print(f"  ! {filnavn}: Fant ikke faktureringsperiode — hopper over.")
        return None
    måned_nr = int(m.group(2))
    år = 2000 + int(m.group(3))
    slutt_måned = int(m.group(5))
    slutt_år = 2000 + int(m.group(6))
    forbruk_kwh = _sf(m.group(7))   # totalt forbruk — nevner for strømstøtte-beregning
    spot_ore = _sf(m.group(8))

    # Beregn månedsdifferanse — normal faktura: start og slutt er én måned fra hverandre
    # (f.eks. 01.04.26–01.05.26). Advar kun ved mer enn 1 måneds differanse.
    if (slutt_år * 12 + slutt_måned) - (år * 12 + måned_nr) > 1:
        print(f"  ! {filnavn}: Perioden dekker mer enn én måned — kontroller fakturaen manuelt.")
        return None

    # Dato-prefix brukt for å låse alle linjesøk til riktig faktureringsperiode.
    # Fakturaer med korreksjonsrader for andre måneder (f.eks. etterregning) inneholder
    # linjer for flere perioder — uten låsing plukkes feil linje opp.
    dato = rf"\d{{2}}\.{måned_nr:02d}\.{år % 100:02d}"

    # Strømstøtte: nettobeløp i kr delt på totalt forbruk kWh (fra Strømpris-linja).
    # "borettslag" er valgfritt: eldre fakturaer har det på samme linje som datoen,
    # nyere fakturaer har det på neste linje (da fanget opp av \s+).
    # Perioden låses slik at korreksjonsposter for andre måneder ikke plukkes opp.
    m = re.search(
        rf"Midlertidig str[øo]mst[øo]nad for(?:\s+borettslag)?\s+"
        rf"{dato}-\d{{2}}\.\d{{2}}\.\d{{2}}\s+"
        rf"-[\d ,]+kWh\s+[\d,]+\s+øre/kWh\s+\d+\s+(-[\d ]+,\d{{2}})(?=\s)",
        tekst,
    )
    stonad_kr = abs(_sf(m.group(1))) if m else 0.0
    stromstotte_ore = (stonad_kr / forbruk_kwh * 100) if forbruk_kwh > 0 else 0.0

    # Energiledd hverdag øre/kWh — låst til riktig periode
    m = re.search(
        rf"Energiledd hverdag[^\n]*?{dato}-\d{{2}}\.\d{{2}}\.\d{{2}}[^\n]*?kWh\s+([\d,]+)\s+øre/kWh",
        tekst,
    )
    energiledd_ore = _sf(m.group(1)) if m else 0.0

    # Elavgift øre/kWh — låst til riktig periode
    m = re.search(
        rf"Elavgift\s+{dato}-\d{{2}}\.\d{{2}}\.\d{{2}}\s+[\d ,]+kWh\s+([\d,]+)\s+øre/kWh",
        tekst,
    )
    elavgift_ore = _sf(m.group(1)) if m else 0.0

    spot_m_mva = spot_ore * 1.25
    nettleie_m_mva = (energiledd_ore + elavgift_ore) * 1.25
    totalt_pr_kwh = round(spot_m_mva - stromstotte_ore + nettleie_m_mva, 4)

    print(f"  PDF {filnavn}: {MÅNEDER[måned_nr]} {år} (anlegg: {anleggsref})")
    return (år, måned_nr, totalt_pr_kwh)


def les_strompriser():
    """
    Leser strømpriser fra PDF-fakturaer (primær) og Excel (fallback).
    PDF-priser overstyrer Excel for samme måned.
    Skriver ut hvilken kilde som brukes per måned.
    """
    # Fallback: Excel-fanen Strømpriser
    excel_priser = {}
    wb = openpyxl.load_workbook(EXCEL_FIL, data_only=True)
    ws = wb["Strømpriser"]
    for row in ws.iter_rows(values_only=True):
        dato, pris = row[0], row[10]
        if isinstance(dato, datetime) and isinstance(pris, (int, float)):
            excel_priser[(dato.year, dato.month)] = pris

    # Primær: PDF-fakturaer i Faktura-mappen
    pdf_priser = {}
    pdf_filer = sorted(glob.glob(os.path.join(FAKTURA_MAPPE, "*.pdf")))
    pdf_feil = []
    for filsti in pdf_filer:
        resultat = les_pris_fra_pdf(filsti)
        if resultat:
            år, måned_nr, pris = resultat
            pdf_priser[(år, måned_nr)] = pris
        else:
            pdf_feil.append(os.path.basename(filsti))

    if pdf_feil:
        print(f"  ! Kunne ikke lese pris fra: {', '.join(pdf_feil)}")

    # Meld fra om kilde per måned
    alle_nøkler = set(excel_priser) | set(pdf_priser)
    priser = {}
    for nøkkel in sorted(alle_nøkler):
        if nøkkel in pdf_priser:
            priser[nøkkel] = pdf_priser[nøkkel]
        else:
            priser[nøkkel] = excel_priser[nøkkel]

    print(f"  Strømpriser lastet: {len(pdf_priser)} fra PDF, "
          f"{len(alle_nøkler) - len(pdf_priser)} fra Excel")

    return priser


# ── Beboere ────────────────────────────────────────────────────────────────

def les_beboere():
    wb = openpyxl.load_workbook(EXCEL_FIL, data_only=True)
    ws = wb["Beboere"]
    beboere = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        nr, navn, hnummer = row[0], row[1], row[2]
        if nr and navn:
            beboere[int(nr)] = {
                "navn": str(navn).strip(),
                "hnummer": str(hnummer).strip() if hnummer else "",
            }
    return beboere


# ── CloudCharge CSV ────────────────────────────────────────────────────────

def norsk_float(verdi):
    if pd.isna(verdi):
        return 0.0
    if isinstance(verdi, (int, float)):
        return float(verdi)
    return float(str(verdi).replace("\xa0", "").replace(" ", "").replace(",", "."))


def les_csv_filer():
    filer = sorted(glob.glob(os.path.join(CLOUDCHARGE_MAPPE, "*.csv")))
    if not filer:
        raise FileNotFoundError(
            f"Ingen CSV-filer funnet i {CLOUDCHARGE_MAPPE}. "
            "Legg CloudCharge-fil(er) i mappen CloudCharge/ og prøv igjen."
        )

    print(f"  {len(filer)} CSV-fil(er) fra CloudCharge:")
    for f in filer:
        print(f"    {os.path.basename(f)}")

    deler = [pd.read_csv(f, sep=";", encoding="utf-8-sig", dtype=str) for f in filer]
    return pd.concat(deler, ignore_index=True)


# ── Summering ──────────────────────────────────────────────────────────────

def legg_til_summer(resultater):
    """Injiserer en Sum-rad etter hver gruppe av måneder per ladepunkt."""
    output = []
    i = 0
    while i < len(resultater):
        ladepunkt = resultater[i]["Ladepunkt"]
        gruppe = []
        while i < len(resultater) and resultater[i]["Ladepunkt"] == ladepunkt:
            gruppe.append(resultater[i])
            i += 1

        output.extend(gruppe)

        if len(gruppe) > 1:
            output.append({
                "Ladepunkt": ladepunkt,
                "Navn": gruppe[0]["Navn"],
                "Leilighet": gruppe[0]["Leilighet"],
                "År": "Sum",
                "Måned": None,
                "Forbruk (kWh)": round(sum(r["Forbruk (kWh)"] for r in gruppe), 2),
                "Strømpris inkl. 20% påslag (øre/kWh)": None,
                "Strømkost (kr)": round(sum(r["Strømkost (kr)"] for r in gruppe), 2),
                "_summary": True,
            })

    return output


# ── Hovedlogikk ────────────────────────────────────────────────────────────

def beregn(fra=None, til=None):
    if fra and til:
        print(f"Periode: {fra[0]}.{fra[1]:02d} – {til[0]}.{til[1]:02d}")
    print("Leser strømpriser (PDF + Excel)...")
    priser = les_strompriser()

    print("Leser beboerregister fra Excel...")
    beboere = les_beboere()

    print("Leser CloudCharge CSV...")
    data = les_csv_filer()

    data["Energi_kWh"] = data["Energi (kWh)"].apply(norsk_float)
    data["Sluttdato_dt"] = pd.to_datetime(data["Sluttdato"], format="%Y-%m-%d", errors="coerce")
    data["Uttak"] = pd.to_numeric(data["Uttak nummer"], errors="coerce")
    data["BeboerNr"] = data["Uttak"] + 100
    data = data.dropna(subset=["Sluttdato_dt", "Uttak"])

    # Sluttdato for månedstildeling — samsvarer med CloudCharges eksportformat
    # (CloudCharge inkluderer sesjoner i perioden basert på sluttdato).
    data["År"] = data["Sluttdato_dt"].dt.year.astype(int)
    data["Måned"] = data["Sluttdato_dt"].dt.month.astype(int)
    data["BeboerNr"] = data["BeboerNr"].astype(int)

    måneder_i_csv = sorted({(int(r["År"]), int(r["Måned"])) for _, r in data.iterrows()})
    if fra and til:
        måneder_i_csv = [m for m in måneder_i_csv if fra <= m <= til]
    gyldige_måneder = [m for m in måneder_i_csv if m in priser]
    utelatte_måneder = [m for m in måneder_i_csv if m not in priser]

    forbruk_data = data[data["Energi_kWh"] > 0]
    gruppert = forbruk_data.groupby(["BeboerNr", "År", "Måned"])["Energi_kWh"].sum()

    resultater = []

    for beboer_nr in sorted(beboere.keys()):
        beboer = beboere[beboer_nr]
        for (år, måned) in gyldige_måneder:
            forbruk = round(gruppert.get((beboer_nr, år, måned), 0.0), 2)
            pris_ore = round(priser[(år, måned)] * PRISPÅSLAG, 4)
            kostnad = round(forbruk * pris_ore / 100, 2)

            resultater.append({
                "Ladepunkt": beboer_nr,
                "Navn": beboer["navn"],
                "Leilighet": beboer["hnummer"],
                "År": år,
                "Måned": MÅNEDER[måned],
                "_sort": (beboer_nr, år, måned),
                "_summary": False,
                "Forbruk (kWh)": forbruk,
                "Strømpris inkl. 20% påslag (øre/kWh)": pris_ore,
                "Strømkost (kr)": kostnad,
            })

    if utelatte_måneder:
        utelatt_str = ", ".join(f"{MÅNEDER[m]} {å}" for å, m in utelatte_måneder)
        print(f"\n  Følgende måneder utelatt (ingen PDF-faktura eller Excel-pris): {utelatt_str}")

    if not resultater:
        print("\nIngen resultater å skrive. Sjekk at filer er lagt i riktige mapper.")
        return

    resultater.sort(key=lambda x: x["_sort"])
    for r in resultater:
        del r["_sort"]

    resultater = legg_til_summer(resultater)

    tidsstempel = datetime.now().strftime("%Y%m%d_%H%M")
    if fra and til:
        periode_str = f"{fra[0]}{fra[1]:02d}-{til[0]}{til[1]:02d}_"
    else:
        periode_str = ""
    output_fil = os.path.join(MAPPE, f"Fakturering_{periode_str}{tidsstempel}.xlsx")
    skriv_excel(resultater, output_fil)

    rader = [r for r in resultater if not r["_summary"]]
    print(f"\nTotalt forbruk : {sum(r['Forbruk (kWh)'] for r in rader):,.2f} kWh")
    print(f"Total kostnad  : kr {sum(r['Strømkost (kr)'] for r in rader):,.2f}")
    print(f"\nFerdig! Fil lagret: {os.path.basename(output_fil)}")


# ── Excel-output ───────────────────────────────────────────────────────────

def skriv_excel(resultater, filsti):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fakturering"

    kolonner = [
        "Ladepunkt", "Navn", "Leilighet", "År", "Måned",
        "Forbruk (kWh)", "Strømpris inkl. 20% påslag (øre/kWh)", "Strømkost (kr)",
    ]
    kolonnebredder = [12, 32, 12, 8, 14, 16, 32, 16]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    summary_font = Font(bold=True)
    summary_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    topp_kant = Border(top=Side(style="thin"))
    annenhver_fill = PatternFill(start_color="EEF3F9", end_color="EEF3F9", fill_type="solid")

    for col, navn in enumerate(kolonner, 1):
        celle = ws.cell(row=1, column=col, value=navn)
        celle.font = header_font
        celle.fill = header_fill
        celle.alignment = Alignment(horizontal="center")

    forrige_ladepunkt = None
    gruppe_teller = 0

    for rad_nr, rad in enumerate(resultater, 2):
        er_summary = rad.get("_summary", False)

        if rad["Ladepunkt"] != forrige_ladepunkt:
            forrige_ladepunkt = rad["Ladepunkt"]
            gruppe_teller += 1

        fyll = summary_fill if er_summary else (annenhver_fill if gruppe_teller % 2 == 0 else None)

        def sett(col, verdi, tallformat=None):
            c = ws.cell(row=rad_nr, column=col, value=verdi)
            if fyll:
                c.fill = fyll
            if tallformat:
                c.number_format = tallformat
            if er_summary:
                c.font = summary_font
                c.border = topp_kant
            return c

        sett(1, rad["Ladepunkt"])
        sett(2, rad["Navn"])
        sett(3, rad["Leilighet"])
        sett(4, rad["År"])
        sett(5, rad["Måned"])
        sett(6, rad["Forbruk (kWh)"], "#,##0.00")
        sett(7, rad["Strømpris inkl. 20% påslag (øre/kWh)"], "#,##0.00")
        sett(8, rad["Strømkost (kr)"], "#,##0.00")

    for col, bredde in enumerate(kolonnebredder, 1):
        ws.column_dimensions[get_column_letter(col)].width = bredde

    ws.freeze_panes = "A2"
    wb.save(filsti)


def parse_periode(args):
    """
    Leser valgfrie periodeargumenter på formen YYYY.MM YYYY.MM.
    Returnerer (fra_tuple, til_tuple) eller (None, None) om ingen er oppgitt.
    """
    if len(args) == 0:
        return None, None
    if len(args) != 2:
        print("Bruk: python beregn_lading.py [YYYY.MM YYYY.MM]")
        print("Eksempel: python beregn_lading.py 2026.01 2026.04")
        raise SystemExit(1)
    fra_str, til_str = args
    try:
        fra = tuple(int(x) for x in fra_str.split("."))
        til = tuple(int(x) for x in til_str.split("."))
        assert len(fra) == 2 and len(til) == 2
        assert 1 <= fra[1] <= 12 and 1 <= til[1] <= 12
        assert fra <= til
    except Exception:
        print(f"Ugyldig periode: '{fra_str}' '{til_str}'. Forventet format: YYYY.MM YYYY.MM")
        raise SystemExit(1)
    return fra, til


if __name__ == "__main__":
    import sys
    fra, til = parse_periode(sys.argv[1:])
    beregn(fra, til)
