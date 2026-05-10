#!/usr/bin/env python3
"""
Elbillading - strømkostnad beregning, Køhnkvartalet Sameie

Leser CloudCharge CSV-filer og strømpriser fra Excel,
beregner kostnad per ladepunkt per måned og skriver
ut Excel-fil til regnskapsfører.

Bruk:
    python beregn_lading.py

Legg CloudCharge CSV-fil(er) i samme mappe som scriptet før kjøring.
Måneder uten registrert strømpris utelates automatisk fra rapporten.
"""

import os
import glob
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MAPPE = os.path.dirname(os.path.abspath(__file__))
EXCEL_FIL = os.path.join(MAPPE, "Elbillading Strøm+Beboere.xlsx")

MÅNEDER = {
    1: "Januar", 2: "Februar", 3: "Mars", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

# Påslag på spotpris for å dekke usikkerhet rundt ladetidspunkt på døgnet
PRISPÅSLAG = 1.20


def les_strompriser():
    wb = openpyxl.load_workbook(EXCEL_FIL, data_only=True)
    ws = wb["Strømpriser"]
    priser = {}
    for row in ws.iter_rows(values_only=True):
        dato, pris = row[0], row[10]  # første og siste kolonne
        if isinstance(dato, datetime) and isinstance(pris, (int, float)):
            priser[(dato.year, dato.month)] = pris
    return priser


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


def norsk_float(verdi):
    if pd.isna(verdi):
        return 0.0
    if isinstance(verdi, (int, float)):
        return float(verdi)
    return float(str(verdi).replace("\xa0", "").replace(" ", "").replace(",", "."))


def les_csv_filer():
    filer = glob.glob(os.path.join(MAPPE, "*.csv"))
    if not filer:
        raise FileNotFoundError("Ingen CSV-filer funnet i mappen. Legg CloudCharge-fil(er) i mappen og prøv igjen.")

    print(f"Fant {len(filer)} CSV-fil(er):")
    for f in filer:
        print(f"  {os.path.basename(f)}")

    deler = []
    for filsti in filer:
        df = pd.read_csv(filsti, sep=";", encoding="utf-8-sig", dtype=str)
        deler.append(df)

    return pd.concat(deler, ignore_index=True)


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
            total_kwh = round(sum(r["Forbruk (kWh)"] for r in gruppe), 2)
            total_kr = round(sum(r["Strømkost (kr)"] for r in gruppe), 2)
            output.append({
                "Ladepunkt": ladepunkt,
                "Navn": gruppe[0]["Navn"],
                "Leilighet": gruppe[0]["Leilighet"],
                "Måned": "Sum",
                "Forbruk (kWh)": total_kwh,
                "Strømpris (øre/kWh)": None,
                "Strømkost (kr)": total_kr,
                "_summary": True,
            })

    return output


def beregn():
    print("Leser strømpriser fra Excel...")
    priser = les_strompriser()

    print("Leser beboerregister fra Excel...")
    beboere = les_beboere()

    print("Leser CloudCharge CSV...")
    data = les_csv_filer()

    # Konverter nøkkelkolonner
    data["Energi_kWh"] = data["Energi (kWh)"].apply(norsk_float)
    data["Sluttdato_dt"] = pd.to_datetime(data["Sluttdato"], format="%Y-%m-%d", errors="coerce")
    data["Uttak"] = pd.to_numeric(data["Uttak nummer"], errors="coerce")
    data["BeboerNr"] = data["Uttak"] + 100

    # Behold bare rader med gyldig dato, uttak og positivt forbruk
    data = data.dropna(subset=["Sluttdato_dt", "Uttak"])
    data = data[data["Energi_kWh"] > 0]

    data["År"] = data["Sluttdato_dt"].dt.year.astype(int)
    data["Måned"] = data["Sluttdato_dt"].dt.month.astype(int)
    data["BeboerNr"] = data["BeboerNr"].astype(int)

    # Summer forbruk per beboer per måned
    gruppert = (
        data.groupby(["BeboerNr", "År", "Måned"])["Energi_kWh"]
        .sum()
        .reset_index()
    )

    resultater = []
    advarsler_pris = set()
    advarsler_beboer = set()

    for _, rad in gruppert.iterrows():
        beboer_nr = int(rad["BeboerNr"])
        år = int(rad["År"])
        måned = int(rad["Måned"])
        forbruk = round(rad["Energi_kWh"], 2)

        if beboer_nr not in beboere:
            advarsler_beboer.add(f"Ukjent ladepunkt {beboer_nr} (CloudCharge uttak {beboer_nr - 100}) — mangler i beboerregisteret")
            continue

        pris_nøkkel = (år, måned)
        if pris_nøkkel not in priser:
            # Utelat måneder uten kjent strømpris
            advarsler_pris.add(f"{MÅNEDER[måned]} {år}")
            continue

        beboer = beboere[beboer_nr]
        pris_ore = round(priser[pris_nøkkel] * PRISPÅSLAG, 4)
        kostnad = round(forbruk * pris_ore / 100, 2)

        resultater.append({
            "Ladepunkt": beboer_nr,
            "Navn": beboer["navn"],
            "Leilighet": beboer["hnummer"],
            "Måned": f"{MÅNEDER[måned]} {år}",
            "_sort": (beboer_nr, år, måned),
            "_summary": False,
            "Forbruk (kWh)": forbruk,
            "Strømpris (øre/kWh)": pris_ore,
            "Strømkost (kr)": kostnad,
        })

    if advarsler_pris:
        print(f"\n  Følgende måneder er utelatt (strømpris ikke registrert): {', '.join(sorted(advarsler_pris))}")
    if advarsler_beboer:
        print("\nAdvarsler:")
        for a in sorted(advarsler_beboer):
            print(f"  ! {a}")

    if not resultater:
        print("\nIngen resultater å skrive. Sjekk at CSV-filer og strømpriser er oppdatert.")
        return

    resultater.sort(key=lambda x: x["_sort"])
    for r in resultater:
        del r["_sort"]

    resultater = legg_til_summer(resultater)

    tidsstempel = datetime.now().strftime("%Y%m%d_%H%M")
    output_fil = os.path.join(MAPPE, f"Fakturering_{tidsstempel}.xlsx")
    skriv_excel(resultater, output_fil)

    rader_uten_sum = [r for r in resultater if not r["_summary"]]
    total_kwh = round(sum(r["Forbruk (kWh)"] for r in rader_uten_sum), 2)
    total_kr = round(sum(r["Strømkost (kr)"] for r in rader_uten_sum), 2)
    print(f"\nTotalt forbruk : {total_kwh:,.2f} kWh")
    print(f"Total kostnad  : kr {total_kr:,.2f}")
    print(f"\nFerdig! Fil lagret: {os.path.basename(output_fil)}")


def skriv_excel(resultater, filsti):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fakturering"

    kolonner = [
        "Ladepunkt", "Navn", "Leilighet", "Måned",
        "Forbruk (kWh)", "Strømpris inkl. 20% (øre/kWh)", "Strømkost (kr)",
    ]
    kolonnebredder = [12, 32, 12, 18, 16, 26, 16]

    # Stiler
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

        if er_summary:
            fyll = summary_fill
        else:
            fyll = annenhver_fill if gruppe_teller % 2 == 0 else None

        def sett(col, verdi, tallformat=None, bold=False):
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
        sett(4, rad["Måned"])
        sett(5, rad["Forbruk (kWh)"], "#,##0.00")
        sett(6, rad["Strømpris (øre/kWh)"], "#,##0.00")
        sett(7, rad["Strømkost (kr)"], "#,##0.00")

    for col, bredde in enumerate(kolonnebredder, 1):
        ws.column_dimensions[get_column_letter(col)].width = bredde

    ws.freeze_panes = "A2"
    wb.save(filsti)


if __name__ == "__main__":
    beregn()
