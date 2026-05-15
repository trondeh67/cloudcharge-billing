# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# CloudCharge Billing — Køhnkvartalet Sameie

## Prosjekt
Python-script som beregner strømkostnader for elbil-ladepunkter i sameiet og genererer Excel-rapport til regnskapsfører. Fakturering skjer 4 ganger per år (kvartalsvis).

## Oppsett
```
pip install -r requirements.txt
```
Ingen tester, ingen linting-konfig. Verifisering gjøres ved å kjøre scriptet mot testdata.

## Filstruktur
```
beregn_lading.py                  # Hovedscript
requirements.txt                  # Python-avhengigheter
Elbillading Strøm+Beboere.xlsx    # IKKE i git — persondata og strømpriser (Excel-fallback)
CloudCharge/                      # IKKE i git — CSV-filer fra CloudCharge
Faktura/                          # IKKE i git — PDF-fakturaer fra strømleverandør
Fakturering_*.xlsx                # IKKE i git — output til regnskapsfører
```

## Kodearkitektur

Alt ligger i `beregn_lading.py`. Flyten er:

```
parse_args()
  └─ beregn()
       ├─ les_strompriser()
       │    ├─ les_pris_fra_pdf()   ← regex mot Ustekveikja Energi-format
       │    └─ Excel-fallback       ← openpyxl, fane "Strømpriser", kolonne 11
       ├─ les_ladepunkt_nokkel()    ← Ladepunkter.csv (ladepunkt_id → hnummer)
       ├─ les_beboerliste()         ← nyeste *-alle seksjoner.csv (hnummer → Eier)
       ├─ koblet_beboere()          ← slår sammen de to til {lp_id: {navn, hnummer}}
       ├─ les_csv_filer()           ← pandas concat av alle *.csv i CloudCharge-mappen
       ├─ legg_til_summer()         ← injiserer Sum-rad per ladepunkt
       └─ skriv_excel()             ← fane "Fakturering" + fane "Oppsummert"
```

Nøkkelkonstanter øverst i scriptet:
- `PRISPÅSLAG = 1.20` — 20% påslag på spotpris
- `FORVENTET_ANLEGGSREFERANSE = "Strøm Fellesareal"` — filtrerer bort feil PDF-fakturaer

## Datakilder

### CloudCharge CSV ("Charge sessions"-rapport)
- Legges i undermappen `CloudCharge/`
- Format: semikolonseparert, norsk tallformat (komma desimal, mellomrom tusenskiller)
- Nøkkelkolonner: `Uttak nummer` (1–16), `Sluttdato` (YYYY-MM-DD), `Energi (kWh)`
- Én rad per ladeøkt — scriptet summerer per uttak per måned
- Månedstildeling basert på **`Sluttdato`**: CloudCharge inkluderer sesjoner i en eksportperiode basert på sluttdato, så sluttdato gir konsistent samsvar mellom CSV-fil og hvilken måned sesjonen telles i.
- Alle CSV-filer i mappen behandles — legg gjerne inn flere månedsfiler samtidig

### PDF-fakturaer (Ustekveikja Energi AS)
- Legges i undermappen `Faktura/`
- Scriptet leser automatisk alle PDF-er og beregner `Totalt pr/kWh` etter formelen:

```
totalt_pr_kWh = Total fakturabeløp inkl. MVA (kr) / Strømpris forbruk (kWh) × 100
```

- Felter som leses fra PDF: `Strømpris` (periode + forbruk kWh), `Total sum` (fakturabeløp inkl. MVA)
- Leverandør: Ustekveikja Energi AS — regex-mønstre er tilpasset dette faktura-formatet
- Fakturaer med feil Anleggsreferanse avvises automatisk med varsel

### Ladepunkter.csv
Ligger i samme mappe som scriptet. Kobler ladepunkt-ID til leilighetsnummer:
- Kolonne `ladepunkt_id`: 101–116
- Kolonne `hnummer`: leilighetsnummer (f.eks. H0209)
- 16 ladepunkter totalt; nr 110 og 116 peker på samme H-nummer (samme beboer)

### `<YYYY-MM-DD>-alle seksjoner.csv` (beboerregister fra styret.com)
Semikolonseparert, Windows-1252-kodet, med `sep=;` som første linje (Excel-eksport).
Scriptet finner automatisk filen med nyeste dato i filnavnet (alfabetisk sortering).
- Kobling mot `Ladepunkter.csv` via kolonne `H-nummer`
- Beboernavn hentes fra kolonne `Eier` og brukes slik den står (kan inneholde flere navn separert med komma)

### Excel: Elbillading Strøm+Beboere.xlsx (kun for strømpriser)
Fane **Strømpriser** (fallback for måneder uten PDF-faktura):
- Kolonne 1: dato (første dag i måneden)
- Kolonne 11 (siste): `Totalt pr/kwh`

## Kjøring

```
python beregn_lading.py                                         # alle måneder, standardmapper
python beregn_lading.py 2026.01 2026.04                        # kun januar–april 2026
python beregn_lading.py 2026.01 2026.04 -F <sti> -C <sti>     # egne mappestier
```

| Argument | Beskrivelse |
|---|---|
| `YYYY.MM YYYY.MM` | Valgfri periode. Begge må oppgis eller ingen. |
| `-F MAPPE` | Sti til PDF-fakturamappe (standard: `Faktura/` ved siden av scriptet) |
| `-C MAPPE` | Sti til CloudCharge CSV-mappe (standard: `CloudCharge/` ved siden av scriptet) |

Ugyldig mappe eller feil periodeformat gir feilmelding og avslutter. Bruk `--help` for full oversikt.

## Output
Excel-fil med navn:
- `Fakturering_YYYYMMDD_HHMM.xlsx` — uten periodeargument
- `Fakturering_YYYYMM-YYYYMM_YYYYMMDD_HHMM.xlsx` — med periodeargument

### Fane 1: Fakturering
Månedlig detaljert oversikt. Kolonner:
`Ladepunkt | Navn | Leilighet | År | Måned | Forbruk (kWh) | Strømpris inkl. 20% påslag (øre/kWh) | Strømkost (kr)`

Alle 16 ladepunkter vises for alle gyldige måneder, også de med null forbruk. Ladepunkter med forbruk i flere måneder får en **Sum-rad** (fet, blå bakgrunn) etter månedradene.

### Fane 2: Oppsummert
Fakturaunderlag til forretningsfører — aggregerte totaler per ladepunkt for hele perioden. Kolonner:
`Ladepunkt | Navn | Leilighet | Forbruk (kWh) | Strømkost (kr)`

Alle 16 ladepunkter vises inkludert de med null forbruk. Øverst: tittelrad med perioden. Nederst: totallinje.

Måneder uten registrert strømpris (verken PDF eller Excel) **utelates helt** fra rapporten.

## Prispåslag
Spotprisen multipliseres med **1,20 (20%)** før utregning. Påslaget kompenserer for at gjennomsnittlig spotpris brukes, men faktisk lading kan skje i dyrere timer på døgnet. Konstanten `PRISPÅSLAG` øverst i scriptet kan justeres ved behov.

## Viktige hensyn
- Excel leses med `data_only=True` — formler evalueres ikke, kun cachet verdi leses. Excel-filen må ha vært åpnet og lagret i Excel for at verdiene skal være cachet.
- Siden strømfaktura kommer måneden etter, vil siste måned i CloudCharge-data typisk mangle pris og utelates automatisk.
- CloudCharge har ikke offentlig API — CSV lastes ned manuelt fra portal.cloudcharge.se eller mottas på e-post (månedlig automatisk rapport kan konfigureres i portalen).
