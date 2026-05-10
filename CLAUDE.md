# CloudCharge Billing — Køhnkvartalet Sameie

## Prosjekt
Python-script som beregner strømkostnader for elbil-ladepunkter i sameiet og genererer Excel-rapport til regnskapsfører. Fakturering skjer 4 ganger per år (kvartalsvis).

## Filstruktur
```
beregn_lading.py                  # Hovedscript
requirements.txt                  # Python-avhengigheter
Elbillading Strøm+Beboere.xlsx    # IKKE i git — persondata og strømpriser
*.csv                             # IKKE i git — CloudCharge rådata
Fakturering_*.xlsx                # IKKE i git — output til regnskapsfører
```

## Datakilder

### CloudCharge CSV ("Charge sessions"-rapport)
- Format: semikolonseparert, norsk tallformat (komma desimal, mellomrom tusenskiller)
- Nøkkelkolonner: `Uttak nummer` (1–16), `Sluttdato` (YYYY-MM-DD), `Energi (kWh)`
- Én rad per ladeøkt — scriptet summerer per uttak per måned
- Månedstildeling basert på `Sluttdato`
- Filer legges i prosjektmappen før kjøring — alle CSV-filer i mappen behandles

### Excel: Elbillading Strøm+Beboere.xlsx
Fane **Strømpriser**:
- Kolonne 1: dato (første dag i måneden)
- Kolonne 11 (siste): `Totalt pr/kwh` — pris i øre inkl. nettleie og MVA, fratrukket strømstøtte

Fane **Beboere**:
- Kolonne `Nr`: ladepunktnummer = 100 + `Uttak nummer` fra CloudCharge (uttak 4 → Nr 104)
- Kolonne `Navn`: beboernavn
- Kolonne `H-nummer`: leilighetsnummer (f.eks. H0209)
- 16 ladepunkter totalt (101–116); nr 110 og 116 tilhører samme beboer

## Output
Excel-fil `Fakturering_YYYYMMDD_HHMM.xlsx` med kolonner:
`Ladepunkt | Navn | Leilighet | Måned | Forbruk (kWh) | Strømpris (øre/kWh) | Strømkost (kr)`

## Viktige hensyn
- Excel leses med `data_only=True` — formler evalueres ikke, kun cachet verdi leses. Excel-filen må ha vært åpnet og lagret i Excel for at verdiene skal være cachet.
- Strømpris for inneværende måned legges inn manuelt i Excel etter at strømfaktura er godkjent i styret.com
- CloudCharge har ikke offentlig API — CSV lastes ned manuelt fra portal.cloudcharge.se eller mottas på e-post
- Kjør `pip install -r requirements.txt` ved første gangs oppsett
