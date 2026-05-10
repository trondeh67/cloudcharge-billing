# CloudCharge Billing — Køhnkvartalet Sameie

## Prosjekt
Python-script som beregner strømkostnader for elbil-ladepunkter i sameiet og genererer Excel-rapport til regnskapsfører. Fakturering skjer 4 ganger per år (kvartalsvis).

## Filstruktur
```
beregn_lading.py                  # Hovedscript
requirements.txt                  # Python-avhengigheter
Elbillading Strøm+Beboere.xlsx    # IKKE i git — persondata og strømpriser (Excel-fallback)
CloudCharge/                      # IKKE i git — CSV-filer fra CloudCharge
Faktura/                          # IKKE i git — PDF-fakturaer fra strømleverandør
Fakturering_*.xlsx                # IKKE i git — output til regnskapsfører
```

## Datakilder

### CloudCharge CSV ("Charge sessions"-rapport)
- Legges i undermappen `CloudCharge/`
- Format: semikolonseparert, norsk tallformat (komma desimal, mellomrom tusenskiller)
- Nøkkelkolonner: `Uttak nummer` (1–16), `Startdato` (YYYY-MM-DD), `Energi (kWh)`
- Én rad per ladeøkt — scriptet summerer per uttak per måned
- Månedstildeling basert på **`Startdato`** (ikke sluttdato): mesteparten av ladingen skjer tidlig i økten, og lange sesjoner der kabelen står i bilen over tid tilhører naturlig måneden de startet. Konsistent med tidligere PowerBI-rapport.
- Alle CSV-filer i mappen behandles — legg gjerne inn flere månedsfiler samtidig

### PDF-fakturaer (Ustekveikja Energi AS)
- Legges i undermappen `Faktura/`
- Scriptet leser automatisk alle PDF-er og beregner `Totalt pr/kWh` etter formelen:

```
spot_m_mva       = spotpris (øre/kWh eks. MVA) × 1,25
strømstøtte_øre  = strømstøtte_kr / forbruk_kWh × 100
nettleie_m_mva   = (energiledd_hverdag + elavgift) × 1,25
totalt_pr_kWh    = spot_m_mva − strømstøtte_øre + nettleie_m_mva
```

- Felter som leses fra PDF: `Strømpris` (spotpris + forbruk), `Midlertidig strømstønad for borettslag` (nettobeløp), `Energiledd hverdag` (øre/kWh), `Elavgift` (øre/kWh)
- Leverandør: Ustekveikja Energi AS — regex-mønstre er tilpasset dette faktura-formatet

### Excel: Elbillading Strøm+Beboere.xlsx (fallback for eldre måneder)
Fane **Strømpriser**:
- Kolonne 1: dato (første dag i måneden)
- Kolonne 11 (siste): `Totalt pr/kwh` — brukes for måneder uten PDF-faktura

Fane **Beboere**:
- Kolonne `Nr`: ladepunktnummer = 100 + `Uttak nummer` fra CloudCharge (uttak 4 → Nr 104)
- Kolonne `Navn`: beboernavn
- Kolonne `H-nummer`: leilighetsnummer (f.eks. H0209)
- 16 ladepunkter totalt (101–116); nr 110 og 116 tilhører samme beboer

## Output
Excel-fil `Fakturering_YYYYMMDD_HHMM.xlsx` med kolonner:
`Ladepunkt | Navn | Leilighet | År | Måned | Forbruk (kWh) | Strømpris inkl. 20% påslag (øre/kWh) | Strømkost (kr)`

Når et ladepunkt har forbruk i flere måneder, legges det til en **Sum-rad** (fet, blå bakgrunn) etter månedradene med totalt forbruk og total kostnad.

Måneder uten registrert strømpris (verken PDF eller Excel) **utelates helt** fra rapporten.

## Prispåslag
Spotprisen multipliseres med **1,20 (20%)** før utregning. Påslaget kompenserer for at gjennomsnittlig spotpris brukes, men faktisk lading kan skje i dyrere timer på døgnet. Konstanten `PRISPÅSLAG` øverst i scriptet kan justeres ved behov.

## Viktige hensyn
- Excel leses med `data_only=True` — formler evalueres ikke, kun cachet verdi leses. Excel-filen må ha vært åpnet og lagret i Excel for at verdiene skal være cachet.
- Siden strømfaktura kommer måneden etter, vil siste måned i CloudCharge-data typisk mangle pris og utelates automatisk.
- CloudCharge har ikke offentlig API — CSV lastes ned manuelt fra portal.cloudcharge.se eller mottas på e-post (månedlig automatisk rapport kan konfigureres i portalen).
- Kjør `pip install -r requirements.txt` ved første gangs oppsett
