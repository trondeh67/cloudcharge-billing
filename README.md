# CloudCharge Billing

Beregner strømkostnader for elbil-ladepunkter i sameie basert på ladedata fra CloudCharge og strømpris-PDF fra leverandør. Genererer Excel-rapport klar for videreformidling til regnskapsfører.

## Forutsetninger

- Python 3.9 eller nyere
- Tilgang til CloudCharge-portalen for nedlasting av ladedata
- PDF-fakturaer fra strømleverandør (Ustekveikja Energi AS)
- Excel-filen `Elbillading Strøm+Beboere.xlsx` med beboerregister og historiske strømpriser (ikke inkludert i repoet — inneholder persondata)

## Installasjon

```bash
pip install -r requirements.txt
```

## Bruk

### 1. Legg inn ladedata fra CloudCharge

Logg inn på [portal.cloudcharge.se](https://portal.cloudcharge.se/reports), velg rapporttype **Charge sessions**, sett ønsket tidsperiode og last ned CSV-filen. Legg filen i undermappen **`CloudCharge/`**.

> CloudCharge kan konfigureres til å sende månedlige CSV-rapporter automatisk til en e-postadresse.

### 2. Legg inn strømfaktura

Lagre PDF-fakturaen fra Ustekveikja Energi i undermappen **`Faktura/`**. Scriptet leser automatisk ut spotpris, strømstøtte og nettleie og beregner totalpris per kWh.

> Strømfaktura for en gitt måned er kjent først måneden etter. Måneder som mangler faktura utelates automatisk fra rapporten — scriptet melder fra om hvilke måneder som hoppes over.

### 3. Kjør scriptet

```bash
python beregn_lading.py
```

Scriptet kombinerer alle CSV-filer i `CloudCharge/` og alle PDF-er i `Faktura/` automatisk. Legg gjerne inn flere månedsfiler for et helt kvartal.

### 4. Hent output

En fil med navn `Fakturering_YYYYMMDD_HHMM.xlsx` opprettes i rotmappen. Denne sendes til regnskapsfører.

## Filstruktur

```
cloudcharge-billing/
├── beregn_lading.py              # Hovedscript
├── requirements.txt              # Python-avhengigheter
├── .gitignore
└── README.md

Ikke i versjonskontroll (legg til manuelt):
├── Elbillading Strøm+Beboere.xlsx      # Beboerregister + historiske priser
├── CloudCharge/                        # CSV-filer fra CloudCharge
│   └── Charge sessions *.csv
├── Faktura/                            # PDF-fakturaer fra strømleverandør
│   └── Faktura *.pdf
└── Fakturering_*.xlsx                  # Genererte rapporter
```

## Strømprisberegning fra PDF

Scriptet leser følgende felter fra Ustekveikja Energi-fakturaen og beregner totalpris per kWh:

| Felt i PDF | Brukes til |
|---|---|
| Strømpris — øre/kWh | Spotpris eks. MVA |
| Strømpris — kWh | Totalt forbruk (for strømstøtte-beregning) |
| Midlertidig strømstønad for borettslag — Nettobeløp | Strømstøtte i kr |
| Energiledd hverdag — øre/kWh | Nettleie-komponent |
| Elavgift — øre/kWh | Nettleie-komponent |

**Formel:**
```
spot_m_mva      = spotpris × 1,25
strømstøtte_øre = strømstøtte_kr / forbruk_kWh × 100
nettleie_m_mva  = (energiledd_hverdag + elavgift) × 1,25
totalt_pr_kWh   = spot_m_mva − strømstøtte_øre + nettleie_m_mva
```

For måneder uten PDF-faktura (historiske data) hentes prisen fra fanen **Strømpriser** i Excel-filen.

## Kolonnebeskrivelse — output Excel

| Kolonne | Beskrivelse |
|---|---|
| Ladepunkt | Nummer 101–116 (100 + uttaksnummer fra CloudCharge) |
| Navn | Beboernavn fra register |
| Leilighet | H-nummer (f.eks. H0209) |
| År | Årstall |
| Måned | Månedsnavn, eller **Sum** |
| Forbruk (kWh) | Forbruk for måneden, eller totalt for Sum-raden |
| Strømpris inkl. 20% påslag (øre/kWh) | Totalpris × 1,20, tom på Sum-rad |
| Strømkost (kr) | Forbruk × pris / 100, eller totalsum for Sum-raden |

## Månedstildeling ved sesjoner over månedsskiftet

En ladeøkt tildeles måneden den **avsluttet** (`Sluttdato`). Dette samsvarer med hvordan CloudCharge selv periodiserer sesjoner i eksportfilen: en sesjon er inkludert i perioden der sluttdatoen faller. Dermed blir summen i scriptet konsistent med hva som faktisk ligger i CSV-filen for en gitt periode.

## Prispåslag

Totalpris per kWh multipliseres med **20%** før kostnad beregnes. Bakgrunnen er at gjennomsnittlig månedspris brukes, men lading kan skje i dyrere timer på døgnet — påslaget dekker denne usikkerheten.

Påslaget er definert som konstanten `PRISPÅSLAG = 1.20` øverst i `beregn_lading.py`.

## Advarsler ved kjøring

| Situasjon | Hva skjer |
|---|---|
| Måned mangler både PDF-faktura og Excel-pris | Måneden utelates, melding i konsollen |
| PDF kan ikke leses/parses | Advarsel i konsollen, faller tilbake på Excel |
