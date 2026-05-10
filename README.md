# CloudCharge Billing

Beregner strømkostnader for elbil-ladepunkter i sameie basert på ladedata fra CloudCharge og strømpriser fra Excel. Genererer Excel-rapport klar for videreformidling til regnskapsfører.

## Forutsetninger

- Python 3.9 eller nyere
- Tilgang til CloudCharge-portalen for nedlasting av ladedata
- Excel-filen `Elbillading Strøm+Beboere.xlsx` med strømpriser og beboerregister (ikke inkludert i repoet — inneholder persondata)

## Installasjon

```bash
pip install -r requirements.txt
```

## Bruk

### 1. Hent ladedata fra CloudCharge

Logg inn på [portal.cloudcharge.se](https://portal.cloudcharge.se/reports), velg rapporttype **Charge sessions**, sett ønsket tidsperiode og last ned CSV-filen. Legg filen i prosjektmappen.

> CloudCharge kan også konfigureres til å sende månedlige CSV-rapporter automatisk til en e-postadresse.

### 2. Oppdater strømpriser

Åpne `Elbillading Strøm+Beboere.xlsx`, gå til fanen **Strømpriser** og legg inn pris (øre/kWh) for aktuelle måneder. Prisen hentes fra strømfakturaen (kolonne: `Totalt pr/kwh`).

Lagre og lukk Excel-filen før kjøring av scriptet.

### 3. Kjør scriptet

```bash
python beregn_lading.py
```

Scriptet leser automatisk alle CSV-filer i mappen. Flere CSV-filer kan legges inn samtidig (f.eks. tre månedsfiler for et kvartal).

### 4. Hent output

En fil med navn `Fakturering_YYYYMMDD_HHMM.xlsx` opprettes i mappen. Denne sendes til regnskapsfører.

## Filstruktur

```
cloudcharge-billing/
├── beregn_lading.py              # Hovedscript
├── requirements.txt              # Python-avhengigheter
├── .gitignore
└── README.md

Ikke i versjonskontroll (legg til manuelt):
├── Elbillading Strøm+Beboere.xlsx   # Strømpriser og beboerregister
├── *.csv                            # CloudCharge ladedata
└── Fakturering_*.xlsx               # Genererte rapporter
```

## Kolonnebeskrivelse — output Excel

| Kolonne | Beskrivelse |
|---|---|
| Ladepunkt | Nummer 101–116 (100 + uttaksnummer fra CloudCharge) |
| Navn | Beboernavn fra register |
| Leilighet | H-nummer (f.eks. H0209) |
| Måned | Månedsnavn og år (f.eks. Mai 2026) |
| Forbruk (kWh) | Totalt forbruk i perioden |
| Strømpris (øre/kWh) | Pris hentet fra Excel for aktuell måned |
| Strømkost (kr) | Forbruk × pris / 100 |

## Advarsler

Scriptet varsler ved kjøring hvis:
- En måned i CSV-dataene mangler strømpris i Excel
- Et uttaksnummer i CSV ikke finnes i beboerregisteret
