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

> Strømfaktura for en gitt måned er kjent først måneden etter. Måneder som mangler pris utelates automatisk fra rapporten — scriptet melder fra om hvilke måneder som hoppes over.

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
| Måned | Månedsnavn og år (f.eks. April 2026), eller **Sum** |
| Forbruk (kWh) | Forbruk for måneden, eller totalt for Sum-raden |
| Strømpris inkl. 20% påslag (øre/kWh) | Spotpris fra Excel × 1,20, tom på Sum-rad |
| Strømkost (kr) | Forbruk × pris / 100, eller totalsum for Sum-raden |

## Månedstildeling ved sesjoner over månedsskiftet

En ladeøkt tildeles måneden den **startet** (`Startdato`), ikke måneden den avsluttet.

Begrunnelsen er todelt:
- Mesteparten av energien overføres tidlig i en ladeøkt — bilen lader raskt til den er full og trickle-lader deretter
- Lange sesjoner der kabelen blir stående i bilen over tid (f.eks. fra fredag til mandag) tilhører naturlig måneden ladingen begynte

Dette er konsistent med hvordan tallene tidligere ble beregnet i PowerBI.

## Prispåslag

Spotprisen fra Excel multipliseres med **20%** før kostnad beregnes. Bakgrunnen er at gjennomsnittlig månedspris brukes, men lading kan skje i dyrere timer på døgnet — påslaget dekker denne usikkerheten.

Påslaget er definert som konstanten `PRISPÅSLAG = 1.20` øverst i `beregn_lading.py` og kan justeres der ved behov.

Når et ladepunkt har forbruk i flere måneder, legges en **Sum-rad** til etter månedradene med totalt forbruk og total kostnad for perioden.

## Advarsler ved kjøring

| Situasjon | Hva skjer |
|---|---|
| Måned mangler strømpris i Excel | Måneden utelates fra rapporten, melding i konsollen |
| Uttaksnummer i CSV finnes ikke i beboerregisteret | Raden hoppes over, advarsel i konsollen |
