# ProduksjonsKalkyle

ProduksjonsKalkyle er en Python-basert applikasjon for å forvalte en produksjonsmodell med blant annet:

- produktregister
- BOM (stykklistestruktur)
- routing / operasjoner / arbeidssentre
- kostsatser
- biproduktregler
- produksjonsscenarier

Modellen brukes til å beregne produksjonskost og støtte simuleringer for trelastindustri / høvleri.

## Hovedformål

Applikasjonen brukes til å:

1. vedlikeholde innsatsfaktorer i modellen
2. importere og eksportere modellgrunnlag via Excel
3. beregne produksjonskost per produkt
4. simulere endringer i kostdrivere og produksjonsforutsetninger
5. generere rapporter i PDF, Excel og JSON

Innsatsfaktorer i modellen oppdateres kontrollert og sjelden. Dette er derfor ikke en høyfrekvent transaksjonsapplikasjon, men et fagverktøy for styrt modellforvaltning og beregning.

## Anbefalt målarkitektur

Prosjektet bør driftes i Azure med:

- **Azure App Service** for web-applikasjonen
- **Azure Database for PostgreSQL Flexible Server** for datamodellen
- **Microsoft Entra ID / EasyAuth** for innlogging til web-UI
- på sikt et **API-lag** for uthenting av beregningsresultater til andre systemer

Selv om datamengden er relativt liten, er PostgreSQL anbefalt som produksjonsbackend fordi løsningen forventes å kunne eksponere **beregninger** til andre systemer senere.

## Status i repoet

Følgende er allerede på plass:

- Marimo-basert web-app i `src/varekost_app.py`
- Docker-oppsett via `Dockerfile` og `startup.sh`
- PostgreSQL-driver via `psycopg[binary]`
- backend-valg via `DATABASE_URL`
- påbegynt PostgreSQL-støtte i `src/data_repo.py`
- Azure-dokumentasjon i `azure/README.md`

Følgende gjenstår før repoet er fullt klart for produksjon med PostgreSQL i Azure:

- fullføre PostgreSQL-kompatibilitet i `src/data_repo.py`
- fullføre PostgreSQL-kompatibilitet i `src/excel_bridge.py`
- lage migrering fra eksisterende SQLite-data til PostgreSQL
- teste hele arbeidsflyten ende-til-ende mot ekte PostgreSQL
- oppdatere lanseringsdokumentasjonen til PostgreSQL som primærspor

## Viktige dokumenter

- `azure/README.md` – konkret Azure-lanseringsveiledning
- `docs/Azure_GoLive.md` – runbook for produksjonssetting i Azure
- `docs/PostgreSQL_Ferdigstillelse.md` – teknisk sjekkliste for å fullføre PostgreSQL-sporet
- `docs/API_Strategi.md` – anbefalt retning for uthenting av beregningsresultater
- `docs/Produksjonsmodell_Dokumentasjon.md` – teknisk domenedokumentasjon
- `docs/Brukermanual_Produksjonsmodell.md` – brukerdokumentasjon

## Lokal kjøring

Installer avhengigheter:

```bash
pip install -r requirements.txt
```

Start appen lokalt:

```bash
marimo run src/varekost_app.py
```

Start appen via startup-scriptet:

```bash
./startup.sh
```

## Database-backend

Applikasjonen velger backend slik:

- hvis `DATABASE_URL` er satt: PostgreSQL
- ellers: SQLite

Målet for produksjon er PostgreSQL i Azure.

## Fremtidig API-retning

Når andre systemer skal hente data, bør de ikke nødvendigvis lese råtabeller direkte. Anbefalt retning er å eksponere **beregningsresultater** via API, for eksempel:

- kost per produkt
- kostnedbrytning per produkt
- scenario-beregninger
- simuleringsresultater

Dette gjør at andre systemer bruker samme beregningslogikk som GUI-et, uten å måtte forstå og reimplementere BOM-, routing- og kostreglene selv.