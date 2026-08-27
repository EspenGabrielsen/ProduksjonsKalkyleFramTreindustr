# Azure Go Live – ProduksjonsKalkyle

Dette dokumentet beskriver anbefalt produksjonssetting av ProduksjonsKalkyle i Azure.

## 1. Målarkitektur

Anbefalt oppsett:

- **Azure App Service** for web-applikasjonen
- **Azure Database for PostgreSQL Flexible Server** for datamodellen
- **Microsoft Entra ID** for autentisering
- eventuelt **Azure Blob Storage** senere hvis opplastede filer skal flyttes ut av databasen

## 2. Hva som deployes

Applikasjonen kjøres som container eller som Python-app med startup-kommando. Repoet inneholder:

- `Dockerfile`
- `startup.sh`
- `runtime.txt`

Startkommandoen i praksis er:

```bash
marimo run src/varekost_app.py --host 0.0.0.0 --port 8000 --headless
```

## 3. Azure-ressurser

Minimum:

1. **Resource Group**
2. **App Service Plan**
3. **Web App / App Service**
4. **Azure Database for PostgreSQL Flexible Server**
5. **Entra App Registration / Authentication-oppsett**

Senere ved behov:

6. **Application Insights**
7. **Storage Account / Blob Container**

## 4. App Settings

Sett følgende i **App Service → Configuration → Application settings**:

| Navn | Formål |
|------|--------|
| `WEBSITES_PORT` | Port Azure sender trafikk til |
| `PORT` | Port brukt av `startup.sh` |
| `PRODUKSJONSKALKYLE_TEST` | Skal være `false` i produksjon |
| `DATABASE_URL` | PostgreSQL-tilkobling |
| `ALLOWED_TENANT_ID` | Tillatt tenant-id |
| `MARIMO_BASE_URL` | Valgfri dersom appen ligger bak base path |

Eksempel:

```text
WEBSITES_PORT=8000
PORT=8000
PRODUKSJONSKALKYLE_TEST=false
DATABASE_URL=postgresql://user:password@host:5432/dbname?sslmode=require
ALLOWED_TENANT_ID=00000000-0000-0000-0000-000000000000
MARIMO_BASE_URL=/
```

## 5. Authentication

I **App Service → Authentication**:

1. Legg til **Microsoft** som identity provider
2. Bruk en **single-tenant** app registration
3. Sett **Require authentication**
4. Sett redirect/login for uautentiserte kall

I Entra ID:

- **Supported account types** = kun egen organisasjon

I applikasjonen valideres tenant også via `ALLOWED_TENANT_ID`.

## 6. PostgreSQL-oppsett

Anbefalinger:

- bruk **Azure Database for PostgreSQL Flexible Server**
- krev TLS (`sslmode=require` eller strengere)
- opprett egen database for løsningen
- opprett minst én applikasjonsbruker
- vurder egen read-only-bruker senere for rapportering/API-konsumenter

## 7. Første produksjonssetting

Anbefalt rekkefølge:

1. Opprett PostgreSQL-server og database
2. Sett `DATABASE_URL` i App Service
3. Deploy appen
4. Kjør database-initialisering
5. Migrer eksisterende SQLite-data til PostgreSQL
6. Verifiser nøkkeltabeller
7. Verifiser beregninger for et utvalg produkter
8. Slå på tilgang for sluttbrukere

## 8. Verifisering etter deploy

Verifiser minst følgende:

### Innlogging
- brukere blir sendt til Entra login
- kun brukere i riktig tenant slipper inn

### Database
- appen starter med PostgreSQL-backend
- tabeller opprettes korrekt
- import fra Excel fungerer
- historikk / endringslogg fungerer

### Beregning
- beregning for kjente produkter gir forventet resultat
- simulering fungerer
- eksport til PDF/Excel fungerer

## 9. Go-live kriterier

Løsningen bør ikke regnes som produksjonsklar før:

- PostgreSQL-sporet er testet ende-til-ende
- datamigrering er dokumentert og verifisert
- Azure auth er validert
- minst ett sett med faglige referanseberegninger er bekreftet

## 10. Drift etter go-live

Siden modellen endres sjelden og kontrollert, anbefales denne arbeidsformen:

1. ta backup før månedlige modellendringer
2. gjennomfør endring/import i kontrollert vindu
3. verifiser beregninger
4. dokumenter endringer

Dette passer godt med applikasjonens karakter som fagmodell og beregningsmotor, ikke høyfrekvent transaksjonssystem.