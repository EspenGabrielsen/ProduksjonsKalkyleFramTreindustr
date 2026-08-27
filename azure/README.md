# Azure App Service-oppsett for ProduksjonsKalkyle

Dette prosjektet er ment å kjøres på Azure App Service med Microsoft Entra ID som innlogging og PostgreSQL som produksjonsdatabase.

SQLite kan brukes lokalt og som midlertidig fallback under utvikling, men anbefalt produksjonsspor er:

- **Azure App Service** for applikasjonen
- **Azure Database for PostgreSQL Flexible Server** for datamodellen
- **Microsoft Entra ID / EasyAuth** for autentisering

## 1. Obligatoriske app settings

Sett følgende under **App Service → Configuration → Application settings**:

| Navn | Eksempelverdi | Formål |
|------|---------------|--------|
| `WEBSITES_PORT` | `8000` | Port Azure sender trafikk til |
| `PORT` | `8000` | Brukes av `startup.sh` |
| `PRODUKSJONSKALKYLE_TEST` | `false` | Sikrer at appen ikke bruker testdatabase |
| `DATABASE_URL` | `postgresql://user:pass@host:5432/db?sslmode=require` | PostgreSQL-tilkobling |
| `ALLOWED_TENANT_ID` | `00000000-0000-0000-0000-000000000000` | Ekstra validering i kode mot tillatt tenant |
| `MARIMO_BASE_URL` | `/` eller `/produksjonskalkyle` | Valgfritt dersom appen ligger bak en base path |

Hvis opplastede filer senere flyttes ut av databasen, legg også til:

| Navn | Formål |
|------|--------|
| `AZURE_STORAGE_CONNECTION_STRING` | Tilkobling til Blob Storage |
| `AZURE_STORAGE_CONTAINER` | Container for opplastede Excel-filer |

## 2. Startup command

Hvis du **ikke** bruker Docker-deploy, sett Startup Command til:

```bash
./startup.sh
```

Alternativt direkte:

```bash
marimo run src/varekost_app.py --host 0.0.0.0 --port 8000 --headless
```

## 3. Authentication / EasyAuth

I **App Service → Authentication**:

1. Legg til **Microsoft** som identity provider
2. Velg en **single-tenant** app registration i Entra ID
3. Sett **Require authentication**
4. Sett **Unauthenticated requests** til redirect/login

Dette gjør at bare innloggede brukere får tilgang.

## 4. Begrens til kun egen tenant

I Entra-appregistreringen:

- **Supported account types** = `Accounts in this organizational directory only`

I tillegg validerer appen tenant-id i kode via `ALLOWED_TENANT_ID`.

## 5. Valgfri ytterligere begrensning til grupper

Hvis bare enkelte brukere eller grupper skal få tilgang:

1. Gå til **Enterprise applications**
2. Velg appen
3. Sett **Assignment required = Yes**
4. Tildel riktige brukere/grupper

## 6. App Service-innstillinger

Anbefalte innstillinger:

- **Web Sockets** = On
- **Always On** = On
- **HTTP version** = 2.0

## 7. Docker-deploy

Dette repoet inneholder:

- `Dockerfile`
- `startup.sh`
- `runtime.txt`

Du kan derfor deploye enten:

1. som vanlig Python App Service med Oryx + `startup.sh`, eller
2. som container-basert App Service via `Dockerfile`

## 8. Databasevalg

Applikasjonen velger database-backend slik:

- hvis `DATABASE_URL` er satt: PostgreSQL
- ellers: SQLite

For Azure-produksjon er PostgreSQL anbefalt og planlagt som primær backend.

## 9. Nåværende status for PostgreSQL-sporet

Repoet har allerede:

- `psycopg[binary]` i `requirements.txt`
- støtte for `DATABASE_URL`
- PostgreSQL-tilkobling i `DataRepo`
- PostgreSQL-skjema i `src/data_repo.py`

Det gjenstår fortsatt arbeid før PostgreSQL-sporet kan regnes som ferdig:

- rydde SQLite-spesifikke rester i `src/data_repo.py`
- rydde SQLite-spesifikke rester i `src/excel_bridge.py`
- kjøre ende-til-ende validering mot ekte PostgreSQL
- dokumentere og gjennomføre migrering fra eksisterende SQLite-data

Se også:

- `docs/Azure_GoLive.md`
- `docs/PostgreSQL_Ferdigstillelse.md`
- `docs/API_Strategi.md`

## 10. Integrasjoner og API

Det anbefales at andre systemer på sikt henter **beregningsresultater** via API, fremfor å lese rå BOM/routing-tabeller direkte.

PostgreSQL er derfor valgt som anbefalt produksjonsdatabase ikke først og fremst på grunn av datamengde, men fordi løsningen skal være en ryddig Azure-tjeneste og kunne videreutvikles med API og integrasjoner.