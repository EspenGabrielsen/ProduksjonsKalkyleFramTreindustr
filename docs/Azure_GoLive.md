# Azure Go Live – ProduksjonsKalkyle

Dette dokumentet beskriver anbefalt produksjonssetting av ProduksjonsKalkyle i Azure.

## 1. Målarkitektur

Anbefalt oppsett:

- **Azure App Service** for web-applikasjonen
- **Azure Database for PostgreSQL Flexible Server** for datamodellen
- **Microsoft Entra ID / App Service Authentication** for brukerinnlogging
- **Managed Identity / Microsoft Entra-autentisering** som foretrukket produksjonsspor mellom App Service og PostgreSQL
- eventuelt **Azure Blob Storage** senere hvis opplastede filer skal flyttes ut av databasen

PostgreSQL og App Service-containeren valideres nå automatisk i GitHub Actions på `azure-readiness`.

## 2. Hva som deployes

Applikasjonen kjøres som container. Repoet inneholder:

- `Dockerfile`
- `startup.sh`
- `runtime.txt`

Startkommandoen i praksis er:

```bash
marimo run src/varekost_app.py --host 0.0.0.0 --port 8000 --headless
```

GitHub Actions bygger det faktiske Docker-imaget på hver endring i Azure-branchen. App Service skal konfigureres med `WEBSITES_PORT=8000` fordi containeren lytter på port 8000.

## 3. Azure-ressurser

Minimum:

1. **Resource Group**
2. **App Service Plan**
3. **Web App / App Service**
4. **Azure Database for PostgreSQL Flexible Server**
5. **Entra App Registration / App Service Authentication-oppsett**
6. **Managed Identity** på App Service dersom passordløs databaseautentisering brukes

Senere ved behov:

7. **Application Insights**
8. **Storage Account / Blob Container**
9. **Azure Container Registry** dersom containerdeploy skal gå via eget registry

## 4. App Settings

Følgende innstillinger er relevante:

| Navn | Formål |
|------|--------|
| `WEBSITES_PORT` | Port Azure sender HTTP-trafikk til; settes til `8000` |
| `PORT` | Port brukt av `startup.sh`; settes til `8000` |
| `PRODUKSJONSKALKYLE_TEST` | Skal være `false` i produksjon |
| `DATABASE_URL` | Brukes av dagens PostgreSQL-tilkobling og i testmiljø |
| `ALLOWED_TENANT_ID` | Tillatt Entra tenant-id |
| `MARIMO_BASE_URL` | Valgfri dersom appen ligger bak base path |

I første Azure-pilot kan `DATABASE_URL` brukes for å verifisere funksjonaliteten. Før endelig produksjonssetting bør Managed Identity/Entra vurderes som hovedspor slik at et permanent databasepassord ikke må lagres som app-setting.

## 5. Authentication for sluttbrukere

I **App Service → Authentication**:

1. Legg til **Microsoft** som identity provider.
2. Bruk single-tenant-oppsett for FramTre-miljøet.
3. Krev autentisering før tilgang til applikasjonen.
4. Behold applikasjonens tenant-validering via `ALLOWED_TENANT_ID` som ekstra kontroll.

Før go-live må det i tillegg avgjøres hvilke Entra-grupper/roller som skal kunne:

- lese og simulere
- importere Excel
- endre stamdata
- administrere historikk og andre sensitive funksjoner

Tenant-tilhørighet alene skal ikke automatisk bety administratortilgang.

## 6. PostgreSQL-oppsett

Anbefalt produksjonsspor:

- Azure Database for PostgreSQL Flexible Server
- egen database for ProduksjonsKalkyle
- TLS på databaseforbindelsen
- Microsoft Entra-autentisering aktivert
- App Service Managed Identity gitt en PostgreSQL-rolle med minst mulige nødvendige rettigheter
- separat administratorrolle for databaseadministrasjon

I CI brukes lokal PostgreSQL 16 med `sslmode=disable`. Dette gjelder kun den midlertidige GitHub Actions-containeren og skal ikke kopieres til produksjonsoppsettet.

## 7. Databaseberedskap

Datalaget har nå:

- PostgreSQL-støtte via `DataRepo`
- schema-baseline `CURRENT_SCHEMA_VERSION = 1`
- `schema_version` i databasen
- eksplisitt stopp ved inkompatibel schema-versjon
- atomiske CRUD- og audit-transaksjoner
- testet Excel-import og -eksport
- testet migreringsscript fra SQLite til PostgreSQL
- migrering av `change_log` og `uploaded_files` inkludert blob-data
- radtallsvalidering etter ren migrering

Permanent database-CI:

```text
.github/workflows/postgres-compatibility.yml
```

## 8. Første Azure-pilot

Anbefalt rekkefølge:

1. Opprett PostgreSQL Flexible Server og en tom testdatabase.
2. Konfigurer nettverk/TLS og databaseautentisering.
3. Opprett App Service og aktiver Entra-innlogging.
4. Deploy containeren fra `azure-readiness` eller en senere godkjent branch.
5. Verifiser `schema_version=1` og at appen bruker PostgreSQL-backend.
6. Kjør smoke-test mot Azure-databasen.
7. Migrer en kopi av eksisterende SQLite-data.
8. Sammenlign radtall og et utvalg viktige dataobjekter.
9. Verifiser kjente produksjonskalkyler mot dagens forventede resultat.
10. Test flere samtidige brukere, reconnect og Excel-import.

Ingen produksjonsdata skal flyttes før pilotløpet er godkjent.

## 9. CI/CD

Repoet har nå CI som:

- bygger App Service-containeren
- kompilerer Python-koden
- starter en midlertidig PostgreSQL 16-instans
- tester databaseinitialisering og schema-versjon
- tester Excel begge veier
- tester SQLite -> PostgreSQL-migrering
- tester atomiske transaksjoner på PostgreSQL og SQLite

Det som mangler er CD-delen: faktisk push av containerimage og deploy til en konkret Azure App Service. Den workflowen bør først legges inn når Azure-ressurs, registry/deploymetode og miljønavn er bestemt.

## 10. Go-live kriterier

Løsningen bør ikke regnes som produksjonsklar før:

- Azure PostgreSQL og App Service er opprettet og testet
- databaseautentisering og nettverk/TLS er verifisert
- Entra-grupper/roller er avklart og testet
- migrering av reelle data er gjennomført i et testmiljø og kontrollert
- faglige referansekalkyler gir forventet resultat
- flerbrukerdrift er testet i App Service
- deployløpet fra GitHub til Azure er kontrollert

## 11. Drift etter go-live

Skjemaendringer skal heretter behandles som eksplisitte databaseversjoner. Dagens baseline er versjon 1. Når datamodellen endres, skal kode og database-migrering følge samme versjonsendring og testes i CI før produksjonsdeploy.

Før større modell- eller datamigreringer skal det tas databasebackup og gjennomføres verifisering av sentrale kalkyler etter endringen.
