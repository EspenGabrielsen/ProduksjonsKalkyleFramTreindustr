# Azure App Service-oppsett for ProduksjonsKalkyle

Dette prosjektet kan kjøres på Azure App Service med Microsoft Entra ID (EasyAuth) slik at kun brukere i egen tenant får tilgang.

## 1. Obligatoriske app settings

Sett følgende under **App Service → Configuration → Application settings**:

| Navn | Eksempelverdi | Formål |
|------|---------------|--------|
| `WEBSITES_PORT` | `8000` | Port Azure sender trafikk til |
| `PORT` | `8000` | Brukes av `startup.sh` |
| `PRODUKSJONSKALKYLE_TEST` | `false` | Sikrer at appen ikke bruker testdatabase |
| `PRODUKSJONSKALKYLE_DB_PATH` | `/home/data/produksjonskalkyle.db` | Persistent SQLite-path inntil PostgreSQL er på plass |
| `ALLOWED_TENANT_ID` | `00000000-0000-0000-0000-000000000000` | Ekstra validering i kode mot tillatt tenant |
| `MARIMO_BASE_URL` | `/` eller `/produksjonskalkyle` | Valgfritt dersom appen ligger bak en base path |

Når PostgreSQL/Blob Storage tas i bruk, legg også til:

| Navn | Formål |
|------|--------|
| `DATABASE_URL` | PostgreSQL-tilkobling |
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

## 6. App Service innstillinger

Anbefalte innstillinger:

- **Web Sockets** = On
- **Always On** = On
- **HTTP version** = 2.0

## 7. Docker-deploy

Dette repoet inneholder nå:

- `Dockerfile`
- `startup.sh`
- `runtime.txt`

Du kan derfor deploye enten:

1. som vanlig Python App Service med Oryx + `startup.sh`, eller
2. som container-basert App Service via `Dockerfile`

## 8. Viktig om dagens database

Prosjektet bruker fortsatt SQLite lokalt i koden. For Azure-produksjon bør dette erstattes med PostgreSQL før appen tas i ordinær drift med flere brukere.

Inntil PostgreSQL er på plass, bruk alltid en persistent path som:

```text
/home/data/produksjonskalkyle.db
```

Applikasjonen støtter nå dette via `PRODUKSJONSKALKYLE_DB_PATH`.