# PostgreSQL-ferdigstillelse

Dette dokumentet beskriver status for PostgreSQL-sporet i ProduksjonsKalkyle etter arbeidet på branchen `azure-readiness`.

## Status

PostgreSQL er nå en testet applikasjonsbackend, ikke bare et planlagt spor. Følgende valideres automatisk i GitHub Actions mot en helt ny PostgreSQL 16-instans:

- opprettelse av tom database og skjema
- eksplisitt `schema_version`
- CRUD via `DataRepo`
- lesing gjennom `DatabaseData`, som kalkylemotoren bruker
- Excel-import til PostgreSQL
- Excel-eksport fra PostgreSQL
- transportflagg fra Excel-import
- SQLite -> PostgreSQL-migrering
- migrering av endringslogg og opplastede filer/blob
- radtallsvalidering mellom kilde og mål ved ren migrering
- atomisk rollback av data og audit-logg ved mislykket mutasjon

SQLite kjøres i tillegg gjennom egen atomisitetstest for å beskytte dagens lokale/testbaserte flyt.

Containeren som er beregnet for Azure App Service bygges også automatisk i GitHub Actions.

## Implementerte PostgreSQL-rettelser

`src/data_repo.py` håndterer nå blant annet:

- `DATABASE_URL` og `psycopg`
- backend-nøytral placeholder-konvertering
- PostgreSQL `executemany()` via cursor
- sikker initialisering av helt tom PostgreSQL-database
- sitert `"user"` i `change_log`
- atomiske CRUD-operasjoner der dataendring og audit-logg committes eller rulles tilbake sammen
- `schema_version` med baseline versjon 1
- stopp dersom databasen har en schema-versjon som ikke støttes av applikasjonen

SQLite-spesifikk `PRAGMA`- og legacy-migreringslogikk ligger fortsatt eksplisitt bak `backend == "sqlite"` og kjøres ikke på PostgreSQL.

## Excel-bro

`src/excel_bridge.py` er validert ende-til-ende mot PostgreSQL. CI lager en faktisk `.xlsx`, importerer den til PostgreSQL, verifiserer dataene og eksporterer databasen tilbake til Excel.

Dette betyr at Excel-broen ikke lenger regnes som en PostgreSQL-blocker.

## Migrering fra SQLite

Migreringsscript:

```text
src/scripts/migrate_sqlite_to_postgres.py
```

Ved normal migrering:

1. SQLite-kilden åpnes uten å kjøre schema-migreringer mot kildefilen.
2. PostgreSQL-målet initialiseres og tømmes.
3. Stamdata, optimeringsdata, transportdata, `change_log` og `uploaded_files` migreres.
4. Opprinnelig audit-historikk og filblobber bevares.
5. Radtall sammenlignes mellom kilde og mål.
6. Ved avvik eller feil tømmes målet igjen for å unngå en halvferdig ren migrering.

`--keep-target-data` finnes for spesielle merge-scenarier. Eksakt radtallsvalidering er da deaktivert fordi målet kan inneholde legitime eksisterende data.

## Automatiske tester

Permanent CI ligger i:

```text
.github/workflows/postgres-compatibility.yml
.github/workflows/azure-container.yml
```

Viktige testscripts:

```text
src/scripts/postgres_smoke_test.py
src/scripts/postgres_migration_test.py
src/scripts/postgres_transaction_test.py
src/scripts/database_schema_version_test.py
```

## Det som fortsatt krever Azure-miljø

Følgende kan ikke ferdigstilles bare i kildekoden og GitHub-testmiljøet:

- opprette Azure Database for PostgreSQL Flexible Server
- opprette/deploye Azure App Service
- bestemme og konfigurere produksjonsautentisering mot PostgreSQL, helst Managed Identity/Entra fremfor permanent databasepassord
- produksjons-TLS og nettverksregler
- definere hvilke Entra-grupper/roller som kan lese, simulere, importere og endre data
- CI/CD som faktisk deployer til en konkret Azure App Service
- flerbrukertest av Marimo/WebSocket/reconnect i virkelig App Service
- faglig sammenligning av kjente referansekalkyler før go-live

## Konklusjon

PostgreSQL-sporet er nå teknisk egnet for neste fase: test mot en ekte Azure PostgreSQL-instans og deretter pilotdeploy av App Service.

Det betyr ikke at produksjonsmiljøet er ferdig. De viktigste gjenstående punktene er nå Azure-infrastruktur, sikkerhet/tilgang, faktisk deploy og faglig go-live-verifisering – ikke en grunnleggende omskriving av datalaget.
