# PostgreSQL-ferdigstillelse

Dette dokumentet beskriver hva som konkret gjenstår for å gjøre PostgreSQL til fullverdig produksjonsbackend i ProduksjonsKalkyle.

## Mål

Målet er at applikasjonen skal kunne kjøre i Azure med PostgreSQL som primær databasebackend, uten å være avhengig av SQLite-spesifikk oppførsel.

## 1. `src/data_repo.py`

### Status

Det finnes allerede:

- deteksjon av `DATABASE_URL`
- `psycopg.connect(...)`
- `POSTGRES_SCHEMA_SQL`
- flere backend-nøytrale hjelpefunksjoner
- flere `ON CONFLICT ... DO UPDATE ... RETURNING id`-mønstre

### Gjenstående arbeid

Gå gjennom alle metoder og fjern/erstatt SQLite-spesifikke antakelser:

- `sqlite3.Row` i typehint og returforventninger
- `sqlite3.OperationalError` som kontrollflyt
- `PRAGMA`-basert logikk
- SQLite-spesifikke migreringer som ikke gir mening for PostgreSQL
- eventuelle antakelser om SQLite-returnerte radtyper

### Særlig å kontrollere

- statistikkfunksjoner
- endringslogg-funksjoner
- upsert-funksjoner for alle hovedtabeller
- `demand`, `historical_sales`, `uploaded_files`
- eksportfunksjoner og tabell-lesing

## 2. `src/excel_bridge.py`

Denne filen må gjennomgås for backend-spesifikk SQL.

### Eksempler som må kontrolleres

- `CREATE TABLE ... AUTOINCREMENT`
- `datetime('now')`
- direkte SQLite-tabellopprettelse i hjelpefunksjoner
- SQL som forutsetter SQLite-syntaks

Målet er at Excel-import/-eksport skal fungere likt mot PostgreSQL.

## 3. Migrering av eksisterende data

Det må lages et kontrollert migreringsløp fra SQLite til PostgreSQL.

Migreringen bør:

1. lese alle relevante tabeller fra SQLite
2. skrive til PostgreSQL i riktig rekkefølge
3. verifisere radantall per tabell
4. verifisere et utvalg nøkkelobjekter
5. verifisere et utvalg beregninger etter migrering

## 4. Testløp som må være grønt

Følgende må valideres mot ekte PostgreSQL:

1. `DataRepo().initialize()`
2. import fra Excel
3. endringslogg
4. historikk for opplastede filer
5. datamodell-innsyn i appen
6. beregning av produksjonskost
7. simuleringer
8. eksport til PDF / Excel / JSON

## 5. Produksjonskrav før PostgreSQL kan regnes som ferdig

PostgreSQL-sporet er ferdig når:

- appen starter stabilt med `DATABASE_URL`
- alle hovedfunksjoner virker uten SQLite-fallback
- migrering fra eksisterende database er gjennomført og validert
- Azure-lanseringsdokumentasjon er oppdatert til PostgreSQL som primærspor

## 6. Anbefalt videre arbeid

Prioritert rekkefølge:

1. rydde `src/data_repo.py`
2. rydde `src/excel_bridge.py`
3. lage migreringsscript
4. kjøre ende-til-ende verifisering mot PostgreSQL
5. oppdatere teknisk dokumentasjon