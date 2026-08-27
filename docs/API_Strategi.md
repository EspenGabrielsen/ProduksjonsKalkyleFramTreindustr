# API-strategi for beregningsresultater

Dette dokumentet beskriver anbefalt retning for å eksponere ProduksjonsKalkyle til andre systemer.

## 1. Hva andre systemer bør hente

Det anbefales at andre systemer henter:

- beregnet produksjonskost per produkt
- kostnedbrytning (materiale, operasjon, setup, biprodukt, transport)
- scenario- og simuleringsresultater

Det anbefales **ikke** at integrasjoner må lese rå BOM-, routing- og kosttabeller direkte for å utføre egne beregninger.

## 2. Begrunnelse

Beregningen i denne løsningen er domene- og regelstyrt. Hvis andre systemer leser råtabeller direkte, må de også:

- forstå datamodellen i detalj
- reimplementere kostlogikken
- håndtere endringer i regler og struktur

Dette skaper høy kobling og større risiko for avvik mellom systemene.

## 3. Anbefalt prinsipp

Bruk databasen som kilde for modellen, men eksponer **beregning som tjeneste**.

Anbefalt mønster:

- PostgreSQL lagrer modellgrunnlaget
- beregningsmotoren i Python utfører kalkylene
- et API eksponerer resultatene til andre systemer

## 4. Foreslåtte API-use-cases

Eksempler på fremtidige endepunkter:

- `GET /api/products/{item_no}/cost`
- `GET /api/products/{item_no}/cost-breakdown`
- `POST /api/calculate`
- `POST /api/simulations`

Typiske input kan være:

- produktnummer
- lokasjon
- scenario-navn
- overstyringer av sentrale parametere

Typiske output kan være:

- netto produksjonskost
- materialkost
- operasjonskost
- setupkost
- biproduktverdi
- transportpåslag
- detaljlinjer per komponent og operasjon

## 5. Teknisk anbefaling

Når API-et bygges, anbefales et dedikert API-lag, for eksempel med **FastAPI**.

Hvorfor:

- tydelige HTTP-endepunkter
- automatisk OpenAPI-dokumentasjon
- god støtte for Azure-deploy
- enkel autentisering og autorisasjon senere

## 6. Foreslått rekkefølge

1. fullfør PostgreSQL-sporet
2. deploy web-appen i Azure
3. ekstraher beregningslogikk til tydelig service-lag ved behov
4. bygg API for beregningsresultater

## 7. Avgrensning

API-et bør i første omgang fokusere på **lesing/beregning**, ikke full CRUD i modellen.

Det passer godt med arbeidsformen der innsatsfaktorer vedlikeholdes kontrollert og sjelden av et avgrenset fagmiljø.