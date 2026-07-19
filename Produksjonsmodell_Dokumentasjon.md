# Produksjonsmodell - Dokumentasjon

> **Produksjonsmodell for trelastindustrien**
> Beregning av standardkost, simulering av produksjon og analyse av lønnsomhet
> Basert pa Kodal Hovleri som referanseeksempel

---

## Innholdsfortegnelse

1. [Innledning](#1-innledning)
2. [Dataark - oversikt](#2-dataark--oversikt)
3. [Product Master](#3-product-master)
4. [Locations](#4-locations)
5. [Work Centers](#5-work-centers)
6. [Operation Master](#6-operation-master)
7. [Item Costs](#7-item-costs)
8. [BOM (Stykkliste)](#8-bom-stykkliste)
9. [Routing](#9-routing)
10. [By Product Rules](#10-by-product-rules)
11. [Beregningsmodell](#11-beregningsmodell)
12. [Eksempel - Kodal Hovleri](#12-eksempel--kodal-hovleri)
13. [Brukerveiledning](#13-brukerveiledning)
14. [JSON-eksport](#14-json-eksport)
15. [Tilpasning og vedlikehold](#15-tilpasning-og-vedlikehold)


---

## 1. Innledning

### 1.1 Formal

Modellen skal brukes til:

- **Beregning av standardkost** per produkt (materialkost, operasjonskost, setupkost)
- **Simulering av produksjon** - hva koster det a produsere X antall enheter?
- **Analyse av lonnsomhet** per produkt og per produktgruppe
- **Analyse av kapasitetsutnyttelse** - hvor mange timer trengs per arbeidssenter?
- **Beregning av biproduktverdi** - hovelspon, flis, bark
- **Omstillingsanalyse** - hva koster det a bytte mellom produkter pa samme maskin?
- **Datagrunnlag** for Business Central og Power BI

### 1.2 Overordnet arkitektur

```
+-----------------------------------------------------------+
|                     Excel-datamodell                      |
|  +----------+  +----------+  +----------+  +------------+ |
|  | Product  |  |  Work    |  |   BOM    |  |  Routing   | |
|  |  Master  |  | Centers  |  |          |  |            | |
|  +----------+  +----------+  +----------+  +------------+ |
|  +----------+  +----------+                               |
|  |  Item    |  |  By Prod |                               |
|  |  Costs   |  |  Rules   |                               |
|  +----------+  +----------+                               |
+---------------------------+-------------------------------+
                            |
                            v
+-----------------------------------------------------------+
|              Python-beregningsmotor                       |
|  +------------------------------------------------------+ |
|  |  CostCalculator                                      | |
|  |  - Materialkost (fra BOM + Item Costs)               | |
|  |  - Operasjonskost (fra Routing + Work Centers)       | |
|  |  - Setupkost (fordelt pa batch size)                 | |
|  |  - Biproduktverdi (fra By Product Rules)             | |
|  |  - Dynamisk cost roll-up (FG -> FG-avhengigheter)    | |
|  +------------------------------------------------------+ |
+---------------------------+-------------------------------+
                            |
                            v
          +-----------------------------------+
          |  Resultater                       |
          |  - Konsoll (tekst)                |
          |  - JSON (produktkalkyle.json)     |
          |  - Marimo (analyse og simulering) |
          +-----------------------------------+
```


### 1.3 Vareflyt - eksempel

```
48x198 Skrulast (RM001)
         |
         v
    +------------+
    | Oppdeling  |  <- HOVEDHOVEL
    +------------+
         |
         v
    +------------+
    |  Hovling   |  <- HOVEDHOVEL
    +------------+
         |
    +----+------------------+
    v                       v
+------------+       +--------------+
| Profilering|       | Hovelspon    |
| (SPESIAL-  |       | (BP001)      |
|  HOVEL)    |       | Flis (BP002) |
+------------+       +--------------+
    |
    v
+------------+
|  Pakking   |  <- PAKKELINJE
+------------+
    |
    v
Utvendig Panel 21x95 (FG001)
```

---

## 2. Dataark - oversikt

Modellen bestar av **8 ark** i Excel. Hvert ark har en spesifikk rolle:

| # | Arknavn | Innhold | Nokkelkolonner |
|---|---------|---------|----------------|
| 1 | **Product Master** | Vareregister | Item No, Description, Item Type |
| 2 | **Locations** | Fabrikker og lagre | Location Code, Location Name |
| 3 | **Work Centers** | Arbeidssentre med kostsatser | Work Center Code, Labor/Machine/Overhead Cost |
| 4 | **Operation Master** | Standardoperasjoner | Operation Code, Description |
| 5 | **Item Costs** | Kostpriser per vare | Item No, Unit Cost, Currency |
| 6 | **BOM** | Stykkliste (hva bestar produktet av) | Parent Item, Component, Quantity Per |
| 7 | **Routing** | Produksjonsflyt (operasjoner, tider) | Item No, Operation No, Setup/Run Time |
| 8 | **By Product Rules** | Biprodukter og verdsetting | Parent Item, By Product, Market Value |


---

## 3. Product Master

### 3.1 Formal

Register over alle varer i virksomheten. Dette er hovedkatalogen over alt som finnes - ravarer, halvfabrikata, ferdigvarer, biprodukter og handelsvarer.

### 3.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Item No** | Tekst | Unik identifikator for varen | `RM001`, `FG001`, `BP001` |
| **Description** | Tekst | Beskrivende navn pa varen | `Skrulast 48x198`, `Utvendig Panel 21x95` |
| **Item Type** | Tekst | Type vare | `Raw Material`, `Semi Finished`, `Finished Good`, `By Product`, `Trading Item` |
| **Product Group** | Tekst | Gruppering av varer | `Skrulast`, `Panel`, `Kledning`, `Spon` |
| **Base Unit of Measure** | Tekst | Standard maleenhet | `LM`, `M3`, `KG`, `PCS` |
| **Active** | Ja/Nei | Angir om varen er aktiv | `Ja` |

### 3.3 Varetyper

```
Raw Material    -> Ravares som kjopes inn (f.eks. skrulast, maling)
Semi Finished   -> Halvfabrikat (mellomprodukt)
Finished Good   -> Ferdigvare som selges (f.eks. panel, terrassebord)
By Product      -> Biprodukt som oppstar i produksjon (f.eks. hovelspon, flis)
Trading Item    -> Handelsvare (kjopes og selges uendret)
```

### 3.4 Testdata - Kodal Hovleri

| Item No | Description | Item Type | Product Group | UOM | Aktiv |
|---------|-------------|-----------|---------------|-----|-------|
| RM001 | Skrulast 48x198 | Raw Material | Skrulast | M3 | Ja |
| RM002 | Gran 36x148 | Raw Material | Skrulast | M3 | Ja |
| RM003 | Maling - Hvit | Raw Material | Maling | LTR | Ja |
| FG001 | Utvendig Panel 21x95 | Finished Good | Panel | LM | Ja |
| FG002 | Terrassebord 28x120 | Finished Good | Terrasse | LM | Ja |
| FG003 | Kledning 18x120 | Finished Good | Kledning | LM | Ja |
| FG004 | Utvendig Panel 21x95 - Malt | Finished Good | Panel | LM | Ja |
| BP001 | Hovelspon | By Product | Spon | KG | Ja |
| BP002 | Flis | By Product | Spon | KG | Ja |
| BP003 | Bark | By Product | Spon | KG | Ja |

---

## 4. Locations

### 4.1 Formal

Register over fabrikker og lokasjoner. Hvert arbeidssenter er knyttet til en lokasjon.

### 4.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Location Code** | Tekst | Unik kode for lokasjonen | `KOD` |
| **Location Name** | Tekst | Navn pa lokasjonen | `Kodal Fabrikk` |
| **Location Type** | Tekst | Type lokasjon | `Factory`, `Warehouse`, `Distribution Center`, `Sales Office` |
| **Active** | Ja/Nei | Angir om lokasjonen er aktiv | `Ja` |

### 4.3 Testdata

| Location Code | Location Name | Location Type | Aktiv |
|---------------|---------------|---------------|-------|
| KOD | Kodal Fabrikk | Factory | Ja |
| SKI | Skien Lager | Warehouse | Ja |

---

## 5. Work Centers

### 5.1 Formal

Register over produksjonsressurser - maskiner og arbeidsplasser. Hvert arbeidssenter har timekostnader som brukes til a beregne operasjonskost.

### 5.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Work Center Code** | Tekst | Unik identifikator | `HOVEDHOVEL` |
| **Description** | Tekst | Beskrivende navn | `Hovedhovel` |
| **Location Code** | Tekst | Fabrikken det tilhorer | `KOD` |
| **Labor Cost per Hour** | Desimal | Arbeidskostnad per time (lonn, arbeidsgiveravgift, pensjon, feriepenger) | `550` |
| **Machine Cost per Hour** | Desimal | Maskinkostnad per time (avskrivninger, service, leasing, vedlikehold, energi) | `900` |
| **Overhead Cost per Hour** | Desimal | Indirekte produksjonskostnader (produksjonsledelse, kvalitet, vedlikeholdsadm., intern logistikk) | `150` |
| **Capacity Hours per Day** | Desimal | Tilgjengelige timer per dag | `16` |
| **Effective Capacity %** | Prosent | Hvor stor del av tiden som faktisk kan brukes til produksjon (tar hensyn til stopp, vedlikehold, feil) | `85` |
| **Active** | Ja/Nei | Angir om arbeidssenteret er aktivt | `Ja` |

### 5.3 Timekostnad (Justert for Effektiv Kapasitet)

For å sikre at alle kalkyler tar høyde for uunngåelig ståtid, mikrostopp og planlagt vedlikehold, justeres den timeprisen som belastes produktet opp med maskinens effektivitetsgrad (**Effective Capacity %**):

$$\text{Effektiv Timepris} = \frac{\text{Nominell Timepris (Lønn + Maskin + Overhead)}}{\frac{\text{Effective Capacity \%}}{100}}$$

#### Eksempel - HOVEDHOVEL (85 % effektivitet):
* Nominell timepris = $550 \text{ kr (lønn)} + 900 \text{ kr (maskin)} + 150 \text{ kr (overhead)} = 1600 \text{ kr/time}$
* Effektiv timepris = $\frac{1600 \text{ kr}}{0,85} = \mathbf{1882,35 \text{ kr/time}}$

Dette betyr at produktet belastes 1882,35 kr per time i stedet for 1600 kr. På denne måten blir de 15 % med tapt produksjonstid automatisk og nøyaktig bakt inn i produktkalkylen!

### 5.4 Effektiv kapasitet

```
Effektive timer per dag = Capacity Hours x Effective Capacity %

Eksempel - HOVEDHOVEL:
  16 timer x 85 % = 13,6 effektive timer
```

### 5.5 Testdata

| Work Center | Beskrivelse | Lokasjon | Labor | Maskin | Overhead | Timer/dag | Eff. % | Aktiv |
|-------------|-------------|----------|-------|--------|----------|-----------|--------|-------|
| HOVEDHOVEL | Hovedhovel | KOD | 550 | 900 | 150 | 16 | 85 | Ja |
| SPESIALHOVEL | Spesialhovel | KOD | 550 | 950 | 150 | 16 | 85 | Ja |
| MALINGSLINJE | Malingslinje | KOD | 500 | 400 | 120 | 16 | 80 | Ja |
| PAKKELINJE | Pakkelinje | KOD | 450 | 300 | 100 | 8 | 90 | Ja |

---

## 6. Operation Master

### 6.1 Formal

Standardisert liste over operasjoner som kan brukes i routing. Gir en felles "ordbok" for produksjonsprosesser.

### 6.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Operation Code** | Tekst | Unik operasjonskode | `RIP`, `PLANING`, `PROFILE` |
| **Description** | Tekst | Beskrivelse av operasjonen | `Oppdeling`, `Hovling`, `Profilering` |
| **Default Work Center** | Tekst | Anbefalt arbeidssenter | `HOVEDHOVEL` |
| **Standard Unit** | Tekst | Maleenhet for produksjonstid | `Minutes`, `Hours` |
| **Active** | Ja/Nei | Angir om operasjonen er aktiv | `Ja` |

### 6.3 Testdata

| Operation Code | Description | Default Work Center | Standard Unit | Aktiv |
|----------------|-------------|---------------------|---------------|-------|
| RIP | Oppdeling | HOVEDHOVEL | Minutes | Ja |
| PLANING | Hovling | HOVEDHOVEL | Minutes | Ja |
| PROFILE | Profilering | SPESIALHOVEL | Minutes | Ja |
| MALING | Maling | MALINGSLINJE | Minutes | Ja |
| PACKING | Pakking | PAKKELINJE | Minutes | Ja |

---

## 7. Item Costs

### 7.1 Formal

Samlet register over kostpriser for alle varer. For ravarer er dette innkjopspris. For ferdigvarer settes prisen til 0 (beregnes automatisk). For biprodukter er dette markedsverdi.

### 7.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Item No** | Tekst | Referanse til varen (fra Product Master) | `RM001` |
| **Cost Type** | Tekst | Type kostpris | `Standard Cost`, `Last Direct Cost`, `Forecast Cost`, `Budget Cost` |
| **Unit Cost** | Desimal | Kostpris per enhet | `3000.00` |
| **Currency** | Tekst | Valuta | `NOK`, `EUR` |
| **Effective Date** | Dato | Dato kostprisen gjelder fra | `2026-01-01` |

### 7.3 Testdata

| Item No | Cost Type | Unit Cost | Currency | Effective Date |
|---------|-----------|-----------|----------|----------------|
| RM001 | Standard Cost | 3 000,00 | NOK | 2026-01-01 |
| RM002 | Standard Cost | 2 500,00 | NOK | 2026-01-01 |
| RM003 | Standard Cost | 120,00 | NOK | 2026-01-01 |
| FG001 | Standard Cost | 0,00 | NOK | 2026-01-01 |
| FG002 | Standard Cost | 0,00 | NOK | 2026-01-01 |
| FG003 | Standard Cost | 0,00 | NOK | 2026-01-01 |
| FG004 | Standard Cost | 0,00 | NOK | 2026-01-01 |
| BP001 | Standard Cost | 1,50 | NOK | 2026-01-01 |
| BP002 | Standard Cost | 0,80 | NOK | 2026-01-01 |
| BP003 | Standard Cost | 0,50 | NOK | 2026-01-01 |

> **Merk:** Ferdigvarer (FG001-FG004) har kostpris 0,00 fordi kostnaden beregnes automatisk fra BOM og Routing.

---

## 8. BOM (Stykkliste)

### 8.1 Formal

Beskriver hvilke komponenter som inngar i et produkt. En BOM-linje sier: "For a lage X trenger du Y, og du far Z enheter output per enhet input."

### 8.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Parent Item No** | Tekst | Produktet som produseres | `FG001` |
| **Component Item No** | Tekst | Komponenten som forbrukes | `RM001` |
| **Quantity Per** | Desimal | Antall output-enheter per input-enhet | `400` |
| **Unit of Measure** | Tekst | Maleenhet for forholdet | `LM` |
| **Scrap %** | Prosent | Forventet materialsvinn | `5.0` |
| **Valid From** | Dato | Gyldig fra dato | `2026-01-01` |
| **Valid To** | Dato | Gyldig til dato (tom = alltid gyldig) | |

### 8.3 Hvordan Quantity Per fungerer

```
Quantity Per = antall output-enheter per input-enhet

Forbruk per output = 1 / Quantity Per

Eksempel - FG001 (Panel) fra RM001 (Skrulast):
  Quantity Per = 400 LM per M3
  Forbruk per LM panel = 1 / 400 = 0,0025 M3 per LM
```

### 8.4 Materialkost-beregning

```
Materialkost per enhet = (Unit Cost / Quantity Per) x (1 + Scrap% / 100)

Eksempel - FG001:
  = (3 000 kr/M3 / 400 LM/M3) x (1 + 0,05)
  = 7,50 x 1,05
  = 7,875 kr/LM
```

### 8.5 Testdata

| Parent | Component | Quantity Per | UOM | Scrap % | Valid From |
|--------|-----------|-------------|-----|---------|------------|
| FG001 | RM001 | 400 | LM | 5,0 | 2026-01-01 |
| FG002 | RM002 | 250 | LM | 4,0 | 2026-01-01 |
| FG003 | RM001 | 420 | LM | 6,0 | 2026-01-01 |
| FG004 | FG001 | 1 | LM | 2,0 | 2026-01-01 |
| FG004 | RM003 | 20 | LTR | 3,0 | 2026-01-01 |

> **Merk:** FG004 (malt panel) har to BOM-linjer: den bruker FG001 (ubehandlet panel) som komponent i tillegg til maling. Dette kalles en **produksjonskjede** - FG004 bygger pa FG001.

---

## 9. Routing

### 9.1 Formal

Beskriver produksjonsprosessen - hvilke operasjoner som utfores, i hvilken rekkefolge, pa hvilket arbeidssenter, og hvor lang tid hver operasjon tar.

### 9.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Item No** | Tekst | Produkt som produseres | `FG001` |
| **Operation No** | Heltall | Sekvensnummer (stigende rekkefolge) | `10`, `20`, `30` |
| **Operation Code** | Tekst | Hvilken operasjon som utfores | `RIP`, `PLANING` |
| **Work Center Code** | Tekst | Arbeidssenter som utforer operasjonen | `HOVEDHOVEL` |
| **Setup Time Minutes** | Desimal | Tid til klargjoring (omstilling, knivbytte, innkjoring, kontrollmaling) | `15.0` |
| **Run Time Minutes** | Desimal | Produksjonstid per enhet | `0.15` |
| **Batch Size** | Desimal | Normal ordrestorrelse (brukes til a fordele setupkostnad) | `500` |
| **Valid From** | Dato | Gyldig fra dato | `2026-01-01` |
| **Valid To** | Dato | Gyldig til dato | |

### 9.3 Operasjonskost-beregning

```
Kjoretid per enhet (timer) = Run Time Minutes / 60
Kjorekost per enhet = Kjoretid x Timekost

Setupkost per enhet = (Setup Time / 60 x Timekost) / Batch Size

Eksempel - Op 10 (Oppdeling) for FG001:
  Kjoretid: 0,15 min / 60 = 0,0025 timer
  Kjorekost: 0,0025 x 1 600 = 4,00 kr/LM
  Setupkost: (15/60 x 1 600) / 500 = 400 / 500 = 0,80 kr/LM
```

### 9.4 Testdata - FG001 (Utvendig Panel 21x95)

| Op.nr | Operasjon | Arbeidssenter | Setup (min) | Kjoretid (min) | Batch |
|-------|-----------|---------------|-------------|----------------|-------|
| 10 | Oppdeling | HOVEDHOVEL | 15,0 | 0,15 | 500 |
| 20 | Hovling | HOVEDHOVEL | 10,0 | 0,10 | 500 |
| 30 | Profilering | SPESIALHOVEL | 20,0 | 0,12 | 500 |
| 40 | Pakking | PAKKELINJE | 5,0 | 0,05 | 500 |

### 9.5 Produksjonsflyt - visuell

```
FG001 - Utvendig Panel 21x95
-----------------------------
  Op 10: RIP       @ HOVEDHOVEL    (15 min setup + 0,15 min/LM)
  Op 20: PLANING   @ HOVEDHOVEL    (10 min setup + 0,10 min/LM)
  Op 30: PROFILE   @ SPESIALHOVEL  (20 min setup + 0,12 min/LM)
  Op 40: PACKING   @ PAKKELINJE    ( 5 min setup + 0,05 min/LM)

FG002 - Terrassebord 28x120
-----------------------------
  Op 10: RIP       @ HOVEDHOVEL    (12 min setup + 0,18 min/LM)
  Op 20: PLANING   @ HOVEDHOVEL    ( 8 min setup + 0,12 min/LM)
  Op 30: PROFILE   @ SPESIALHOVEL  (15 min setup + 0,15 min/LM)

FG003 - Kledning 18x120
-----------------------------
  Op 10: RIP       @ HOVEDHOVEL    (10 min setup + 0,10 min/LM)
  Op 20: PLANING   @ HOVEDHOVEL    ( 8 min setup + 0,08 min/LM)

FG004 - Utvendig Panel 21x95 - Malt
-----------------------------
  Op 10: MALING    @ MALINGSLINJE  (30 min setup + 0,20 min/LM)
  Op 20: PACKING   @ PAKKELINJE    ( 5 min setup + 0,05 min/LM)

```

---

## 10. By Product Rules

### 10.1 Formal

Beskriver hvordan biprodukter handteres okonomisk. I trelastproduksjon oppstar det alltid biprodukter som hovelspon, flis og bark. Disse har en verdi som skal trekkes fra hovedproduktets kostnad.

### 10.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Parent Item No** | Tekst | Produktet (ferdigvaren) som skaper biproduktet | `FG001` |
| **By Product Item No** | Tekst | Biproduktet | `BP001` |
| **Expected Quantity** | Desimal | Forventet mengde biprodukt per enhet hovedprodukt | `0.5` |
| **Unit of Measure** | Tekst | Maleenhet for biproduktet | `KG` |
| **Market Value** | Desimal | Forventet markedspris per enhet | `1.50` |
| **Allocation Method** | Tekst | Hvordan verdien skal handteres | `Reduce Main Product Cost` |

### 10.3 Allokeringsmetoder

| Metode | Beskrivelse |
|--------|-------------|
| **Reduce Main Product Cost** | Verdien trekkes fra hovedproduktets kost (anbefalt) |
| **Separate Profit Center** | Biproduktet behandles som eget salgbart produkt |
| **Informational Only** | Brukes kun til rapportering, pavirker ikke kost |

### 10.4 Biproduktverdi-beregning

```
Biproduktverdi per enhet = Expected Quantity x Market Value

Eksempel - FG001:
  BP001 (Hovelspon): 0,5 KG x 1,50 kr/KG = 0,75 kr/LM
  BP002 (Flis):      0,3 KG x 0,80 kr/KG = 0,24 kr/LM
  Total:             0,99 kr/LM
```

### 10.5 Testdata

| Parent | By Product | Forventet mengde | UOM | Markedsverdi | Allokeringsmetode |
|--------|------------|-----------------|-----|-------------|-------------------|
| FG001 | BP001 | 0,5 | KG | 1,50 | Reduce Main Product Cost |
| FG001 | BP002 | 0,3 | KG | 0,80 | Reduce Main Product Cost |
| FG002 | BP002 | 0,4 | KG | 0,80 | Reduce Main Product Cost |
| FG002 | BP003 | 0,2 | KG | 0,50 | Reduce Main Product Cost |
| FG003 | BP001 | 0,5 | KG | 1,50 | Reduce Main Product Cost |
| FG003 | BP002 | 0,3 | KG | 0,80 | Reduce Main Product Cost |
| FG004 | BP001 | 0,5 | KG | 1,50 | Reduce Main Product Cost |
| FG004 | BP002 | 0,3 | KG | 0,80 | Reduce Main Product Cost |

### 10.6 Sortering og 2. sortering (A- og B-vare)

I trelastproduksjon (høvlerier) opplever man ofte at en viss andel av produksjonen ikke tilfredsstiller kravene til 1. sortering (A-vare) under den visuelle kontrollen, og blir derfor nedgradert til 2. sortering (B-vare). 

B-varen bruker akkurat det samme råstoffet (tømmeret) og den samme maskintiden som A-varen, men selges til en lavere markedspris. Dette modelleres som et biprodukt for å allokere kostnadene helt korrekt etter bransjestandard:

#### Matematisk eksempel (94 % A-vare og 6 % B-vare):

1. **Brutto råvareforbruk per meter A-vare øker**:
   Når bare 94 % av tømmeret blir til A-vare, betyr det at vi må kjøre mer råstoff gjennom høvelen per ferdige meter A-vare vi legger på lager. 
   Hvis det teoretiske utbyttet av tømmeret er **553,34 meter/m³**, blir det reelle utbyttet av A-vare:
   $$553,34 \text{ meter/m³} \times 0,94 = \mathbf{520,14 \text{ meter A-vare per m³}}$$
   Dette tallet legges inn i **BOM (stykklisten)** under `Quantity Per` for A-varen.
   Dette gjør at råvarekosten per meter A-vare stiger fra **5,12 kr** til **5,44 kr** (fordi vi må fordele råvarekostnaden på færre godkjente meter).

2. **Inntekt fra B-vare trekkes fra som kreditt**:
   For hver meter A-vare vi legger på lager, oppstår det uunngåelig også en andel B-vare:
   $$\text{B-vare andel} = 6 \% / 94 \% \approx \mathbf{0,0638 \text{ meter B-vare per meter A-vare}}$$
   Dette legges inn under `Expected Quantity` i **By Product Rules** for A-varen, med B-varen (`JD16073-B`) som biprodukt.
   
   Hvis B-varen selges for en redusert pris på **3,00 kr/meter** (lagt inn i *Item Costs*), genererer dette en kreditt på:
   $$0,0638 \text{ meter} \times 3,00 \text{ kr/meter} = \mathbf{0,19 \text{ kr/meter}}$$

3. **Netto materialkostnad for A-vare**:
   Netto materialkostnad blir dermed brutto råvarekost fratrukket verdien av B-varene vi selger:
   $$\text{Netto materialkost} = 5,44 \text{ kr} - 0,19 \text{ kr} = \mathbf{5,25 \text{ kr/meter}}$$

Denne metoden er 100 % revisorgodkjent og sikrer et nøyaktig bilde av lønnsomheten, samtidig som den tar høyde for at man ikke ender opp med kun A-vare på lager under produksjon.

---

## 11. Beregningsmodell

### 11.1 Materialkost


```
Materialkost = (Unit Cost / Quantity Per) x (1 + Scrap% / 100)

Hvor:
  Unit Cost    = enhetskost for ravaren (fra Item Costs)
  Quantity Per = antall output-enheter per input-enhet (fra BOM)
  Scrap%       = forventet materialsvinn (fra BOM)

Eksempel - FG001 (Panel) fra RM001 (Skrulast):
  = (3 000 kr/M3 / 400 LM/M3) x (1 + 5/100)
  = 7,50 x 1,05
  = 7,875 kr/LM
```

### 11.2 Operasjonskost


```
Operasjonskost per enhet = (Run Time Minutes / 60) x Timekost

Hvor:
  Run Time Minutes = produksjonstid per enhet (fra Routing)
  Timekost         = Labor + Machine + Overhead (fra Work Centers)

Eksempel - Op 10 (Oppdeling) for FG001:
  Kjoretid: 0,15 min / 60 = 0,0025 timer
  Timekost: 550 + 900 + 150 = 1 600 kr/time
  = 0,0025 x 1 600 = 4,00 kr/LM
```

### 11.3 Setupkost

```
Setupkost per enhet = ((Setup Time Minutes / 60) x Timekost) / Batch Size

Hvor:
  Setup Time Minutes = klargjoringstid per ordre (fra Routing)
  Timekost           = Labor + Machine + Overhead (fra Work Centers)
  Batch Size         = normal ordrestorrelse (fra Routing)

Eksempel - Op 10 (Oppdeling) for FG001:
  Setup-tid: 15 min / 60 = 0,25 timer
  Timekost: 1 600 kr/time
  Batch: 500 LM
  = (0,25 x 1 600) / 500 = 400 / 500 = 0,80 kr/LM
```

### 11.4 Brutto produksjonskost

```
Brutto produksjonskost = Materialkost + Operasjonskost + Setupkost

Eksempel - FG001:
  = 7,875 + 10,675 + 2,575
  = 21,125 kr/LM
```

### 11.5 Biproduktverdi

```
Biproduktverdi per enhet = Sum (Expected Quantity x Market Value)

Eksempel - FG001:
  BP001 (Hovelspon): 0,5 KG x 1,50 kr/KG = 0,750 kr/LM
  BP002 (Flis):      0,3 KG x 0,80 kr/KG = 0,240 kr/LM
  Total:             0,990 kr/LM
```

### 11.6 Netto produksjonskost

```
Netto produksjonskost = Brutto produksjonskost - Biproduktverdi

Eksempel - FG001:
  = 21,125 - 0,990
  = 20,135 kr/LM
```

### 11.7 Oppsummering - FG001

```
+-----------------------------------------------------------+
|  Materialkost:              7,8750 kr/LM                  |
|    RM001: Skrulast 48x198                                 |
|      Pris: 3 000,00 kr/M3                                 |
|      Output per input: 400 LM                             |
|      Svinn: 5 %  ->  7,8750 kr                            |
+-----------------------------------------------------------+
|  Operasjonskost:          10,6750 kr/LM                   |
|  Setupkost:                2,5750 kr/LM                   |
|    Op 10: Oppdeling @ HOVEDHOVEL                          |
|      Kjoretid: 0,15 min  |  Setup: 15 min / batch 500    |
|      Timekost: 1 600 kr/t  ->  4,8000 kr                  |
|    Op 20: Hovling @ HOVEDHOVEL                            |
|      Kjoretid: 0,10 min  |  Setup: 10 min / batch 500    |
|      Timekost: 1 600 kr/t  ->  3,2000 kr                  |
|    Op 30: Profilering @ SPESIALHOVEL                      |
|      Kjoretid: 0,12 min  |  Setup: 20 min / batch 500    |
|      Timekost: 1 650 kr/t  ->  4,4000 kr                  |
|    Op 40: Pakking @ PAKKELINJE                            |
|      Kjoretid: 0,05 min  |  Setup: 5 min / batch 500     |
|      Timekost: 850 kr/t  ->  0,8500 kr                    |
+-----------------------------------------------------------+
|  Brutto produksjonskost:    21,1250 kr/LM                 |
+-----------------------------------------------------------+
|  Biproduktverdi:            0,9900 kr/LM                  |
|    BP001: Hovelspon  0,5 KG x 1,50 kr = 0,7500 kr       |
|    BP002: Flis       0,3 KG x 0,80 kr = 0,2400 kr       |
+-----------------------------------------------------------+
|  NETTO PRODUKSJONSKOST:    20,1350 kr/LM                  |
+-----------------------------------------------------------+
```

---

## 12. Eksempel - Kodal Hovleri

### 12.1 Produktkalkyle for alle produkter

#### FG001 - Utvendig Panel 21x95

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (RM001: Skrulast) | 7,8750 |
| Operasjonskost (4 operasjoner) | 10,6750 |
| Setupkost | 2,5750 |
| **Brutto produksjonskost** | **21,1250** |
| Biproduktverdi (spon + flis) | -0,9900 |
| **Netto produksjonskost** | **20,1350** |

#### FG002 - Terrassebord 28x120

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (RM002: Gran) | 10,4000 |
| Operasjonskost (3 operasjoner) | 12,1250 |
| Setupkost | 2,3646 |
| **Brutto produksjonskost** | **24,8896** |
| Biproduktverdi (flis + bark) | -0,4200 |
| **Netto produksjonskost** | **24,4696** |

#### FG003 - Kledning 18x120

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (RM001: Skrulast) | 7,5714 |
| Operasjonskost (2 operasjoner) | 4,8000 |
| Setupkost | 0,8000 |
| **Brutto produksjonskost** | **13,1714** |
| Biproduktverdi (spon + flis) | -0,9900 |
| **Netto produksjonskost** | **12,1814** |

#### FG004 - Utvendig Panel 21x95 - Malt

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (FG001 + maling) | 27,7275 |
| Operasjonskost (2 operasjoner) | 4,1083 |
| Setupkost | 1,1617 |
| **Brutto produksjonskost** | **32,9975** |
| Biproduktverdi (spon + flis) | -0,9900 |
| **Netto produksjonskost** | **32,0075** |

---

## 13. Brukerveiledning

### 13.1 Forutsetninger

- Python 3.10 eller nyere
- Biblioteker: `pandas`, `openpyxl`

Installasjon:
```bash
pip install pandas openpyxl
```

### 13.2 Tilgjengelige kommandoer

#### Full produktkalkyle

```bash
python kostberegning.py --excel Produksjonsmodell_Testdata_v3.xlsx
```

Dette beregner:
- Produktkalkyle for alle ferdigvarer
- Eksporterer JSON til `produktkalkyle.json`

#### Ett spesifikt produkt

```bash
python kostberegning.py --excel Produksjonsmodell_Testdata_v3.xlsx --product FG001
```

#### Uten JSON-eksport

```bash
python kostberegning.py --excel Produksjonsmodell_Testdata_v3.xlsx --no-json
```

#### Testdata (innebygget, uten Excel)

```bash
python kostberegning.py --test
```

### 13.3 Generere testdata-Excel

```bash
python lag_testdata_v3.py
```

Dette oppretter `Produksjonsmodell_Testdata_v3.xlsx` med:
- 8 ark med testdata
- Beskrivende kommentarer pa header-radene

### 13.4 Alle flagg

| Flagg | Beskrivelse |
|-------|-------------|
| `--excel FIL` | Excel-fil med data (standard: `Produksjonsmodell_Mal.xlsx`) |
| `--test` | Bruk innebygget testdata |
| `--product ITEMNO` | Beregn kost for ett produkt |
| `--no-json` | Ikke eksporter til JSON |
| `--json-dir MAPPE` | Mappe for JSON-eksport (standard: `.`) |

> **Merk:** Scenario-simulering (--scenario) er flyttet til Marimo for analyse.

---

## 14. JSON-eksport

### 14.1 produktkalkyle.json

```json
{
  "type": "product_cost_calculation",
  "exported_at": "2026-07-10T12:00:00",
  "results": [
    {
      "product_no": "FG001",
      "product_desc": "Utvendig Panel 21x95",
      "product_group": "Panel",
      "base_uom": "LM",
      "material_cost": 7.875,
      "operation_cost": 10.675,
      "setup_cost": 2.575,
      "gross_production_cost": 21.125,
      "by_product_value": 0.99,
      "net_production_cost": 20.135,
      "material_details": [...],
      "operation_details": [...],
      "byproduct_details": [...]
    }
  ]
}
```

### 14.2 Bruk i Power BI

JSON-filen kan lastes direkte inn i Power BI:

1. **Hent data** -> **JSON**
2. Velg `produktkalkyle.json`
3. Power BI tolker JSON-strukturen automatisk
4. Bygg rapporter pa kost per produkt, produktgruppe, etc.

---

## 15. Tilpasning og vedlikehold

### 15.1 Legge til nytt produkt

1. **Product Master**: Legg til ny rad med Item No, beskrivelse, type
2. **Item Costs**: Legg til kostpris (0 for ferdigvarer)
3. **BOM**: Legg til stykklistelinjer (hva bestar produktet av?)
4. **Routing**: Legg til produksjonsflyt (operasjoner, tider, arbeidssentre)
5. **By Product Rules**: Legg til biprodukter hvis relevant

### 15.2 Endre kostsatser

Oppdater `Unit Cost` i **Item Costs** for ravarer, eller `Labor/Machine/Overhead Cost per Hour` i **Work Centers**.

### 15.3 Oppdatere testdata-Excel

```bash
python lag_testdata_v3.py
```

Dette genererer Excel-filen pa nytt med all testdata og beskrivelser.

---

> **Dokumentasjon versjon 2.1**
> Sist oppdatert: juli 2026
> Basert pa Kodal Hovleri som referanseeksempel
> **Merk:** `Changeover Time Minutes` er fjernet fra datamodellen. Omstillingskost håndteres gjennom `Setup Time Minutes` (neste produkts setup-tid).