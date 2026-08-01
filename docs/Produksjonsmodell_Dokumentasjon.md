# ProduksjonsKalkyle — Systemdokumentasjon

> **Produksjonsmodell for trelastindustrien**
> Beregning av standardkost, simulering av produksjon og analyse av lønnsomhet
> Basert på Kodal Hovleri som referanseeksempel
>
> **System:** Excel-datamodell → SQLite-database → Python-beregningsmotor → Marimo web-app → PDF/Excel-rapport

---

## Om dette dokumentet

Dette er den tekniske dokumentasjonen for **ProduksjonsKalkyle** — et system for standardkost-beregning og produksjonssimulering bygget for Fram Treindustri / Kodal Hovleri.

Dokumentasjonen dekker:
- **Datamodellen** — Excel-arkene og SQLite-tabellene som utgjør grunnlaget for alle beregninger
- **Beregningslogikken** — hvordan materialkost, operasjonskost, setupkost og biproduktverdi regnes ut
- **Marimo web-appen** — hvordan den brukes til simulering, dataimport og eksport
- **Import/eksport** — hvordan du laster opp data, eksporterer PDF-rapporter og Excel-filer
- **Målgrupper og roller** — hvem som gjør hva i systemet

For sluttbrukerveiledning i Excel-arket, se egen **Brukermanual_Produksjonsmodell.md**.

---

## Innholdsfortegnelse

1. [Målgrupper og roller](#1-målgrupper-og-roller)
2. [Systemarkitektur](#2-systemarkitektur)
3. [Dataflyt](#3-dataflyt)
4. [Marimo web-app — full gjennomgang](#4-marimo-web-app--full-gjennomgang)
   - 4.1 Fane 1: Simulering & Analyse
   - 4.2 Fane 2: Dataimport & Versjoner
   - 4.3 Fane 3: Datamodell (Innsyn)
   - 4.4 Fane 4: Endringslogg
5. [Dataark — oversikt](#5-dataark--oversikt)
6. [Product Master](#6-product-master)
7. [Locations](#7-locations)
8. [Work Centers](#8-work-centers)
9. [Operation Master](#9-operation-master)
10. [Item Costs](#10-item-costs)
11. [BOM (Stykkliste)](#11-bom-stykkliste)
12. [Routing](#12-routing)
13. [By Product Rules](#13-by-product-rules)
14. [Capacity Calendar](#14-capacity-calendar)
15. [Production Scenario](#15-production-scenario)
16. [Beregningsmodell](#16-beregningsmodell)
17. [Eksempel — Kodal Hovleri](#17-eksempel--kodal-hovleri)
18. [Import av data med Excel](#18-import-av-data-med-excel)
19. [Eksport av PDF-rapport](#19-eksport-av-pdf-rapport)
20. [Eksport av Excel-rapport](#20-eksport-av-excel-rapport)
21. [JSON-eksport](#21-json-eksport)
22. [Tilpasning og vedlikehold](#22-tilpasning-og-vedlikehold)
23. [Kommandolinje-verktøy](#23-kommandolinje-verktøy)

---

## 1. Målgrupper og roller

Systemet brukes av flere roller i virksomheten. Tabellen under forklarer hva hver rolle gjør og hvordan.

| Rolle | Ansvarsområde | Hvordan de bruker systemet |
|-------|--------------|---------------------------|
| **Produksjonsleder / Drift** | Simulere "what-if"-scenarioer for produksjonskost | Åpner Marimo web-app, justerer parametere (timelønn, svinn, batch-størrelse), kjører simulering, eksporterer PDF-rapport til ledelsen |
| **Økonomi / Controller** | Vedlikeholde kostpriser, kostsatser og produktstruktur | Redigerer Excel-datamodellen direkte (Item Costs, Work Centers, BOM), laster opp via Marimo, eksporterer Excel til Power BI |
| **Innkjøper** | Oppdatere råvarepriser og vurdere innkjøpsstrategi | Oppdaterer Unit Cost i Item Costs-arket, importerer via Marimo, ser konsekvensen av prisendringer gjennom simulering |
| **Kvalitetsansvarlig** | Justere svinnprosenter og co-produkt-andeler | Endrer Scrap % og Co-Prod % i BOM-arket, simulerer effekt på netto produksjonskost |
| **IT / Superbruker** | Administrere databasen, versjonshistorikk, feilsøking | Bruker kommandolinje-verktøy (`data_repo.py`, `excel_bridge.py`), overvåker endringslogg, gjeninnlaster tidligere versjoner ved behov |
| **Ledelse** | Motta rapporter og ta beslutninger | Leser PDF-rapporter (ledelsessammendrag + nøkkeltall) og Excel-analyse fra økonomiavdelingen |

### Arbeidsflyt — typisk månedlig syklus

```
Uke 1: Innkjøper oppdaterer råvarepriser → laster opp Excel
Uke 2: Kvalitet justerer svinn-prosenter basert på erfaringstall
Uke 3: Produksjonsleder kjører simuleringer for kommende måneds produksjon
Uke 4: Controller eksporterer PDF-rapport med ledelsessammendrag
```

---

## 2. Systemarkitektur

### 2.1 Komponenter

Systemet består av fire hovedlag:

```
┌─────────────────────────────────────────────────────────────────────┐
│                      Excel-datamodell (.xlsx)                       │
│   Product Master  ·  Work Centers  ·  BOM  ·  Routing  ·  ...      │
│   (10 ark med stamdata)                                             │
└───────────────────────────────────┬─────────────────────────────────┘
                                    │ import
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│              SQLite-database (src/produksjonskalkyle.db)             │
│   DataRepo — CRUD · endringslogg · versjonerte opplastede filer    │
│   change_log · uploaded_files · products · bom_lines · ...          │
└───────────────────────────────────┬─────────────────────────────────┘
                                    │ les
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    Python-beregningsmotor (src/)                     │
│   CostCalculator  (materialkost, operasjonskost, setupkost)         │
│   SimulationEngine  (what-if-sammenligning)                         │
│   SimulationOverride  (overstyrte parametere)                       │
└───────────────────────────────┬─────────────────────────────────────┘
                                │
            ┌───────────────────┼───────────────────────┐
            ▼                   ▼                       ▼
┌───────────────────────┐ ┌────────────┐ ┌───────────────────────────┐
│   Marimo web-app      │ │   JSON     │ │   PDF / Excel-rapporter   │
│   src/varekost_app.py │ │  (eksport) │ │   src/generer_pdf_rapport │
│   Simulering + import  │ │            │ │   src/generer_excel_rapp │
└───────────────────────┘ └────────────┘ └───────────────────────────┘
```

### 2.2 Filer og deres rolle

| Fil | Formål |
|-----|--------|
| `src/kostberegning.py` | **Kjernelogikk:** ExcelData-parsing, SqliteData (databaseavlesning), CostCalculator (kostnadskalkyle), SimulationEngine (what-if-sammenligning), SimulationOverride, export_product_costs_to_json |
| `src/data_repo.py` | SQLite-databasehåndtering: DataRepo-klasse med CRUD, endringslogg, versjonering av opplastede filer. Database: `src/produksjonskalkyle.db` |
| `src/excel_bridge.py` | Import/eksport mellom Excel og SQLite. validate_excel() (validering), import_excel_to_sqlite() (import), export_sqlite_to_excel() (eksport til 11 ark inkl. endringslogg) |
| `src/varekost_app.py` | **Marimo web-app** med 4 faner: simulering & analyse, dataimport & versjoner, datamodell-innsyn, endringslogg |
| `src/generer_pdf_rapport.py` | PDF-rapportgenerering med ReportLab. Ledelsessammendrag, detaljert kostnadsoversikt, kapasitetssammendrag, dokumentasjon-PDF |
| `src/generer_excel_rapport.py` | Excel-rapportgenerering for simuleringsresultater (med logo, tabeller, fargeprofil) |
| `src/scripts/lag_testdata_v3.py` | Generering av testdata-Excel (`src/Produksjonsmodell_Testdata_v3.xlsx`) |
| `src/scripts/lag_baseline.py` | Beregning av baseline-kalkyle fra kommandolinje (til `output/`) |
| `src/scripts/sjekk_diff.py` | Sammenligning av to kjøringer |
| `src/scripts/oppdater_mal.py` | Oppdatering av Excel-mal (`src/Produksjonsmodell_Mal.xlsx`) |
| `docs/STYLING.md` | Fargepalett, typografi, CSS-klasser for Marimo-app, PDF og Excel |

---

## 3. Dataflyt

### 3.1 Hovedflyt

1. **Excel-import:** Bruker laster opp Excel-fil i Marimo-appen (eller via CLI)
   - `validate_excel()` sjekker ark- og kolonne-strukturer, kryssreferanser og datatyper
   - `import_excel_to_sqlite()` lagrer data i SQLite med felt-for-felt endringslogg
   - Original Excel-fil lagres som BLOB i `uploaded_files`-tabellen for versjonering

2. **SQLite → Python-objekter:** `SqliteData`-klassen leser SQLite-tabeller og bygger strukturerte lister:
   - `Product`, `WorkCenter`, `BOMLine`, `RoutingLine`, `ByProductRule`, etc.

3. **Python-beregning:** `CostCalculator` utfører kalkyle i denne rekkefølgen:
   - Materialkost (fra BOM + Item Costs)
   - Operasjonskost (fra Routing + Work Centers)
   - Setupkost (fordelt på batch-størrelse)
   - Brutto produksjonskost (sum)
   - Biproduktverdi (fra By Product Rules)
   - Netto produksjonskost (brutto - biproduktverdi)
   - **Dynamisk cost roll-up:** Hvis en BOM-komponent er en ferdigvare, brukes dens dynamiske brutto produksjonskost i stedet for statisk Item Cost

4. **Simulering:** `SimulationOverride` inneholder overstyrte parametere → `SimulationEngine.compare_all()` sammenligner original vs simulert

5. **Eksport:** PDF (ledelsessammendrag + detaljer), Excel (simuleringsresultater), JSON, komplett database-Excel (11 ark)

### 3.2 Dataflyt — visuell

```
Excel (.xlsx) ── import ──→ SQLite (src/produksjonskalkyle.db) ──→ Python-dataobjekter ──→ Kostnadsberegning
                                   ↑                            ↓
                          Marimo-app (src/varekost_app.py)    JSON / PDF / Excel-eksport
```

---

## 4. Marimo web-app — full gjennomgang

Marimo web-appen (`src/varekost_app.py`) er hovedgrensesnittet for sluttbrukere. Den har 4 faner, tilgjengelig via toppen av skjermen.

Start appen med:
```bash
marimo run src/varekost_app.py
```

### 4.1 Fane 1: Simulering & Analyse

Dette er hovedfanen for å justere parametere, kjøre simuleringer og eksportere resultater.

#### Filter
- **🔍 Filtrer på varenummer og beskrivelse:** Tekstfelt i sidebaren. Skriver du f.eks. "Panel" vises kun produkter med "Panel" i navnet, og tilhørende BOM, routing og arbeidssentre filtreres automatisk (kaskade-filtrering).

#### Trinn 1: Juster parametere
Fire accordion-seksjoner (kan utvides enkeltvis):

| Seksjon | Hva du kan endre | Eksempel |
|---------|------------------|----------|
| **🪵 Råvarepriser** | Unit Cost for råvarer | Endre pris på Skrulast fra 3 000 → 3 500 kr/M3 |
| **🗑️ Svinn- og kapp-prosenter** | Scrap % og Co-Prod % per BOM-linje | Øk svinn fra 5 % → 7 % |
| **🏭 Arbeidssentre (timekostnad)** | Lønn, maskin, overhead per time, effektiv kapasitet % | Øk timelønn fra 550 → 600 kr |
| **📋 Routing (stykkpris/tider)** | Run time, setup time, batch size per operasjon | Reduser run time fra 0,15 → 0,12 min |

Endringer lagres automatisk når du skriver i tabellene. Verdier som er ulik original vises, og fjernes automatisk når du setter dem tilbake til originalverdi.

#### Trinn 2: Planlagt kvantum
- **📦 Planlagt kvantum:** Angi antall enheter du planlegger å produsere (f.eks. 1 000 stk). Høyere kvantum gir lavere oppstartskostnad per enhet.

#### Trinn 3: Kjør simulering
- **⚡ Start simulering:** Trykk knappen for å sammenligne originale kostnader med simulerte verdier.

#### Resultater

Etter simulering vises:

1. **KPI-kort** øverst:
   - Totalt antall simulerte produkter
   - Samlet kostnadsendring i NOK og prosent
   - Totalt timebehov

2. **Sammenligningstabell** — viser original vs simulert per produkt for:
   - Materialkost
   - Operasjonskost
   - Oppstartkost (setup)
   - Brutto produksjonskost
   - Biproduktverdi
   - Netto produksjonskost (med diff i NOK)

3. **Scenariototaler** — viser total netto kost, kost per enhet og timebehov per produkt

#### Trinn 4: Eksporter
Se egne kapitler for [PDF-eksport](#19-eksport-av-pdf-rapport) og [Excel-eksport](#20-eksport-av-excel-rapport).

### 4.2 Fane 2: Dataimport & Versjoner

Denne fanen håndterer import av Excel-data og administrasjon av versjonshistorikk.

#### Last opp Excel
1. **📄 Velg Excel-fil:** Klikk for å velge en .xlsx-fil
2. **Hva er endret? (valgfri kommentar):** Skriv f.eks. "Oppdaterte råvarepriser Q3"
3. Systemet validerer automatisk:
   - Sjekker at alle påkrevde ark finnes
   - Sjekker kolonne-strukturer
   - Sjekker kryssreferanser (f.eks. at alle Work Center-koder finnes i Work Centers-arket)
   - Rapporterer antall rader, nye produkter, fjernede produkter
   - Ved feil: vises detaljert feilliste — ingenting importeres
4. Ved godkjent validering: data importeres til SQLite, endringslogg oppdateres

#### Versjonshistorikk
- **📜 Velg tidligere versjon:** Dropdown med de 10 siste importene (tidspunkt, filnavn, kommentar)
- Velg en tidligere versjon for å gjeninnlaste den — alle nåværende data erstattes

#### Eksporter komplett datafil
- **📥 Last ned komplett datafil:** Eksporterer hele SQLite-databasen til Excel med 11 ark (alle 10 dataark + endringslogg). Inkluderer ACTION- og Rad ID-kolonner for roundtrip-redigering.

### 4.3 Fane 3: Datamodell (Innsyn)

Viser alle data i databasen som skrivebeskyttede tabeller (filtrert etter aktivt filter i sidebaren):
- Product Master, Locations, Work Centers, Operations, Item Costs, BOM, Routing, By Product Rules, Capacity Calendar, Production Scenario

Perfekt for rask innsyn uten å åpne Excel.

### 4.4 Fane 4: Endringslogg

Viser de 50 siste endringene i databasen, felt-for-felt:
- Tidspunkt, bruker, kilde, tabell, nøkkel, feltnavn, gammel verdi, ny verdi

Alle endringer (import, oppdatering, sletting) loggføres automatisk.

---

## 5. Dataark — oversikt

Modellen består av **10 ark** i Excel, som speiles i SQLite-tabeller:

| # | Arknavn | Innhold | Nøkkelkolonner |
|---|---------|---------|----------------|
| 1 | **Product Master** | Vareregister | Item No, Description, Item Type |
| 2 | **Locations** | Fabrikker og lagre | Location Code, Location Name |
| 3 | **Work Centers** | Arbeidssentre med kostsatser | Work Center Code, Labor/Machine/Overhead Cost |
| 4 | **Operation Master** | Standardoperasjoner | Operation Code, Description |
| 5 | **Item Costs** | Kostpriser per vare | Item No, Unit Cost, Currency |
| 6 | **BOM** | Stykkliste (hva består produktet av) | Parent Item, Component, Quantity Per |
| 7 | **Routing** | Produksjonsflyt (operasjoner, tider) | Item No, Operation No, Setup/Run Time |
| 8 | **By Product Rules** | Biprodukter og verdsetting | Parent Item, By Product, Market Value |
| 9 | **Capacity Calendar** | Kapasitetskalender per arbeidssenter | Work Center, Date, Available Hours |
| 10 | **Production Scenario** | Forhåndsdefinerte scenarioer | Scenario Name, Product, Planned Quantity |

---

## 6. Product Master

### 6.1 Formål

Register over alle varer i virksomheten. Dette er hovedkatalogen over alt som finnes — råvarer, halvfabrikata, ferdigvarer, biprodukter og handelsvarer.

### 6.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Item No** | Tekst | Unik identifikator for varen | `RM001`, `FG001`, `BP001` |
| **Description** | Tekst | Beskrivende navn på varen | `Skrulast 48x198`, `Utvendig Panel 21x95` |
| **Item Type** | Tekst | Type vare | `Raw Material`, `Semi Finished`, `Finished Good`, `By Product`, `Trading Item` |
| **Product Group** | Tekst | Gruppering av varer | `Skrulast`, `Panel`, `Kledning`, `Spon` |
| **Base Unit of Measure** | Tekst | Standard måleenhet | `LM`, `M3`, `KG`, `PCS` |
| **Active** | Ja/Nei | Angir om varen er aktiv | `Ja` |

### 6.3 Varetyper

```
Raw Material    → Råvare som kjøpes inn (f.eks. skrulast, maling)
Semi Finished   → Halvfabrikat (mellomprodukt)
Finished Good   → Ferdigvare som selges (f.eks. panel, terrassebord)
By Product      → Biprodukt som oppstår i produksjon (f.eks. hovelspon, flis)
Trading Item    → Handelsvare (kjøpes og selges uendret)
```

### 6.4 Testdata — Kodal Hovleri

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

## 7. Locations

### 7.1 Formål

Register over fabrikker og lokasjoner. Hvert arbeidssenter er knyttet til en lokasjon.

### 7.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Location Code** | Tekst | Unik kode for lokasjonen | `KOD` |
| **Location Name** | Tekst | Navn på lokasjonen | `Kodal Fabrikk` |
| **Location Type** | Tekst | Type lokasjon | `Factory`, `Warehouse`, `Distribution Center`, `Sales Office` |
| **Active** | Ja/Nei | Angir om lokasjonen er aktiv | `Ja` |

### 7.3 Testdata

| Location Code | Location Name | Location Type | Aktiv |
|---------------|---------------|---------------|-------|
| KOD | Kodal Fabrikk | Factory | Ja |
| SKI | Skien Lager | Warehouse | Ja |

---

## 8. Work Centers

### 8.1 Formål

Register over produksjonsressurser — maskiner og arbeidsplasser. Hvert arbeidssenter har timekostnader som brukes til å beregne operasjonskost.

### 8.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Work Center Code** | Tekst | Unik identifikator | `HOVEDHOVEL` |
| **Description** | Tekst | Beskrivende navn | `Hovedhovel` |
| **Location Code** | Tekst | Fabrikken det tilhører | `KOD` |
| **Labor Cost per Hour** | Desimal | Arbeidskostnad per time (lønn, arbeidsgiveravgift, pensjon, feriepenger) | `550` |
| **Machine Cost per Hour** | Desimal | Maskinkostnad per time (avskrivninger, service, leasing, vedlikehold, energi) | `900` |
| **Overhead Cost per Hour** | Desimal | Indirekte produksjonskostnader (produksjonsledelse, kvalitet, vedlikeholdsadm., intern logistikk) | `150` |
| **Capacity Hours per Day** | Desimal | Tilgjengelige timer per dag | `16` |
| **Effective Capacity %** | Prosent | Hvor stor del av tiden som faktisk kan brukes til produksjon (tar hensyn til stopp, vedlikehold, feil) | `85` |
| **Active** | Ja/Nei | Angir om arbeidssenteret er aktivt | `Ja` |

### 8.3 Timekostnad (Justert for effektiv kapasitet)

For å sikre at alle kalkyler tar høyde for uunngåelig ståtid, mikrostopp og planlagt vedlikehold, justeres den timeprisen som belastes produktet opp med maskinens effektivitetsgrad (**Effective Capacity %**):

$$\text{Effektiv Timepris} = \frac{\text{Nominell Timepris (Lønn + Maskin + Overhead)}}{\frac{\text{Effective Capacity \%}}{100}}$$

#### Eksempel — HOVEDHOVEL (85 % effektivitet):
- Nominell timepris = $550 \text{ kr (lønn)} + 900 \text{ kr (maskin)} + 150 \text{ kr (overhead)} = 1600 \text{ kr/time}$
- Effektiv timepris = $\frac{1600 \text{ kr}}{0,85} = \mathbf{1882,35 \text{ kr/time}}$

Dette betyr at produktet belastes 1882,35 kr per time i stedet for 1600 kr. De 15 % med tapt produksjonstid blir automatisk og nøyaktig bakt inn i produktkalkylen.

### 8.4 Effektiv kapasitet

```
Effektive timer per dag = Capacity Hours × Effective Capacity %

Eksempel — HOVEDHOVEL:
  16 timer × 85 % = 13,6 effektive timer
```

### 8.5 Testdata

| Work Center | Beskrivelse | Lokasjon | Labor | Maskin | Overhead | Timer/dag | Eff. % | Aktiv |
|-------------|-------------|----------|-------|--------|----------|-----------|--------|-------|
| HOVEDHOVEL | Hovedhovel | KOD | 550 | 900 | 150 | 16 | 85 | Ja |
| SPESIALHOVEL | Spesialhovel | KOD | 550 | 950 | 150 | 16 | 85 | Ja |
| MALINGSLINJE | Malingslinje | KOD | 500 | 400 | 120 | 16 | 80 | Ja |
| PAKKELINJE | Pakkelinje | KOD | 450 | 300 | 100 | 8 | 90 | Ja |

---

## 9. Operation Master

### 9.1 Formål

Standardisert liste over operasjoner som kan brukes i routing. Gir en felles "ordbok" for produksjonsprosesser.

### 9.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Operation Code** | Tekst | Unik operasjonskode | `RIP`, `PLANING`, `PROFILE` |
| **Description** | Tekst | Beskrivelse av operasjonen | `Oppdeling`, `Hovling`, `Profilering` |
| **Default Work Center** | Tekst | Anbefalt arbeidssenter | `HOVEDHOVEL` |
| **Standard Unit** | Tekst | Måleenhet for produksjonstid | `Minutes`, `Hours` |
| **Active** | Ja/Nei | Angir om operasjonen er aktiv | `Ja` |

### 9.3 Testdata

| Operation Code | Description | Default Work Center | Standard Unit | Aktiv |
|----------------|-------------|---------------------|---------------|-------|
| RIP | Oppdeling | HOVEDHOVEL | Minutes | Ja |
| PLANING | Hovling | HOVEDHOVEL | Minutes | Ja |
| PROFILE | Profilering | SPESIALHOVEL | Minutes | Ja |
| MALING | Maling | MALINGSLINJE | Minutes | Ja |
| PACKING | Pakking | PAKKELINJE | Minutes | Ja |

---

## 10. Item Costs

### 10.1 Formål

Samlet register over kostpriser for alle varer. For råvarer er dette innkjøpspris. For ferdigvarer settes prisen til 0 (beregnes automatisk). For biprodukter er dette markedsverdi.

### 10.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Item No** | Tekst | Referanse til varen (fra Product Master) | `RM001` |
| **Cost Type** | Tekst | Type kostpris | `Standard Cost`, `Last Direct Cost`, `Forecast Cost`, `Budget Cost` |
| **Unit Cost** | Desimal | Kostpris per enhet | `3000.00` |
| **Currency** | Tekst | Valuta | `NOK`, `EUR` |
| **Effective Date** | Dato | Dato kostprisen gjelder fra | `2026-01-01` |

### 10.3 Testdata

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

> **Merk:** Ferdigvarer (FG001–FG004) har kostpris 0,00 fordi kostnaden beregnes automatisk fra BOM og Routing.

---

## 11. BOM (Stykkliste)

### 11.1 Formål

Beskriver hvilke komponenter som inngår i et produkt. En BOM-linje sier: "For å lage X trenger du Y, og du får Z enheter output per enhet input."

### 11.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Parent Item No** | Tekst | Produktet som produseres | `FG001` |
| **Component Item No** | Tekst | Komponenten som forbrukes | `RM001` |
| **Quantity Per** | Desimal | Antall output-enheter per input-enhet | `400` |
| **Unit of Measure** | Tekst | Måleenhet for forholdet | `LM` |
| **Scrap %** | Prosent | Forventet materialsvinn | `5.0` |
| **Co-Prod %** | Prosent | Andel samprodukt (co-product / B-vare) | `6.0` |
| **Co-Prod Item No** | Tekst | Varenummer for samproduktet | `JD16073-B` |
| **Valid From** | Dato | Gyldig fra dato | `2026-01-01` |
| **Valid To** | Dato | Gyldig til dato (tom = alltid gyldig) | |

### 11.3 Hvordan Quantity Per fungerer

```
Quantity Per = antall output-enheter per input-enhet

Forbruk per output = 1 / Quantity Per

Eksempel — FG001 (Panel) fra RM001 (Skrulast):
  Quantity Per = 400 LM per M3
  Forbruk per LM panel = 1 / 400 = 0,0025 M3 per LM
```

### 11.4 Materialkost-beregning

```
Materialkost per enhet = (Unit Cost / Quantity Per) × (1 + Scrap% / 100)

Eksempel — FG001:
  = (3 000 kr/M3 / 400 LM/M3) × (1 + 0,05)
  = 7,50 × 1,05
  = 7,875 kr/LM
```

### 11.5 Testdata

| Parent | Component | Quantity Per | UOM | Scrap % | Co-Prod % | Co-Prod Item | Valid From |
|--------|-----------|-------------|-----|---------|-----------|--------------|------------|
| FG001 | RM001 | 400 | LM | 5,0 | 0 | | 2026-01-01 |
| FG002 | RM002 | 250 | LM | 4,0 | 0 | | 2026-01-01 |
| FG003 | RM001 | 420 | LM | 6,0 | 0 | | 2026-01-01 |
| FG004 | FG001 | 1 | LM | 2,0 | 0 | | 2026-01-01 |
| FG004 | RM003 | 20 | LTR | 3,0 | 0 | | 2026-01-01 |

> **Merk:** FG004 (malt panel) har to BOM-linjer: den bruker FG001 (ubehandlet panel) som komponent i tillegg til maling. Dette kalles en **produksjonskjede** — FG004 bygger på FG001.

---

## 12. Routing

### 12.1 Formål

Beskriver produksjonsprosessen — hvilke operasjoner som utføres, i hvilken rekkefølge, på hvilket arbeidssenter, og hvor lang tid hver operasjon tar.

### 12.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Item No** | Tekst | Produkt som produseres | `FG001` |
| **Operation No** | Heltall | Sekvensnummer (stigende rekkefølge) | `10`, `20`, `30` |
| **Operation Code** | Tekst | Hvilken operasjon som utføres | `RIP`, `PLANING` |
| **Work Center Code** | Tekst | Arbeidssenter som utfører operasjonen | `HOVEDHOVEL` |
| **Setup Time Minutes** | Desimal | Tid til klargjøring (omstilling, knivbytte, innkjøring, kontrollmåling) | `15.0` |
| **Run Time Minutes** | Desimal | Produksjonstid per enhet | `0.15` |
| **Batch Size** | Desimal | Normal ordrestørrelse (brukes til å fordele setupkostnad) | `500` |
| **Valid From** | Dato | Gyldig fra dato | `2026-01-01` |
| **Valid To** | Dato | Gyldig til dato | |

### 12.3 Operasjonskost-beregning

```
Kjøretid per enhet (timer) = Run Time Minutes / 60
Kjørekost per enhet = Kjøretid × Timekost

Setupkost per enhet = (Setup Time / 60 × Timekost) / Batch Size

Eksempel — Op 10 (Oppdeling) for FG001:
  Kjøretid: 0,15 min / 60 = 0,0025 timer
  Kjørekost: 0,0025 × 1 600 = 4,00 kr/LM
  Setupkost: (15/60 × 1 600) / 500 = 400 / 500 = 0,80 kr/LM
```

### 12.4 Testdata — FG001 (Utvendig Panel 21×95)

| Op.nr | Operasjon | Arbeidssenter | Setup (min) | Kjøretid (min) | Batch |
|-------|-----------|---------------|-------------|----------------|-------|
| 10 | Oppdeling | HOVEDHOVEL | 15,0 | 0,15 | 500 |
| 20 | Hovling | HOVEDHOVEL | 10,0 | 0,10 | 500 |
| 30 | Profilering | SPESIALHOVEL | 20,0 | 0,12 | 500 |
| 40 | Pakking | PAKKELINJE | 5,0 | 0,05 | 500 |

### 12.5 Produksjonsflyt — visuell

```
FG001 - Utvendig Panel 21×95
-----------------------------
  Op 10: RIP       @ HOVEDHOVEL    (15 min setup + 0,15 min/LM)
  Op 20: PLANING   @ HOVEDHOVEL    (10 min setup + 0,10 min/LM)
  Op 30: PROFILE   @ SPESIALHOVEL  (20 min setup + 0,12 min/LM)
  Op 40: PACKING   @ PAKKELINJE    ( 5 min setup + 0,05 min/LM)

FG002 - Terrassebord 28×120
-----------------------------
  Op 10: RIP       @ HOVEDHOVEL    (12 min setup + 0,18 min/LM)
  Op 20: PLANING   @ HOVEDHOVEL    ( 8 min setup + 0,12 min/LM)
  Op 30: PROFILE   @ SPESIALHOVEL  (15 min setup + 0,15 min/LM)

FG003 - Kledning 18×120
-----------------------------
  Op 10: RIP       @ HOVEDHOVEL    (10 min setup + 0,10 min/LM)
  Op 20: PLANING   @ HOVEDHOVEL    ( 8 min setup + 0,08 min/LM)

FG004 - Utvendig Panel 21×95 - Malt
-----------------------------
  Op 10: MALING    @ MALINGSLINJE  (30 min setup + 0,20 min/LM)
  Op 20: PACKING   @ PAKKELINJE    ( 5 min setup + 0,05 min/LM)
```

---

## 13. By Product Rules

### 13.1 Formål

Beskriver hvordan biprodukter håndteres økonomisk. I trelastproduksjon oppstår det alltid biprodukter som hovelspon, flis og bark. Disse har en verdi som skal trekkes fra hovedproduktets kostnad.

### 13.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Parent Item No** | Tekst | Produktet (ferdigvaren) som skaper biproduktet | `FG001` |
| **By Product Item No** | Tekst | Biproduktet | `BP001` |
| **Expected Quantity** | Desimal | Forventet mengde biprodukt per enhet hovedprodukt | `0.5` |
| **Unit of Measure** | Tekst | Måleenhet for biproduktet | `KG` |
| **Market Value** | Desimal | Forventet markedspris per enhet | `1.50` |
| **Allocation Method** | Tekst | Hvordan verdien skal håndteres | `Reduce Main Product Cost` |

### 13.3 Allokeringsmetoder

| Metode | Beskrivelse |
|--------|-------------|
| **Reduce Main Product Cost** | Verdien trekkes fra hovedproduktets kost (anbefalt) |
| **Separate Profit Center** | Biproduktet behandles som eget salgbart produkt |
| **Informational Only** | Brukes kun til rapportering, påvirker ikke kost |

### 13.4 Biproduktverdi-beregning

```
Biproduktverdi per enhet = Expected Quantity × Market Value

Eksempel — FG001:
  BP001 (Hovelspon): 0,5 KG × 1,50 kr/KG = 0,75 kr/LM
  BP002 (Flis):      0,3 KG × 0,80 kr/KG = 0,24 kr/LM
  Total:             0,99 kr/LM
```

### 13.5 Testdata

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

### 13.6 Sortering og 2. sortering (A- og B-vare)

I trelastproduksjon (høvlerier) opplever man ofte at en viss andel av produksjonen ikke tilfredsstiller kravene til 1. sortering (A-vare), og blir derfor nedgradert til 2. sortering (B-vare).

B-varen bruker akkurat det samme råstoffet (tømmeret) og den samme maskintiden som A-varen, men selges til en lavere markedspris. Dette modelleres som et samprodukt (co-product) i BOM-arket via feltene **Co-Prod %** og **Co-Prod Item No**.

#### Matematisk eksempel (94 % A-vare og 6 % B-vare):

1. **Brutto råvareforbruk per meter A-vare øker**:
   Når bare 94 % av tømmeret blir til A-vare, må vi kjøre mer råstoff gjennom høvelen per ferdige meter A-vare.
   Hvis det teoretiske utbyttet av tømmeret er **553,34 meter/m³**, blir det reelle utbyttet av A-vare:
   $$553,34 \text{ meter/m³} \times 0,94 = \mathbf{520,14 \text{ meter A-vare per m³}}$$
   Dette tallet legges inn i **BOM** under `Quantity Per` for A-varen.
   Råvarekosten per meter A-vare stiger fra **5,12 kr** til **5,44 kr** (fordi vi må fordele råvarekostnaden på færre godkjente meter).

2. **B-vare får egen kalkyle**: B-varen får en egen Product Master-post, BOM-linje (med Quantity Per justert for B-vare-andel) og routing, slik at systemet beregner en fullstendig produksjonskostnad for B-varen isolert.

3. **Justert kalkyle for A-vare**: Co-Prod %-feltet i BOM justerer run-time marginalt opp for A-varen, og materialkostnaden fordeles proporsjonalt. Bidraget fra B-varen fremkommer som egen kalkyle under co-product results.

Denne metoden er revisorgodkjent og sikrer et nøyaktig bilde av lønnsomheten, samtidig som den tar høyde for at man ikke ender opp med kun A-vare på lager under produksjon.

---

## 14. Capacity Calendar

### 14.1 Formål

Register over tilgjengelig kapasitet per arbeidssenter per dag. Brukes til å planlegge produksjon og analysere kapasitetsutnyttelse.

### 14.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Work Center** | Tekst | Arbeidssenter (ref. Work Centers) | `HOVEDHOVEL` |
| **Date** | Dato | Dato | `2026-03-16` |
| **Available Hours** | Desimal | Tilgjengelige timer for dagen | `16` |
| **Planned Downtime** | Desimal | Planlagte stopp i timer (vedlikehold, ferie, ombygging) | `4` |

### 14.3 Testdata

| Work Center | Date | Available Hours | Planned Downtime |
|-------------|------|----------------|-----------------|
| HOVEDHOVEL | 2026-03-16 | 16 | 0 |
| HOVEDHOVEL | 2026-03-17 | 16 | 4 |

---

## 15. Production Scenario

### 15.1 Formål

Forhåndsdefinerte produksjonsscenarioer som kan brukes som utgangspunkt for simulering.

### 15.2 Felter

| Kolonne | Type | Beskrivelse | Eksempel |
|---------|------|-------------|----------|
| **Scenario Name** | Tekst | Navn på scenario | `Normal Produksjon` |
| **Product** | Tekst | Produktet som simuleres (Item No) | `FG001` |
| **Planned Quantity** | Desimal | Planlagt produksjonsmengde | `100000` |
| **Start Date** | Dato | Startdato for scenario | `2026-01-01` |
| **End Date** | Dato | Sluttdato for scenario | `2026-12-31` |

### 15.3 Testdata

| Scenario Name | Product | Planned Quantity | Start Date | End Date |
|---------------|---------|-----------------|------------|----------|
| Normal Produksjon | FG001 | 100000 | 2026-01-01 | 2026-12-31 |
| Normal Produksjon | FG002 | 50000 | 2026-01-01 | 2026-12-31 |

---

## 16. Beregningsmodell

### 16.1 Materialkost

```
Materialkost = (Unit Cost / Quantity Per) × (1 + Scrap% / 100)

Hvor:
  Unit Cost    = enhetskost for råvaren (fra Item Costs)
  Quantity Per = antall output-enheter per input-enhet (fra BOM)
  Scrap%       = forventet materialsvinn (fra BOM)

Eksempel — FG001 (Panel) fra RM001 (Skrulast):
  = (3 000 kr/M3 / 400 LM/M3) × (1 + 5/100)
  = 7,50 × 1,05
  = 7,875 kr/LM
```

### 16.2 Operasjonskost

```
Operasjonskost per enhet = (Run Time Minutes / 60) × Timekost

Hvor:
  Run Time Minutes = produksjonstid per enhet (fra Routing)
  Timekost         = Labor + Machine + Overhead (fra Work Centers)

Eksempel — Op 10 (Oppdeling) for FG001:
  Kjøretid: 0,15 min / 60 = 0,0025 timer
  Timekost: 550 + 900 + 150 = 1 600 kr/time
  = 0,0025 × 1 600 = 4,00 kr/LM
```

### 16.3 Setupkost

```
Setupkost per enhet = ((Setup Time Minutes / 60) × Timekost) / Batch Size

Hvor:
  Setup Time Minutes = klargjøringstid per ordre (fra Routing)
  Timekost           = Labor + Machine + Overhead (fra Work Centers)
  Batch Size         = normal ordrestørrelse (fra Routing)

Eksempel — Op 10 (Oppdeling) for FG001:
  Setup-tid: 15 min / 60 = 0,25 timer
  Timekost: 1 600 kr/time
  Batch: 500 LM
  = (0,25 × 1 600) / 500 = 400 / 500 = 0,80 kr/LM
```

### 16.4 Brutto produksjonskost

```
Brutto produksjonskost = Materialkost + Operasjonskost + Setupkost

Eksempel — FG001:
  = 7,875 + 10,675 + 2,575
  = 21,125 kr/LM
```

### 16.5 Biproduktverdi

```
Biproduktverdi per enhet = Sum (Expected Quantity × Market Value)

Eksempel — FG001:
  BP001 (Hovelspon): 0,5 KG × 1,50 kr/KG = 0,750 kr/LM
  BP002 (Flis):      0,3 KG × 0,80 kr/KG = 0,240 kr/LM
  Total:             0,990 kr/LM
```

### 16.6 Netto produksjonskost

```
Netto produksjonskost = Brutto produksjonskost - Biproduktverdi

Eksempel — FG001:
  = 21,125 - 0,990
  = 20,135 kr/LM
```

### 16.7 Oppsummering — FG001

```
+-----------------------------------------------------------+
|  Materialkost:              7,8750 kr/LM                  |
|    RM001: Skrulast 48×198                                 |
|      Pris: 3 000,00 kr/M3                                 |
|      Output per input: 400 LM                             |
|      Svinn: 5 %  →  7,8750 kr                            |
+-----------------------------------------------------------+
|  Operasjonskost:          10,6750 kr/LM                   |
|  Setupkost:                2,5750 kr/LM                   |
|    Op 10: Oppdeling @ HOVEDHOVEL                          |
|      Kjøretid: 0,15 min  |  Setup: 15 min / batch 500    |
|      Timekost: 1 600 kr/t  →  4,8000 kr                  |
|    Op 20: Hovling @ HOVEDHOVEL                            |
|      Kjøretid: 0,10 min  |  Setup: 10 min / batch 500    |
|      Timekost: 1 600 kr/t  →  3,2000 kr                  |
|    Op 30: Profilering @ SPESIALHOVEL                      |
|      Kjøretid: 0,12 min  |  Setup: 20 min / batch 500    |
|      Timekost: 1 650 kr/t  →  4,4000 kr                  |
|    Op 40: Pakking @ PAKKELINJE                            |
|      Kjøretid: 0,05 min  |  Setup: 5 min / batch 500     |
|      Timekost: 850 kr/t  →  0,8500 kr                    |
+-----------------------------------------------------------+
|  Brutto produksjonskost:    21,1250 kr/LM                 |
+-----------------------------------------------------------+
|  Biproduktverdi:            0,9900 kr/LM                  |
|    BP001: Hovelspon  0,5 KG × 1,50 kr = 0,7500 kr       |
|    BP002: Flis       0,3 KG × 0,80 kr = 0,2400 kr       |
+-----------------------------------------------------------+
|  NETTO PRODUKSJONSKOST:    20,1350 kr/LM                  |
+-----------------------------------------------------------+
```

---

## 17. Eksempel — Kodal Hovleri

### 17.1 Produktkalkyle for alle produkter

#### FG001 — Utvendig Panel 21×95

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (RM001: Skrulast) | 7,8750 |
| Operasjonskost (4 operasjoner) | 10,6750 |
| Setupkost | 2,5750 |
| **Brutto produksjonskost** | **21,1250** |
| Biproduktverdi (spon + flis) | -0,9900 |
| **Netto produksjonskost** | **20,1350** |

#### FG002 — Terrassebord 28×120

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (RM002: Gran) | 10,4000 |
| Operasjonskost (3 operasjoner) | 12,1250 |
| Setupkost | 2,3646 |
| **Brutto produksjonskost** | **24,8896** |
| Biproduktverdi (flis + bark) | -0,4200 |
| **Netto produksjonskost** | **24,4696** |

#### FG003 — Kledning 18×120

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (RM001: Skrulast) | 7,5714 |
| Operasjonskost (2 operasjoner) | 4,8000 |
| Setupkost | 0,8000 |
| **Brutto produksjonskost** | **13,1714** |
| Biproduktverdi (spon + flis) | -0,9900 |
| **Netto produksjonskost** | **12,1814** |

#### FG004 — Utvendig Panel 21×95 - Malt

| Kostnadselement | kr/LM |
|----------------|-------|
| Materialkost (FG001 + maling) | 27,7275 |
| Operasjonskost (2 operasjoner) | 4,1083 |
| Setupkost | 1,1617 |
| **Brutto produksjonskost** | **32,9975** |
| Biproduktverdi (spon + flis) | -0,9900 |
| **Netto produksjonskost** | **32,0075** |

---

## 18. Import av data med Excel

### 18.1 Via Marimo web-app (anbefalt)

1. Åpne Marimo-appen: `marimo run src/varekost_app.py`
2. Gå til fanen **📁 Dataimport & Versjoner**
3. Klikk **📄 Velg Excel-fil** og velg din .xlsx-fil
4. Skriv en kommentar i feltet **Hva er endret?** (f.eks. "Oppdaterte råvarepriser Q3")
5. Systemet validerer automatisk:
   - ✓ Sjekker at alle påkrevde ark finnes (10 stk)
   - ✓ Sjekker at alle kolonne-navn er korrekte
   - ✓ Sjekker kryssreferanser (f.eks. at alle Location Code i Work Centers finnes i Locations)
   - ✓ Sjekker datatyper (tall, datoer, Ja/Nei)
   - ✓ Sammenligner med eksisterende data og rapporterer nye/fjernede produkter
   - ❌ Ved feil: vises detaljert feilliste — ingenting importeres!
6. Ved godkjent validering: data importeres til SQLite
   - Hver endring loggføres felt-for-felt i endringsloggen
   - Original Excel-fil lagres som BLOB for versjonshistorikk
   - Gamle overstyringer (simuleringsparametere) nullstilles automatisk

**Valideringsregler ved import:**

| Sjekk | Hva valideres |
|-------|---------------|
| Arkstruktur | Alle 10 ark må finnes. Ekstra ark ignoreres. |
| Kolonner | Hvert ark må ha korrekte kolonne-navn (se dataark-oversikten) |
| Kryssreferanser | Work Center Code → finnes i Work Centers. Component Item No → finnes i Product Master. Osv. |
| Datatyper | Unit Cost må være tall. Active må være Ja/Nei. Effective Date må være dato eller tomt. |
| Unike nøkler | Item No må være unik i Product Master. Work Center Code må være unik i Work Centers. |
| Nye/fjernede produkter | Varsler hvis nye Item No dukker opp, eller hvis tidligere Item No er fjernet. |

### 18.2 Via kommandolinje

```bash
# Importer Excel-data til SQLite (fra src/)
python src/excel_bridge.py --import data.xlsx

# Eksporter SQLite-data til Excel (11 ark inkl. endringslogg)
python src/excel_bridge.py --export utdata.xlsx
```

### 18.3 Versjonshistorikk

Hver import lagres i `uploaded_files`-tabellen i SQLite som BLOB. Du kan:

- **I Marimo:** Velg en tidligere versjon fra dropdown i "Dataimport & Versjoner"-fanen for å gjeninnlaste den
- **I CLI:** Bruk `data_repo.py` for å se versjoner og administrere databasen

```bash
python src/data_repo.py --stats      # Vis tabell-statistikk
python src/data_repo.py --changes    # Vis endringslogg
python src/data_repo.py --clear      # Tøm data (bevar endringslogg)
```

### 18.4 Eksempel på import-workflow

```
1. Åpne Excel-filen og oppdater råvarepriser (Item Costs-arket)
2. Lagre som nytt navn, f.eks. "Kostnadsoppdatering_Q3_2026.xlsx"
3. Åpne Marimo → Dataimport & Versjoner
4. Last opp filen med kommentar "Oppdaterte råvarepriser Q3 2026"
5. Systemet validerer og importerer → "✅ Import fullført!"
6. Gå til Simulering & Analyse → se oppdaterte priser i tabellen
7. Juster eventuelle parametere, kjør simulering, eksporter rapport
```

---

## 19. Eksport av PDF-rapport

### 19.1 Hvordan

1. I Marimo-appen, under fanen **📊 Simulering & Analyse**
2. Kjør en simulering (trykk "⚡ Start simulering")
3. Under **💾 Eksport** → åpne accordionen "📥 Last ned rapporter (PDF / Excel)"
4. Fyll inn valgfritt **Ledelsessammendrag / Kommentar** (vises øverst i PDF)
5. Kryss av **Inkluder tekniske detaljer** hvis du vil ha med material- og operasjonsdetaljer per produkt
6. Trykk **📄 Eksporter til PDF-rapport**
7. Last ned PDF-filen via nedlastingslenken som vises

### 19.2 Hva PDF-rapporten inneholder

| Seksjon | Innhold |
|---------|---------|
| **Forside** | Tittel "Produksjonskost Simulator - Simuleringsrapport", dato, Fram Treindustri-logo |
| **Ledelsessammendrag** | Brukerens egen kommentar (hvis fylt ut) |
| **Kapasitetssammendrag** | Timebehov per arbeidssenter (run-time + setup-time) |
| **Overstyrte parametere** | Liste over hvilke parametere som ble endret i simuleringen (råvarepriser, svinn, timekostnader, routing, co-produkt) |
| **Sammenligningstabell** | Per produkt: original netto, simulert netto, diff i NOK og prosent. Fargekodet (grønn = lavere kostnad, rød = høyere) |
| **Produktdetaljer** (valgfritt) | Hvert produkt med materialdetaljer (komponent, qty per, scrap %, enhetskost, total kost), operasjonsdetaljer (operasjon, arbeidssenter, run-time, batch, kost per time, total kost), biprodukter (biprodukt, kvantum, markedsverdi, total verdi) og co-produkter (produkt, materialkost, operasjonskost, netto) |

### 19.3 Tekniske detaljer

- **Format:** A4, Portrait
- **Farger:** Fram Treindustri-profil (skogsgrønn #14532D, lysegrønn #2F855A)
- **Fonter:** DejaVu (Unicode-støtte) eller Helvetica som fallback
- **Genereres med:** ReportLab via `generer_pdf_rapport.py`

---

## 20. Eksport av Excel-rapport

### 20.1 Simuleringsresultater (via Marimo)

1. I Marimo-appen, under fanen **📊 Simulering & Analyse**
2. Kjør en simulering
3. Under **💾 Eksport** → åpne accordionen
4. Trykk **📊 Eksporter resultater til Excel**
5. Last ned Excel-filen

**Hva Excel-rapporten inneholder:**
- Flere ark med simuleringsresultater
- Sammenligningstabell (original vs simulert per produkt)
- KPI-oversikt
- Detaljer per produkt (material, operasjon, biprodukt)

### 20.2 Komplett database-eksport (via Marimo)

1. Gå til fanen **📁 Dataimport & Versjoner**
2. Trykk **📥 Last ned komplett datafil**
3. Last ned Excel-filen

**Hva Excel-filen inneholder (11 ark):**

| Ark | Innhold |
|-----|---------|
| Product Master | Alle produkter |
| Locations | Alle fabrikker og lagre |
| Work Centers | Alle arbeidssentre med kostsatser |
| Operation Master | Alle standardoperasjoner |
| Item Costs | Alle kostpriser |
| BOM | Alle stykklistelinjer |
| Routing | Alle produksjonsflyt-linjer |
| By Product Rules | Alle biproduktregler |
| Capacity Calendar | Alle kapasitetsdata |
| Production Scenario | Alle scenarioer |
| Change Log | Endringslogg (siste 50 endringer) |

Hvert dataark inkluderer:
- **ACTION-kolonne (kolonne A):** CREATE, UPDATE, DELETE eller tom (for inferering)
- **Rad ID (kolonne B):** Intern database-ID for roundtrip-redigering
- Data-validering (dropdown-lister for Ja/Nei, varetyper, valuta, etc.)
- Kommentarer på header-rader med kolonnebeskrivelser

### 20.3 Via kommandolinje

```bash
# Eksporter hele databasen til Excel (11 ark)
python src/excel_bridge.py --export utdata.xlsx
```

---

## 21. JSON-eksport

### 21.1 produktkalkyle.json

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

### 21.2 Bruk i Power BI

JSON-filen kan lastes direkte inn i Power BI:

1. **Hent data** → **JSON**
2. Velg `produktkalkyle.json`
3. Power BI tolker JSON-strukturen automatisk
4. Bygg rapporter på kost per produkt, produktgruppe, etc.

---

## 22. Tilpasning og vedlikehold

### 22.1 Legge til nytt produkt

1. **Product Master**: Legg til ny rad med Item No, beskrivelse, type
2. **Item Costs**: Legg til kostpris (0 for ferdigvarer)
3. **BOM**: Legg til stykklistelinjer (hva består produktet av?)
4. **Routing**: Legg til produksjonsflyt (operasjoner, tider, arbeidssentre)
5. **By Product Rules**: Legg til biprodukter hvis relevant
6. Last opp Excel-filen i Marimo → dataene er tilgjengelige for simulering

### 22.2 Endre kostsatser

Oppdater `Unit Cost` i **Item Costs** for råvarer, eller `Labor/Machine/Overhead Cost per Hour` i **Work Centers**. Last opp på nytt i Marimo.

### 22.3 Oppdatere testdata-Excel

```bash
python src/scripts/lag_testdata_v3.py
```

Dette genererer `src/Produksjonsmodell_Testdata_v3.xlsx` på nytt med all testdata og beskrivelser.

### 22.4 Overstyringslogikk i Marimo

Når du endrer parametere i Marimo (råvarepriser, svinn, etc.), lagres endringene i et `persistent_overrides`-dict i appens minne:

- Endringer fjernes automatisk når du setter verdien tilbake til original (toleranse: diff < 0,001)
- Overstyringer nullstilles ved ny dataimport
- Overstyringer gjelder kun for gjeldende sesjon — de lagres ikke permanent

### 22.5 Endringslogg

All historikk bevares i `change_log`-tabellen, også etter `clear_all_data()`.

### 22.6 Transportvarer — fler-høvleri-produksjon

Dersom en vare produseres på ett høvleri men viderebearbeides/fraktes til et annet, kan varen flagges som **transportvare** i `transport_flagg`-tabellen i SQLite.

Når flagget settes (`is_transport = 1`), genereres automatisk semi-finished varianter for hver aktiv høvleri-lokasjon (KOD, KV, EIK):

```
Varenummer-suffiks:   {VARE}-KOD, {VARE}-KV, {VARE}-EIK
```

**Ved flagging:**

| Handling | Beskrivelse |
|----------|-------------|
| Semi-finished opprettes | `products`: {VARE}-KOD, {VARE}-KV, {VARE}-EIK som Semi Finished |
| BOM på semi-finished | Semi-finished refererer TIL hovedproduktet: {Semi} → {VARE} (Qty Per = 1) |
| TRANSPORT-routing | Semi-finished får TRANSPORT-operasjon, run_time hentes fra `transport_ruter` |
| **Hovedproduktet** | **Forblir fullstendig urørt** — original BOM/routing endres aldri |

Eksempel (JD16073):

```
JD16073        → RM_50x75 (Qty 533,05, Co 0,5%, JD16073B)   ← urørt
JD16073-KOD    → JD16073  (Qty 1)                            ← materialkost rulles opp
JD16073-KOD    → TRANSPORT (run 45 min fra transport_ruter)

JD16073        → HOVLING @ SPESIALHOVEL                      ← urørt
JD16073-KOD    → TRANSPORT @ TRANSPORT
```

**Ved avflagging (`is_transport = 0`):**

- Alle semi-finished varianter og deres BOM/routing slettes
- Hovedproduktet er aldri rørt — ingen gjenoppretting nødvendig
- Endringsloggen fanger opp alle CREATE/DELETE

**Transportruter (`transport_ruter`-tabellen):**

| Kolonne | Beskrivelse |
|---------|-------------|
| from_loc | Fra-lokasjon (KOD, KV, EIK) |
| to_loc | Til-lokasjon |
| distance_km | Distanse i kilometer |
| run_time_minutes | Kjøretid en vei |
| setup_time_minutes | Laste-/lossetid |
| batch_size | Antall enheter per lass |

Eksporteres til Excel som ark "Transport Ruter". Én TRANSPORT-arbeidssenter brukes for alle ruter — run_time hentes fra rutetabellen per strekning.

**Kjøring:**

```bash
# Test at modulen fungerer (in-memory DB, påvirker ikke aktiv database)
python src/scripts/test_transport.py
```

**Testdekning (5/5 bestått):**

1. Sett flagg → semi-finished opprettes, hovedproduktets BOM/routing er urørt, TRANSPORT på semi-finished med run_time fra rutetabell
2. Fjern flagg → semi-finished slettes, hovedproduktet er intakt — ingen gjenoppretting nødvendig
3. Co-produkt (JD16073/JD16073B) forblir på hovedproduktet
4. Produksjonskjede (JD16098TF/Eksisterende semi-finished) dobles ikke
5. Alle tre varer flagges samtidig → ingen kollisjoner eller duplikater

---

## 23. Kommandolinje-verktøy

Alle kommandoer kjøres fra prosjektroten (`c:\Users\EspenGabrielsen\code\ProduksjonsKalkyle`).

```bash
# Full produktkalkyle (fra src/)
python src/kostberegning.py --excel src/Produksjonsmodell_Testdata_v3.xlsx

# Ett spesifikt produkt
python src/kostberegning.py --excel src/Produksjonsmodell_Testdata_v3.xlsx --product FG001

# Innebygget testdata (uten Excel)
python src/kostberegning.py --test

# Uten JSON-eksport
python src/kostberegning.py --excel src/Produksjonsmodell_Testdata_v3.xlsx --no-json

# JSON til spesifikk mappe
python src/kostberegning.py --excel src/Produksjonsmodell_Testdata_v3.xlsx --json-dir ./rapporter

# Generer testdata-Excel (fra src/scripts/)
python src/scripts/lag_testdata_v3.py

# SQLite-administrasjon (fra src/)
python src/data_repo.py --stats      # Vis tabell-statistikk
python src/data_repo.py --changes    # Vis endringslogg
python src/data_repo.py --clear      # Tøm data (bevar logg)

# Excel ↔ SQLite (fra src/)
python src/excel_bridge.py --import data.xlsx
python src/excel_bridge.py --export utdata.xlsx

# Baseline-verktøy (fra src/scripts/)
python src/scripts/lag_baseline.py              # Generer baseline
python src/scripts/lag_baseline.py --sammenlign # Sammenlign mot baseline
python src/scripts/sjekk_diff.py                # Sjekk diff

# Start Marimo-app
marimo run src/varekost_app.py

# Generer stilede PDF-er fra Markdown-dokumentasjon (fra src/)
python src/generer_dokumentasjon.py docs/fil.md                     # Én fil
python src/generer_dokumentasjon.py --all                           # Alle dokumentasjonsfiler
python src/generer_dokumentasjon.py --all --output ./rapporter      # Til egen mappe
```

### Alle flagg for kostberegning.py

| Flagg | Beskrivelse |
|-------|-------------|
| `--excel FIL` | Excel-fil med data (standard: `Produksjonsmodell_Mal.xlsx` i `src/`) |
| `--test` | Bruk innebygget testdata |
| `--product ITEMNO` | Beregn kost for ett produkt |
| `--no-json` | Ikke eksporter til JSON |
| `--json-dir MAPPE` | Mappe for JSON-eksport (standard: `.`) |

---

> **Dokumentasjon versjon 3.1**
> Sist oppdatert: juli 2026
> Basert på Kodal Hovleri som referanseeksempel
>
> System: Excel-datamodell → SQLite → Python → Marimo web-app → PDF/Excel