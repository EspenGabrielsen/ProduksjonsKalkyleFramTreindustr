Basert på gjennomgangen av regnearket `produksjonsmodell_eksport.xlsx` og prinsippene for industriell produksjonsstyring (ERP/MRP) for både sagbruk og høvleri, er her en detaljert gjennomgang av **hva som bør endres i datamodellen, hvorfor endringene er nødvendige, og hvordan den nye strukturen bør se ut.**

---

## 1. Hva som bør endres (Kjerneendringen)

Nåværende modell har en **asymmetrisk struktur** for sekundære utbytter:

1. **Bi-produkter (Flis, spon, bark):** Er modellert i en fleksibel relasjonstabell ($1:N$) under fanen `Biproduktregler`.
2. **Co-produkter (2. sortering, sidebord, sekundærdimensjoner):** Er «hardkodet» inn direkte i `Stykkliste`-tabellen med to felter (`Co-Prod %` og `Co-Prod Item No`).

### Endringen:

* **Fjern** feltene `Co-Prod %` og `Co-Prod Item No` fra `Stykkliste`-tabellen.
* **Slå sammen** logikken for Bi-produkter og Co-produkter til én felles tabell for alle sekundære utbytter, for eksempel kåret **`Sekundærprodukter`** (eller `Utbytteregler`).

---

## 2. Hvorfor endringen må gjøres

### A. Begrensning på $1:1$ for Co-produkter krasjer på sagbruket

I dagens `Stykkliste` kan hver rad kun peke på **ett enkelt Co-produkt**. På høvleriet har dette fungert for å fange opp én type B-vare (f.eks. `JD16073B`).

* **På sagbruket (Bø):** Et skurmønster (postering) på saga spytter ut **mange produkter samtidig** fra samme stokk:
* Hoveddimensjon (Senterutbytte, 1. sort)
* Planlagte sidebord i en annen dimensjon (Ønsket Co-produkt)
* 2. sortering / Sekunda-trelast (Uønsket Co-produkt)


* 3. sortering / Vrak / Emballasjeved (Uønsket Co-produkt)




* Med dagens struktur har du ikke mulighet til å registrere mer enn én av disse sidebords-/sorterings-variantene per BOM-linje.

### B. Begrepsforvirring og ulike økonomiske logikker

I industriell økonomi behandles sekundære utbytter ulikt i kalkylen:

1. **Planlagte Co-produkter (f.eks. Sidebord):** Er ønskede produkter som produseres parallelt. Tømmerkost og linjetid skal **fordeles (allokeres)** mellom kjerne og sidebord.
2. **Feil / Sekunda Co-produkter (2. og 3. sort):** Er uønsket utbytte som senker farten og effektiviteten for A-varen. De tildeles en lavere verdi/kost, slik at A-varen bærer belastningen.
3. **Bi-produkter (Flis, spon, bark):** Er uunngåelig avfall/råstoff der salgsinntekten (markedsverdien) **trekkes fra** produksjonskostnaden til hovedvaren.

Ved å skille disse med en eksplisitt regel i steden for å blande felter i stykklisten, skaper du en entydig matematisk logikk i beregningsmotoren din.

---

## 3. Hvordan den nye datamodellen vil se ut

### Tabell 1: `Stykkliste` (Forenklet og oppryddet)

`Stykkliste` skal kun håndtere **Input** (hva som går inn i prosessen):

| Feltnavn | Beskrivelse | Eksempel (Sagbruk) | Eksempel (Høvleri) |
| --- | --- | --- | --- |
| `Parent Item No` | Sluttprodukt / Postering | `SKUR_GRAN_18-20` | `JD16073` |
| `Component Item No` | Råvare inn | `TØMMER_GRAN_18-20` | `RM_50x75_US_V_Gran` |
| `Quantity Per` | Mengde råvare inn | `1.85` | `533.05` |
| `Unit of Measure` | Enhet råvare | `M3` | `LM/M3` |
| `Scrap %` | Svinnprosent | `0` | `0` |

*(Feltene `Co-Prod %` og `Co-Prod Item No` fjernes herfra).*

---

### Tabell 2: `Sekundærprodukter` (Ny felles utbyttetabell)

Denne erstatte `Biproduktregler` og dekker både bi- og co-produkter ($1:N$-relasjon):

| Feltnavn | Datatype | Beskrivelse |
| --- | --- | --- |
| `Parent Item No` | Text | Kobling til Hovedvare / Skurserie. |
| `Output Item No` | Text | Varenummer på sekundærvaren. |
| `Output Type` | Dropdown | `Co-Product` (Trelast/B-vare) eller `By-Product` (Flis/spon). |
| `Expected Quantity` | Float | Utbyttemengde per enhet. |
| `Unit of Measure` | Text | Enhet (`M3`, `KG`, `LM`, `%`). |
| `Allocation Method` | Dropdown | Hvordan verdien/kosten håndteres:<br>

<br>• `Reduce Main Cost` (Standard for Bi-produkt)<br>

<br>• `Volume/Value Split` (For planlagte sidebord)<br>

<br>• `Fixed Discount / NRV` (For 2./3. sortering) |
| `Market Value` | Float | Eventuell fast overstyrt markedspris (NOK). |

---

## 4. Eksempel på data i den nye tabellen (Sagbruk vs. Høvleri)

Slik vil rader i den nye `Sekundærprodukter`-tabellen se ut i praksis:

```
─────────────────────────────────────────────────────────────────────────────────────────────────────────────
Parent Item    Output Item    Output Type   Expected Qty   UOM    Allocation Method    Market Value (NOK)
─────────────────────────────────────────────────────────────────────────────────────────────────────────────
// EKSEMPEL: SAGBRUK I BØ (Saging av Gran 18-20)
SKUR_G18-20    45x145_G_1S    Co-Product    0.45           M3     Volume/Value Split   -
SKUR_G18-20    19x100_G_1S    Co-Product    0.12           M3     Volume/Value Split   -
SKUR_G18-20    45x145_G_2S    Co-Product    0.04           M3     Fixed Discount / NRV -
SKUR_G18-20    FLIS_CELL      By-Product    0.28           M3     Reduce Main Cost     280.00
SKUR_G18-20    SPON_BARK      By-Product    0.11           M3     Reduce Main Cost     90.00

// EKSEMPEL: HØVLERI (Eksisterende logikk fra arket ditt)
JD16073        JD16073B       Co-Product    0.005          LM/M3  Fixed Discount / NRV -
JD16073        SA66100        By-Product    0.150          KG     Reduce Main Cost     1.50
─────────────────────────────────────────────────────────────────────────────────────────────────────────────

```

---

## Oppsummering av fordelene

1. **Skalerbarhet ($1:N$):** Sagbruket i Bø kan ha 2, 5 eller 10 utbytte-produkter per serie uten at datamodellen krasjer.
2. **Fleksibilitet:** Høvleriet kan fortsette nøyaktig som før, men får i tillegg muligheten til å registrere flere sorteringsgrader (f.eks. 3. sort vrak).
3. **Renere databasestruktur:** `Stykkliste` handler kun om *Input*, mens `Sekundærprodukter` handler om *Output*.
4. **Enkelt for brukeren:** Brukeren har alt utbytte på én plass og skiller dem enkelt med `Output Type` i stedet for å forholde seg til to separate systemer.

---

# FASE 1: Ny `Sekundærprodukter`-tabell — Detaljert implementasjonsplan

## Mål

Legge til en ny `sekundaerprodukter`-tabell **ved siden av** eksisterende `byproduct_rules` (ikke erstatte den). Eksisterende logikk forblir 100% urørt — dette er en ren additiv endring som forbereder grunnen for sagbrukets 1:N-utbyttemodell.

---

## 1. Database (`data_repo.py`)

### 1.1 Ny tabell

```sql
CREATE TABLE IF NOT EXISTS sekundaerprodukter (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    parent_item_no TEXT NOT NULL,
    output_item_no TEXT NOT NULL,
    output_type TEXT NOT NULL DEFAULT 'Co-Product',  -- 'Co-Product' | 'By-Product'
    expected_quantity REAL NOT NULL DEFAULT 0,
    uom TEXT NOT NULL DEFAULT '',
    allocation_method TEXT NOT NULL DEFAULT 'Reduce Main Cost',
    market_value REAL NOT NULL DEFAULT 0,
    UNIQUE(parent_item_no, output_item_no)
);
```

### 1.2 Indekser

```sql
CREATE INDEX IF NOT EXISTS idx_sek_parent ON sekundaerprodukter(parent_item_no);
CREATE INDEX IF NOT EXISTS idx_sek_output ON sekundaerprodukter(output_item_no);
```

### 1.3 CRUD-funksjoner i `DataRepo`

- `upsert_sekundaerprodukter(rows: list[dict], source: str)` — samme mønster som `upsert_byproduct_rules`
- `get_sekundaerprodukter_for(parent_item_no: str) -> list[dict]`

### 1.4 Change-log

Utvid `log_batch_changes()` til å også logge `sekundaerprodukter`-feltendringer (felt: `expected_quantity`, `market_value`, `allocation_method`, `output_type`).

---

## 2. Domenemodell (`kostberegning.py`)

### 2.1 Ny dataclass

```python
@dataclass
class Sekundaerprodukt:
    parent_item_no: str
    output_item_no: str
    output_type: str          # "Co-Product" | "By-Product"
    expected_quantity: float
    uom: str
    allocation_method: str    # "Reduce Main Cost" | "Volume/Value Split" | "Fixed Discount"
    market_value: float
```

### 2.2 `ExcelData` / `SqliteData`

- Nytt felt: `sekundaerprodukter: list[Sekundaerprodukt]`
- Ny indeks: `_sekundaerprodukt_index: dict[str, list[Sekundaerprodukt]]`
- Ny parser: `_parse_sekundaerprodukter()` (leser fra nytt Excel-ark hvis det finnes)
- Nytt oppslag: `sekundaerprodukter_for(parent_item_no: str) -> list[Sekundaerprodukt]`
- `SqliteData._load_sekundaerprodukter()` — leser fra den nye SQLite-tabellen

**VIKTIG:** Eksisterende `byproduct_rules` og `_calc_byproduct_value()` endres ikke. Den nye tabellen er et ekstra oppslag som `CostCalculator` kan bruke i fremtiden.

---

## 3. Excel-import/eksport (`excel_bridge.py`)

### 3.1 Import

- Hvis Excel-filen har et ark navngitt `Sekundaerprodukter` (eller `Secondary Output`), leses det inn og lagres via `upsert_sekundaerprodukter()`
- Hvis arket ikke finnes, hoppes det over uten feil (additiv)
- Norsk/engelske ark-navn håndteres via `sheet_i18n.json`

### 3.2 Eksport

- Legg til `Sekundaerprodukter` som nytt ark i `export_sqlite_to_excel()` (etter `By Product Rules`)
- Kolonner: `Parent Item No`, `Output Item No`, `Output Type`, `Expected Quantity`, `Unit of Measure`, `Allocation Method`, `Market Value`

### 3.3 `sheet_i18n.json`

- Legg til ark-oppføring for `Sekundaerprodukter` (norsk navn + tabColor)

---

## 4. Kostnadsberegning (`kostberegning.py`)

**Ingen endringer i Fase 1.** Eksisterende co-produkt- og biprodukt-logikk (`_calc_co_product_results()`, `_calc_byproduct_value()`) forblir intakt.

Fremtidig fase 2 vil utvide `CostCalculator` til å konsultere `sekundaerprodukter_for()` for nye allokeringsmetoder.

---

## 5. Excel-mal (`oppdater_mal.py`, `lag_testdata_v3.py`)

- Legg til valgfritt nytt ark `Sekundaerprodukter` i malen
- Eksisterende ark uendret (bakoverkompatibelt)

---

## 6. Verifisering (Fase 1)

1. `python src/data_repo.py --stats` — `sekundaerprodukter`-tabellen skal vises
2. Import av eksisterende Excel-fil (uten nytt ark) skal ikke feile
3. Import av Excel-fil **med** `Sekundaerprodukter`-ark skal lagre rader korrekt
4. `export_sqlite_to_excel()` skal generere arket med korrekt struktur
5. Baseline-beregning endres ikke (alle kalkyler med eksisterende data gir identiske tall)
