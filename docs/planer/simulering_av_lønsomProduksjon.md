# Produksjons- og Batchoptimering med Flernivå-BOM — Prosjektspesifikk Plan

> **Prosjekt:** ProduksjonsKalkyle (Fram Treindustri / Kodal Hovleri)  
> **Dato:** August 2026  
> **Status:** Planleggingsfase — implementasjon ikke påbegynt

---

## 1. Kontekst og mål

Dagens system (`kostberegning.py` → `CostCalculator` + `SimulationEngine`) beregner enhetskostnader per produkt/lokasjon og støtter "what-if"-simuleringer (overstyr råvarepriser, kapasitet, etc.). Det som **mangler**, er en motor som automatisk finner den **mest optimale produksjonsplanen** gitt en etterspørsel, på tvers av flere produksjonssteg og lokasjoner.

**Mål:** Bygge `src/optimization_engine.py` — en frittstående MILP-optimeringsmotor (PuLP) som:
- Leser etterspørsel (hvilke sluttprodukter skal leveres, i hvilke perioder, til hvilke kunder)
- Forstår flernivå-avhengigheter (Skurlast → Ubehandlet → Grunnet → Malt)
- Finner optimal fordeling av produksjonsbatcher på tvers av lokasjoner og perioder
- Minimerer totalkostnad (produksjon + transport + lagerhold)

**Viktig prinsipp:** Eksisterende kode (`kostberegning.py`, `varekost_app.py`, `excel_bridge.py`) skal **ikke endres**. Optimeringsmotoren importerer og gjenbruker det som finnes (`SqliteData`, `CostCalculator`).

---

## 2. Produksjonskjede (eksempel fra datamodellen)

```
[Skurlast / Råvare]          Item Type: Raw Material
       │
       ▼ (Høvling — Work Center: HOVLERI_1 eller HOVLERI_2)
[Ubehandlet Kledning]        Item Type: Semi Finished
       │
       ▼ (Grunning — Work Center: OVERFLATEBEHANDLING_1)
[Grunnet Kledning]           Item Type: Semi Finished
       │
       ▼ (Malingslinje — Work Center: OVERFLATEBEHANDLING_1)
[Malt Kledning]              Item Type: Finished Good
       │
       └──► Etterspørsel fra kunde (demand-tabellen)
```

**Prinsipper for optimeringen:**
- **Etterspørsel trigges fra toppen:** Kun sluttproduktet (Malt Kledning) har direkte etterspørsel. Delprodukter produseres kun fordi de trengs videre i kjeden.
- **Materialbalanse:** For å levere X meter Malt Kledning må modellen produsere (eller ha på lager) X meter Grunnet Kledning, osv.
- **Lokasjonsvalg:** Ubehandlet kledning kan høvles på Høvleri 1, transporteres til Høvleri 2 for grunning — hvis det gir lavest totalkostnad. Transportkostnader hentes fra eksisterende `transport_ruter`-tabell.

---

## 3. Eksisterende komponenter som gjenbrukes

| Komponent | Modul | Hvordan det brukes |
|-----------|-------|-------------------|
| **Stamdata** (produkter, BOM, routing, work centers, lokasjoner, kostpriser) | `kostberegning.SqliteData` | Leses direkte. `optimization_engine.py` importerer `SqliteData` og får tilgang til alle produkter, lokasjoner, arbeidssentre, kapasitetsdata, transportruter |
| **Enhetskostnader per produkt/lokasjon** | `kostberegning.CostCalculator` | `CostCalculator.calculate_all()` kjører *én gang* før MILP-en. Resultatet (materialkost, operasjonskost, setupkost per enhet) mates inn som kostnadskoeffisienter i MILP-objektfunksjonen |
| **Topologisk BOM-sortering** | `CostCalculator` → `get_dependency_order()` | Gjenbrukes for å automatisk populerer `bom_structure`-tabellen fra eksisterende `bom_lines` |
| **Transportkostnader** | `transport_ruter`-tabellen i SQLite | `cost_per_m3` per rute brukes i MILP-ens transportvariabler |
| **Kapasitetsdata** | `work_centers` + `capacity_days` | `capacity_hours_day`, `effective_capacity_pct` og `capacity_days` brukes som kapasitetsbegrensninger per arbeidssenter per periode |

---

## 4. Nye SQLite-tabeller (legges til i `data_repo.py` → `SCHEMA_SQL`)

Kun `SCHEMA_SQL` utvides. Ingen CRUD-funksjoner nødvendig — data settes inn manuelt eller via enkle INSERT-scripts.

### A. `demand` — Sluttetterspørsel

```sql
CREATE TABLE IF NOT EXISTS demand (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id TEXT NOT NULL,          -- Eks: 'KLEDNING_MALT_19X148' (må finnes i products.item_no)
    period INTEGER NOT NULL,           -- Eks: 34 (uke 34), 2026-08 (måned), eller 1, 2, 3...
    quantity REAL NOT NULL,            -- Eks: 5000 (i produktets base_uom, f.eks. LM eller M3)
    location_code TEXT NOT NULL,       -- Eks: 'KV' — hvilken lokasjon skal varen leveres til?
    customer_region TEXT,              -- Valgfritt: for fremtidig transportberegning
    UNIQUE(product_id, period, location_code)
);
```

### B. `bom_structure` — Flernivå-avhengigheter for MILP

```sql
CREATE TABLE IF NOT EXISTS bom_structure (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    parent_product_id TEXT NOT NULL,   -- Eks: 'KLEDNING_GRUNNET_19X148' (den som FORBRUKER)
    child_product_id TEXT NOT NULL,    -- Eks: 'KLEDNING_UBEHANDLET_19X148' (den som KREVES)
    yield_factor REAL NOT NULL DEFAULT 1.0,  -- Forbruk per enhet (1.05 = 5% svinn mellom steg)
    UNIQUE(parent_product_id, child_product_id)
);
```

**Hvordan `bom_structure` populeres:**  
Kan populeres automatisk ved å traversere eksisterende `bom_lines` med topologisk sortering (samme logikk som `CostCalculator.get_dependency_order()`). For hvert parent→child-par i BOM-kjeden legges det inn en rad med `yield_factor = 1 + scrap_pct/100`.

---

## 5. `src/optimization_engine.py` — Arkitektur

Ny frittstående fil. Importerer fra eksisterende moduler uten å endre dem.

```
optimization_engine.py
├── Importerer SqliteData         (fra kostberegning)
├── Importerer CostCalculator     (fra kostberegning)
├── Leser demand + bom_structure  (fra SQLite)
├── Setter opp MILP               (PuLP)
│   ├── Beslutningsvariabler
│   ├── Materialbalanse-betingelser
│   ├── Kapasitetsbetingelser
│   ├── Batch/Rigg-betingelser (Big-M)
│   └── Objektfunksjon
└── Løser og returnerer optimal plan
```

### 5.1 Beslutningsvariabler

| Variabel | Indekser | Type | Beskrivelse |
|----------|----------|------|-------------|
| `X[p, l, t]` | produkt, lokasjon, periode | Kontinuerlig ≥ 0 | Produsert mengde av produkt p på lokasjon l i periode t |
| `Y[p, l, t]` | produkt, lokasjon, periode | Binær | 1 hvis batch startes for produkt p på lokasjon l i periode t |
| `I[p, l, t]` | produkt, lokasjon, periode | Kontinuerlig ≥ 0 | Utgående lagerbeholdning av produkt p på lokasjon l etter periode t |
| `T[p, from_loc, to_loc, t]` | produkt, fra-lok, til-lok, periode | Kontinuerlig ≥ 0 | Transportmengde av produkt p fra lokasjon A til B i periode t |

### 5.2 Betingelser (constraints)

**Materialbalanse (per produkt, lokasjon, periode):**
```
I[p,l,t-1] + X[p,l,t] + SUM(innkommende transport)
    = Demand[p,l,t]
    + SUM( X[parent,l,t] × yield_factor[parent→p] )   ← forbruk i videre produksjon
    + SUM(utgående transport)
    + I[p,l,t]
```

**Kapasitet (per arbeidssenter, lokasjon, periode):**
```
SUM over produkter med routing på dette WC:
    (run_time_min/60 × X[p,l,t] + setup_time_min/60 × Y[p,l,t])
    ≤ available_hours[wc, t] × effective_capacity_pct/100
```

**Big-M batch-kobling (per produkt, lokasjon, periode):**
```
X[p,l,t] ≤ M × Y[p,l,t]    (der M er en stor nok verdi, f.eks. maks batch-størrelse)
```

### 5.3 Objektfunksjon (minimer)

```
Minimer SUM:
  + X[p,l,t] × unit_production_cost[p,l]     (material + operasjon + setup per enhet)
  + Y[p,l,t] × changeover_cost[p,l]          (ekstra riggkost per batch-start)
  + T[p,from,to,t] × unit_transport_cost     (fra transport_ruter.cost_per_m3)
  + I[p,l,t] × unit_holding_cost             (lagerholdskost per enhet per periode)
```

`unit_production_cost[p,l]` hentes fra `CostCalculator.calculate_all()` — dette inkluderer allerede materialkost, operasjonskost, setupkost, biproduktverdi, og effektiv kapasitetsjustering.

---

## 6. Kjøring og grensesnitt

```bash
# Kjør optimering (leser demand fra SQLite)
python src/optimization_engine.py

# Spesifiser database-path
python src/optimization_engine.py --db path/to/produksjonskalkyle.db

# Eksporter resultat til JSON
python src/optimization_engine.py --output resultat.json
```

Output: JSON med optimal produksjonsplan:
```json
{
  "status": "Optimal",
  "total_cost": 123456.78,
  "production_plan": [
    {
      "product": "KLEDNING_UBEHANDLET_19X148",
      "location": "KV",
      "period": 34,
      "quantity": 5000,
      "cost_per_unit": 42.50,
      "total_cost": 212500.00
    }
  ],
  "transport_plan": [...],
  "inventory_levels": [...]
}
```

---

## 7. TODO: Data som må samles inn

Før optimeringsmotoren kan testes med reelle data, må følgende være på plass:

### 7.1 Etterspørselsdata (`demand`-tabellen)

- [ ] Identifiser **hvilke sluttprodukter** som har direkte kunde-etterspørsel (Item Type = Finished Good)
- [ ] For hvert sluttprodukt: angi **uke/periode** (f.eks. uke 34, 35, 36...)
- [ ] For hvert sluttprodukt: angi **kvantum** (i produktets base_uom, f.eks. LM eller M3)
- [ ] Angi **leveringslokasjon** (hvilken fabrikk/lager skal varen leveres til?)
- [ ] Datakilde: Kan komme fra ordresystem, prognoser, eller `production_scenarios`-tabellen (utvides med periode og lokasjon)

### 7.2 Flernivå-BOM (`bom_structure`-tabellen)

- [ ] Kartlegg **produksjonskjeden**: hvilke Finished Goods → Semi Finished → Raw Material finnes i dagens `bom_lines`?
- [ ] Bestem **yield_factor** mellom hvert steg (1.0 = 1:1, 1.05 = 5% svinn)
- [ ] Alternativ: `bom_structure` kan auto-populeres fra `bom_lines` (gjøres i kode)

### 7.3 Lagerholdskostnad

- [ ] Hva koster det å holde 1 enhet på lager i 1 periode? (f.eks. rente, svinn, plasskostnad)
- [ ] Forenkling: Kan settes til 0 i første iterasjon

### 7.4 Batch/changeover-kostnad

- [ ] Hva koster en omstilling/rigg på hvert arbeidssenter? (Dagens `routing_lines` har `setup_time_minutes` og `changeover_time_minutes`)
- [ ] Forenkling: Kan settes til 0 i første iterasjon — riggkost er allerede del av enhetskostnaden

### 7.5 Transportruter

- [ ] Er `transport_ruter`-tabellen fylt med ruter mellom alle relevante lokasjoner? (from_loc → to_loc, cost_per_m3)
- [ ] Hvis ikke: kartlegg alle relevante ruter mellom fabrikker og fyll inn

### 7.6 Kapasitetsdata

- [ ] Er `work_centers` oppdatert med korrekt `capacity_hours_day` og `effective_capacity_pct`?
- [ ] Er `capacity_days` fylt for de aktuelle periodene? (Hvis ikke: perioder uten data antar standard kapasitet fra work_centers)

---

## 8. Sekvensavhengig omstilling (changeover)

### 8.1 Problem

Noen profilbytter på høvelen krever mer tid enn andre. Å bytte fra 19mm tykkelse til 22mm (ribbing-endring) tar mye lengre tid enn å bytte farge på samme dimensjon.

### 8.2 Løsning: Familie-basert (Gemini-tilnærming 3)

**Ikke** full sekvensiering ($Z_{i,j}$-variabler — ville kvele CBC-solveren). I stedet:

- **Produktfamilie** = FTI prefiks + siffer (f.eks. `JD19073` — alle varianter GH/GF/TF/EH tilhører samme familie)
- Ny binærvariabel `YF[f, wc, t]` = 1 hvis minst ett produkt i familie f produseres på arbeidssenter wc i periode t
- **Kobling:** `Y[p,l,t] ≤ YF[familie(p), wc, t]`
- **Kapasitets-straff:** Hver aktiv familie trekker `0.5 × snittstraff` fra tilgjengelig kapasitet (i timer). Dette straffer blanding av mange familier i samme periode uten å spore eksakt rekkefølge.

### 8.3 Changeover-estimater (fallback-generator)

Basert på FTI-nummerstrukturen (se CLINE.md):

| Nivå | Endring | Eksempel | Tidsestimat |
|------|---------|----------|-------------|
| 1 | Suffiks-bytte (samme familie) | GH→GF→VF | 0 min |
| 2 | Bredde-bytte (samme tykkelse) | JD19**073**→JD19**148** | 15-30 min (snitt 22.5) |
| 3 | Tykkelse-bytte | JD**19**148→JD**22**148 | 45-60 min (snitt 52.5) |
| 4 | Prefiks-bytte (annen serie) | JD→JV | ~90 min |

### 8.4 `changeover_matrix`-tabell (data som kan fylles manuelt)

```sql
CREATE TABLE IF NOT EXISTS changeover_matrix (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_center_code TEXT NOT NULL,
    from_family TEXT NOT NULL,
    to_family TEXT NOT NULL,
    changeover_minutes REAL NOT NULL DEFAULT 0,
    changeover_cost REAL NOT NULL DEFAULT 0,
    UNIQUE(work_center_code, from_family, to_family)
);
```

- Når tabellen er fylt, brukes eksakte verdier fra matrisen.
- Når tabellen er tom, brukes fallback-generatoren basert på FTI-strukturen.

### 8.5 Fremtidig: To-trinns dekomponering (utilsiktet)

Når quick-win fungerer, kan en separat TSP-sekvenseringsalgoritme (OR-Tools) kjøre etter MILP-en for å finne optimal rekkefølge innenfor hver periode. Dette er dokumentert som langsiktig veikart i Gemini-analysen.

---

## 9. Implementasjonsrekkefølge

| Steg | Hva | Fil | Avhengigheter |
|------|-----|-----|--------------|
| **1** | Legg til `demand` + `bom_structure` i `SCHEMA_SQL` | `data_repo.py` | Ingen |
| **2** | Bygg `optimization_engine.py` med PuLP | `src/optimization_engine.py` | `SqliteData`, `CostCalculator`, PuLP (`requirements.txt`) |
| **3** | Test med syntetiske data (hardkodet demand + bom_structure) | `optimization_engine.py` | Steg 1+2 |
| **4** | Auto-populer `bom_structure` fra `bom_lines` | `optimization_engine.py` | Steg 2 |
| **5** | Test med reelle data fra Excel/SQLite | `optimization_engine.py` | Steg 3+4 + innsamlede data |

---

## 10. Prognose fra historisk salg (batch-størrelsesavveining)

### 10.1 Historisk salg som prognose

`historical_sales`-tabellen inneholder 12 måneder med faktisk salg per produkt×uke.
Optimeringsmotoren beregner gjennomsnittlig ukentlig salg per (produkt, lokasjon)
og projiserer dette fremover i de neste `--forecast-weeks` periodene (default: 12).
Dette gir modellen et bilde av fremtidig etterspørsel, slik at den kan velge riktig
batch-størrelse: en stor batch sparer setup-kost, men påløper lagerholdskostnaden
(`--holding-cost-pct`).

### 10.2 Setup-kost-terskel: `--forecast-max-setup-pct` (default: 30)

**Bakgrunn:** Uten en terskel genererer prognosen meningsløse mikrobatcher for
nisjeprodukter. For eksempel vil et produkt med 5 LM årlig salg få en prognoselinje
på `5/52 ≈ 0.1 LM` hver uke — noe som aldri kan forsvares økonomisk.

**Løsning (Alternativ A — kun operasjonskost, eksklusiv råvare):**

```
min_mengde = fixed_setup_kost / (enhetskost × max_setup_pct / 100)

Hvis snitt_uke_salg < min_mengde:
    slå sammen N = ceil(min_mengde / snitt_uke_salg) uker
```

| Verdi | Betydning | Effekt |
|-------|-----------|--------|
| **0%** | Ingen grense | Alle prognoselinjer godtas (inkl. 0.1 LM) |
| **10%** | Setup ≤ 10% av operasjonskost | Strengt — kun store batcher |
| **30%** (default) | Setup ≤ 30% av operasjonskost | Moderat — god balanse |
| **60%** | Setup ≤ 60% av operasjonskost | Liberalt — de fleste produkter slipper gjennom |

**Viktig:** `enhetskost` her er **kun operasjonskost/routing-kost per enhet**
(Alternativ A), eksklusiv råvare. Dette er dokumentert som et bevisst valg — det er
routing-kostnaden som driver setup-beslutningen, ikke råvareprisen.

**JD19123-eksempel (30% terskel):**
- Setup-kost: ~2 500 kr, enhetskost (operasjon): ~500 kr/LM
- `min_mengde = 2500 / (500 × 0.30) ≈ 17 LM`
- Snitt ukentlig salg: ~266 LM → 266 > 17 → **får egen ukes-prognose** ✓

**Nisjeprodukt-eksempel (0.1 LM/uke):**
- `min_mengde` ≈ 17 LM, snitt 0.1 LM
- `N = ceil(17 / 0.1) = 170 uker` → **får aldri en prognoselinje** ✓

### 10.3 Mikro-behov-filter etter BOM-ekspansjon

Etter BOM-traversering filtreres alle behov < 1 enhet bort fra modellen
(`MIN_BEHOV = 1.0`). Dette fjerner mikro-støy fra åpne ordrer med ubetydelige
volumer og fra avledet behov i BOM-kjeden, uten å påvirke reelle ordrer.

### 10.4 Bruk

```bash
# Kjøring med prognose og terskel
python src/optimization_engine.py --db src/produksjonskalkyle.db --days 500 \
    --forecast-weeks 12 --forecast-max-setup-pct 30 --output resultat.json
```

---

## 11. Avhengighet: PuLP

`puLP` må legges til i `requirements.txt` hvis det ikke allerede er der:

```
PuLP>=2.7.0
```

Sjekk med: `pip list | grep -i pulp`
