# Batch-optimalisering med SimPy — Implementasjonsplan

> **Mål:** Utvide ProduksjonsKalkyle med et SimPy-basert simuleringsverktøy for batch-optimalisering og "hva-hvis"-scenarioer for produksjonsflyt.  
> **Hovedprinsipp:** `kostberegning.py` og `varekost_app.py` forblir **helt urørt**.  
> All ny funksjonalitet bygges som separate moduler og en egen Marimo-app.
>
> **Status:** Plan godkjent — klar for implementasjon.

---

## Innholdsfortegnelse

1. [Overordnet arkitektur](#1-overordnet-arkitektur)
2. [SimPy-modellering — Detaljert design](#2-simpy-modellering--detaljert-design)
   - 2A. ProductionEnvironment
   - 2B. ProductionOrder
   - 2C. OrderGenerator
   - 2D. KPI-er som samles inn
   - 2E. Kostfunksjon for optimalisering
3. [Hvordan BOM og Routing brukes i simuleringen](#3-hvordan-bom-og-routing-brukes-i-simuleringen)
4. [Nye filer](#4-nye-filer)
   - 4A. `batch_simulator.py` — Kjernemodul
   - 4B. `demand_profile.py` — Etterspørselsdata
   - 4C. `batch_app.py` — Marimo-grensesnitt
5. [Konkret eksempel: Høvleri-scenarioet](#5-konkret-eksempel)
6. [Forretningsnytte](#6-forretningsnytte)
7. [Forutsetninger for å lykkes](#7-forutsetninger-for-a-lykkes)
8. [Data vi må ha stålkontroll pa](#8-data-vi-ma-ha-stalkontroll-pa)
9. [Anbefalt tilnærming — Tre faser](#9-anbefalt-tilnarming--tre-faser)
10. [Trinnvis implementasjonsplan](#10-trinnvis-implementasjonsplan)
11. [Tekniske garantier](#11-tekniske-garantier)

---

## 1. Overordnet arkitektur

```
┌─────────────────────────────────────────────────────────┐
│                    PRODUKSJONSKALKYLE                     │
│                                                          │
│  kostberegning.py (URØRT)                                │
│  ├── ExcelData / SqliteData: leser BOM, Routing, WC     │
│  ├── CostCalculator: beregner enhetskost                 │
│  └── SimulationEngine: what-if per enhet                 │
│                                                          │
│  batch_simulator.py (NY MODUL)                           │
│  ├── ProductionEnvironment: bygger SimPy-miljø           │
│  │   ├── WorkCenterResource: wrapper rundt simpy.Resource│
│  │   ├── OrderGenerator: genererer etterspørselsstrøm    │
│  │   └── ProductionOrder: simpy.Process for én ordre     │
│  ├── BatchOptimizer: kjører simulering for ulike batcher │
│  │   ├── simulate_batch_size(batch): kjør én simulering  │
│  │   ├── evaluate_batch_sizes(batches): test flere       │
│  │   └── find_optimal(): finn laveste totalkost          │
│  └── BatchResult: resultatstruktur (KPI-er)              │
│                                                          │
│  demand_profile.py (NY MODUL)                            │
│  ├── DemandProfile: ordreliste med tidspunkt og kvantum  │
│  ├── DemandGenerator.from_annual_volume()                │
│  └── DemandGenerator.from_historical()                   │
│                                                          │
│  batch_app.py (NY MARIMO-APP, port 8083)                 │
│  ├── Velg produkt / produktgruppe                        │
│  ├── Juster etterspørselsprofil                          │
│  ├── Velg batch-størrelser som skal testes               │
│  ├── Visualisering: kostnadskurver, kapasitetsutnyttelse │
│  └── Eksport: anbefalt batch-størrelse per produkt       │
└─────────────────────────────────────────────────────────┘
```

**Viktige arkitekturregler:**
1. **`kostberegning.py` endres aldri** — Importerer kun `CostCalculator` og `SqliteData` som lese-avhengigheter.
2. **`varekost_app.py` endres aldri** — Ny app kjører på egen port (8083).
3. **Samme SQLite-database** (`endringslogg.db`) — Ingen nye tabeller kreves i første fase.
4. **Hver ny modul har én ansvarlighet.**

---

## 2. SimPy-modellering — Detaljert design

### 2A. ProductionEnvironment (SimPy-miljøet)

```python
class ProductionEnvironment:
    """Holder hele SimPy-simuleringen."""

    def __init__(self, data: SqliteData, product_no: str,
                 batch_size: float, demand_profile: DemandProfile):
        self.env = simpy.Environment()
        self.data = data
        self.batch_size = batch_size
        self.demand_profile = demand_profile

        # Bygg work center resources — kapasitet = 1 (én maskinlinje per WC)
        self.work_centers: dict[str, simpy.PriorityResource] = {}
        for wc in data.work_centers:
            if wc.active:
                self.work_centers[wc.code] = simpy.PriorityResource(
                    self.env, capacity=1
                )

        # Work Center metadata (kostnad, effektiv kapasitet)
        self.wc_meta: dict[str, WorkCenter] = {
            wc.code: wc for wc in data.work_centers if wc.active
        }

        # Mellomlager (WIP) — ett per operasjon per produkt
        self.wip_buffers: dict[str, simpy.Store] = {}

        # KPI-samling
        self.stats = SimulationStats()
```

### 2B. ProductionOrder (simpy.Process)

Hver produksjonsordre går gjennom alle operasjoner i sin routing:

```python
def production_order(env, order_id, product_no, batch_qty,
                     routing_lines, work_centers, wc_meta, stats):
    """
    SimPy-prosess for én produksjonsordre.

    1. Ankommer systemet
    2. For hver operasjon i routing:
       a. Be om work center (med prioritet)
       b. Setup-tid (omstilling)
       c. Kjøretid = run_time_minutes * batch_qty
       d. Frigi work center
    3. Registrer ferdigstillelse
    """
    arrival_time = env.now

    # Prioritet basert på hastegrad (lavere tall = høyere prioritet)
    # Standardprioritet = 1, hasteordre = 0
    priority = 1

    for rl in routing_lines:
        wc = work_centers[rl.work_center_code]
        wc_info = wc_meta[rl.work_center_code]

        # 1. Be om tilgang til work center (med prioritet)
        with wc.request(priority=priority) as req:
            yield req

            # Ventetid i kø
            queue_time = env.now - arrival_time
            stats.record_queue_time(rl.work_center_code, queue_time, order_id)

            # 2. Setup (omstillingstid)
            setup_hours = rl.setup_time_minutes / 60.0
            yield env.timeout(setup_hours)
            stats.record_setup(rl.work_center_code, rl.setup_time_minutes, product_no)

            # 3. Kjøretid (run_time * batch_size)
            run_hours = (rl.run_time_minutes / 60.0) * batch_qty
            yield env.timeout(run_hours)
            stats.record_run(rl.work_center_code, run_hours, batch_qty)

    # 4. Registrer ferdigstillelse
    completion_time = env.now
    stats.record_completion(
        order_id=order_id,
        product_no=product_no,
        arrival_time=arrival_time,
        completion_time=completion_time,
        batch_qty=batch_qty,
        lead_time=completion_time - arrival_time
    )
```

### 2C. OrderGenerator (etterspørsel)

Genererer ordrer i henhold til en etterspørselsprofil:

```python
def order_generator(env, product_no, batch_size, demand_profile,
                    routing_lines, work_centers, wc_meta, stats):
    """Generer produksjonsordrer basert på etterspørselsprofil."""

    def generate_batches(quantity):
        """Del opp en ordre i batcher av batch_size."""
        num_batches = math.ceil(quantity / batch_size)
        for b in range(num_batches):
            yield min(batch_size, quantity - b * batch_size)

    for i, (arrival_time, quantity) in enumerate(demand_profile.orders):
        # Vent til ordren ankommer
        if arrival_time > env.now:
            yield env.timeout(arrival_time - env.now)

        # Del opp i batcher og start produksjonsprosesser
        for j, batch_qty in enumerate(generate_batches(quantity)):
            order_id = f"{product_no}_{i}_{j}"
            env.process(production_order(
                env, order_id, product_no, batch_qty,
                routing_lines, work_centers, wc_meta, stats
            ))
```

### 2D. KPI-er som samles inn

| KPI | Formel | Hvorfor relevant |
|---|---|---|
| **Gjennomsnittlig ledetid** | completion_time − arrival_time | Kundetilfredshet, leveringspresisjon |
| **Total setup-kost** | Σ (setup_timer × timekost × antall_changeovers) | Direkte omstillingskostnad |
| **Total operasjonskost** | Σ (run_timer × timekost) × batch_qty | Direkte produksjonskostnad |
| **Kapasitetsutnyttelse per WC** | (bearbeidet_tid / tilgjengelig_tid) × 100% | Flaskehalsindikator |
| **Gjennomsnittlig kølengde per WC** | Gj.snittlig antall ventende ordrer | Flaskehalsindikator |
| **Gjennomsnittlig køtid per WC** | Ventetid fra ankomst til behandling start | Servicenivå |
| **WIP-nivå** | Gj.snittlig antall ordrer i systemet | Kapitalbinding |
| **Leveringspresisjon** | % ordrer ferdig innen forfallsdato | Servicenivå |
| **Total kost per enhet** | (setup + operasjon + WIP-kost) / totalt_kvantum | **Hoved-optimeringsmål** |
| **Antall changeovers** | Totalt antall omstillinger i perioden | Operatørbelastning |

### 2E. Kostfunksjon for optimalisering

```
Total_kost(batch) = Setup_kost(batch) + Operasjons_kost + WIP_kost(batch) + Forsinkelses_kost(batch)
```

Der:

| Kostnadskomponent | Formel | Oppførsel når batch øker |
|---|---|---|
| **Setup_kost** | `antall_changeovers × Σ(setup_timer × timekost_per_wc)` | **Synker** — færre omstillinger |
| **Operasjons_kost** | `Σ(run_timer × timekost × totalt_kvantum)` | **Konstant** — samme arbeidsmengde |
| **WIP_kost** | `gj.snittlig_WIP × kapitalkost_per_tidsenhet` | **Øker** — mer bundet kapital |
| **Forsinkelses_kost** | `antall_forsinkede × straffekost_per_ordre` | **Øker** — lengre ledetid |

**Optimum:** Der marginalkostnaden av å øke batch = marginalbesparelsen i setup-kost.

<img src="https://upload.wikimedia.org/wikipedia/commons/thumb/2/2e/EOQ_Figure_1.png/640px-EOQ_Figure_1.png" alt="EOQ-kurve" width="400"/>

*Illustrasjon: Klassisk EOQ-kurve. Setup-kost (blå) synker, lagerkost (grønn) stiger. Totalkost (rød) har et minimum.*

---

## 3. Hvordan BOM og Routing brukes i simuleringen

### BOM sin rolle:
- **Materialkost:** `unit_cost / quantity_per × (1 + scrap_pct/100)` gir materialkost per output-enhet
- **Co-produkter:** Når BOM har `co_product_pct > 0`, genererer produksjonen både A-vare og B-vare samtidig — begge må spores i WIP
- **Svinn:** Reduserer reelt utbytte per batch — `batch_qty_effective = batch_qty × (1 − scrap_pct/100)`
- **Dynamisk cost roll-up:** Halvfabrikata (FG/HF-komponenter) som produseres i ett steg og forbrukes i neste må ha konsistente batch-størrelser

### Routing sin rolle:
- **Operasjonssekvens:** Definerer rekkefølgen ordren må gå gjennom — sorteres på `operation_no`
- **Work Center-tilknytning:** Hvilken maskin som brukes — dette er **ressursknutepunktet** i SimPy (`simpy.PriorityResource`)
- **Setup-tid:** `setup_time_minutes` — **hoveddriveren** for batch-optimalisering
- **Run-tid per enhet:** `run_time_minutes` — bestemmer hvor lenge en batch okkuperer work centeret: `run_time × batch_qty`

### Work Centers sin rolle:
- **Flaskehalsidentifisering:** Work centeret med høyest kapasitetsutnyttelse dikterer systemets totale gjennomløp
- **Kø:** Når flere produkter konkurrerer om samme work center, oppstår kø — små batcher = flere changeovers = mer kø
- **Kostnadsrater:** `total_cost_hour` (justert for `effective_capacity_pct`) brukes for å sette kronebeløp på setup-tid og køtid

---

## 4. Nye filer

| Fil | Størrelse | Formål |
|---|---|---|
| `batch_simulator.py` | ~500 linjer | SimPy-kjernelogikk |
| `demand_profile.py` | ~150 linjer | Etterspørselsmodellering |
| `batch_app.py` | ~400 linjer | Marimo-grensesnitt (port 8083) |
| `BATCH_SIMULERING_PLAN.md` | Denne filen | Implementasjonsplan |
| `requirements.txt` | +2 linjer | Legg til `simpy>=4.1.0`, `numpy` |

---

## 4A. `batch_simulator.py` — Kjernemodul (~500 linjer)

**Innhold:**

### Dataklasser:
- `SimulationStats` — Samler alle KPI-er under kjøring
- `BatchResult` — Resultatstruktur for én batch-størrelse
- `OptimizationResult` — Samling av BatchResult for alle testede batcher

### Kjernelogikk:
- `ProductionEnvironment` — Bygger SimPy-miljø med work center resources
- `production_order()` — SimPy-prosess for ordre-gjennomløp
- `order_generator()` — Etterspørselsgenerator (leser DemandProfile)
- `BatchOptimizer` — Orkestrerer hele simuleringsløpet:
  - `simulate_batch_size()` — Kjør én batch-størrelse med N replikasjoner
  - `evaluate_batch_sizes()` — Test flere batch-størrelser
  - `find_optimal()` — Returner batch med lavest totalkost
  - `_run_single_replication()` — Én enkelt SimPy-kjøring

### Avhengigheter:
```python
import simpy
import numpy as np
from dataclasses import dataclass, field
from typing import Optional
from kostberegning import SqliteData, CostCalculator, WorkCenter, RoutingLine
from demand_profile import DemandProfile
```

### BatchOptimizer API:
```python
class BatchOptimizer:
    def __init__(self, data: SqliteData, demand_profile: DemandProfile):
        ...

    def simulate_batch_size(
        self,
        product_no: str,
        batch_size: float,
        location_code: str = "",
        num_replications: int = 30,
        sim_duration_days: float = 365.0,
        warmup_days: float = 30.0,
    ) -> BatchResult:
        """Kjør N replikasjoner av én batch-størrelse."""
        ...

    def evaluate_batch_sizes(
        self,
        product_no: str,
        batch_sizes: list[float],
        location_code: str = "",
        num_replications: int = 30,
    ) -> OptimizationResult:
        """Test flere batch-størrelser og returner resultater."""
        ...

    def find_optimal(
        self,
        product_no: str,
        batch_range: tuple[float, float] = (50, 5000),
        steps: int = 20,
        location_code: str = "",
    ) -> tuple[float, OptimizationResult]:
        """Finn optimal batch-størrelse innenfor et intervall."""
        ...
```

### Må kalles og returneres til batch_app.py:
- `OptimizationResult` inneholder alle `BatchResult`-objekter for plotting
- `BatchResult` inneholder alle KPI-er for én batch-størrelse
- `find_optimal()` returnerer (optimal_batch, OptimizationResult)

---

## 4B. `demand_profile.py` — Etterspørselsdata (~150 linjer)

**Innhold:**

### Dataklasse:
- `DemandProfile` — Inneholder en liste med `(ankomst_tidspunkt_i_timer, kvantum)`

### Generator-funksjoner:
- `DemandGenerator.from_annual_volume(annual_total, num_orders, uom)` — Generer jevnt fordelte ordrer over ett år
- `DemandGenerator.from_annual_volume_seasonal(annual_total, monthly_weights)` — Generer med sesongvariasjon
- `DemandGenerator.from_poisson_arrivals(annual_total, avg_order_size)` — Stokastisk Poisson-ankomst
- `DemandGenerator.from_csv(filepath)` — Les historiske ordredata fra CSV
- `DemandGenerator.from_sqlite(db_path, product_no)` — Les historiske ordre fra database

### API:
```python
@dataclass
class DemandProfile:
    """Holder en liste med ordrer for simulering."""
    product_no: str
    orders: list[tuple[float, float]]  # (arrival_time_hours, quantity)
    total_quantity: float = 0.0
    num_orders: int = 0
    description: str = ""

class DemandGenerator:
    @staticmethod
    def from_annual_volume(annual_total: float, num_orders: int = 250) -> DemandProfile:
        """Jevnt fordelt over 250 arbeidsdager."""
        ...

    @staticmethod
    def from_annual_volume_seasonal(
        annual_total: float,
        monthly_weights: dict[int, float],  # {1: 0.05, 2: 0.06, ...}
    ) -> DemandProfile:
        """Sesongjustert fordeling."""
        ...

    @staticmethod
    def from_poisson_arrivals(
        annual_total: float,
        avg_order_size: float,
        seed: int = 42,
    ) -> DemandProfile:
        """Stokastisk Poisson-ankomst med eksponensielt fordelte ordrestørrelser."""
        ...
```

---

## 4C. `batch_app.py` — Marimo-grensesnitt (~400 linjer)

**Port:** 8083  
**Kjøres med:** `marimo run batch_app.py`

### Layout:

#### Celle 1: Imports og CSS
```python
import marimo as mo
from batch_simulator import BatchOptimizer
from demand_profile import DemandProfile, DemandGenerator
from kostberegning import SqliteData
import plotly.express as px
import pandas as pd
# ... CSS fra STYLING.md
```

#### Celle 2: Datalasting
```python
data = SqliteData()  # Leser fra endringslogg.db

# Produktvelger
product_options = {
    p.item_no: f"{p.item_no} · {p.description} · {p.base_uom}"
    for p in data.products
    if p.item_type in ("Finished Good", "Semi Finished")
}
selected_product = mo.ui.dropdown(
    options=product_options,
    label="Velg produkt",
)
```

#### Celle 3: Etterspørselskonfigurasjon
```python
# Valg av etterspørselsmodell
demand_model = mo.ui.radio(
    options={
        "uniform": "Jevnt fordelt over året",
        "seasonal": "Sesongbasert",
        "poisson": "Stokastisk (Poisson-ankomst)",
    },
    value="uniform",
    label="Etterspørselsmodell",
)

annual_volume = mo.ui.number(
    start=0, stop=10_000_000, value=200_000,
    label="Årlig volum",
)

num_replications = mo.ui.slider(
    start=5, stop=100, step=5, value=30,
    label="Antall replikasjoner per batch",
)
```

#### Celle 4: Batch-konfigurasjon
```python
# Definer hvilke batch-størrelser som skal testes
batch_sizes = mo.ui.array(
    [mo.ui.number(value=100), mo.ui.number(value=250),
     mo.ui.number(value=500), mo.ui.number(value=1000),
     mo.ui.number(value=2000), mo.ui.number(value=5000)],
    label="Batch-størrelser som testes",
)
```

#### Celle 5: Kjør simulering
```python
run_button = mo.ui.run_button(label="🚀 Kjør simulering")

# Progressbar og output
if run_button.value:
    with mo.output.replace():
        # Bygg DemandProfile
        # Kjør BatchOptimizer.evaluate_batch_sizes()
        # Vis progressbar under kjøring
        ...
```

#### Celle 6: Resultater — Tabell
```python
# Tabell: Batch-størrelse → Totalkost, Ledetid, Kap.utnyttelse, WIP, Antall changeovers
results_table = mo.ui.table(
    data=results_df,
    label="Simuleringsresultater per batch-størrelse",
)
```

#### Celle 7: Resultater — Diagrammer
- **Kostnadskurve:** `plotly.express.line` — batch-størrelse vs totalkost, med markert optimum
- **Kapasitetsutnyttelse:** `plotly.express.bar` — per work center per batch
- **Ledetidsfordeling:** `plotly.express.box` — per batch-størrelse
- **Kølengde over tid:** `plotly.express.line` — tidsakse for én replikasjon

#### Celle 8: Anbefaling og eksport
```python
# Vis optimal batch-størrelse med konfidensintervall
mo.md(f"""
## Anbefalt batch-størrelse: **{optimal_batch}** {uom}
- 95% konfidensintervall: {ci_lower} – {ci_upper}
- Estimert årlig besparelse vs. dagens batch: {savings:,.0f} kr
""")

# Eksport-knapper
export_button = mo.ui.button(label="📥 Eksporter resultater (CSV)")
```

---

## 5. Konkret eksempel: Høvleri-scenarioet

Fra `create_test_data()` i `kostberegning.py`:

| Work Center | Lokasjon | Labor | Machine | Overhead | Nominell t/kost | Eff. justert t/kost | Kap/dag (eff.) |
|---|---|---|---|---|---|---|---|
| HOVEDHOVEL | KOD | 550 | 900 | 150 | 1 600 | 1 882 kr/t | 13.6 t |
| SPESIALHOVEL | KOD | 550 | 950 | 150 | 1 650 | 1 941 kr/t | 13.6 t |
| KVHOVEL | KV | 500 | 200 | 100 | 800 | 941 kr/t | 13.6 t |

| Produkt FG001 — KOD-rute | Op | Work Center | Setup (min) | Run (min/enh) | Batch (nå) |
|---|---|---|---|---|---|
| 10: RIP | HOVEDHOVEL | 15 | 0.15 | 500 |
| 20: PLANING | HOVEDHOVEL | 10 | 0.10 | 500 |
| 30: PROFILE | SPESIALHOVEL | 20 | 0.12 | 500 |

| Produkt FG001 — KV-rute | Op | Work Center | Setup (min) | Run (min/enh) | Batch (nå) |
|---|---|---|---|---|---|
| 10: RIP | KVHOVEL | 15 | 0.15 | 1 000 |

### SimPy-test: Batch-størrelser som skal simuleres

`[100, 250, 500, 750, 1 000, 1 500, 2 000, 5 000]`

For hver batch-størrelse simuleres ett års produksjon (200 000 LM) med 30 replikasjoner.

### Forventede KPI-er per batch (illustrative tall):

| Batch | Setup-kost/år | Op.kost/år | WIP (snitt) | Ledetid (snitt t) | Kap.utn. HOVEDHOVEL | Totalkost/enhet |
|---|---|---|---|---|---|---|
| 100 | 1 666 000 | 800 000 | 5 | 2.1 | 98% | 14.23 |
| 250 | 666 000 | 800 000 | 8 | 3.8 | 94% | 12.45 |
| 500 | 333 000 | 800 000 | 12 | 6.2 | 88% | 11.78 |
| 1 000 | 166 000 | 800 000 | 18 | 10.5 | 82% | **11.22** ← OPTIMUM |
| 2 000 | 83 000 | 800 000 | 28 | 18.3 | 78% | 11.45 |
| 5 000 | 33 000 | 800 000 | 52 | 38.7 | 75% | 12.10 |

*Illustrative tall — faktiske resultater avhenger av reelle data.*

### Spørsmål simuleringen svarer på:

| Spørsmål | Beslutningsstøtte |
|---|---|
| "Hva er optimal batch-størrelse for FG001 på HOVEDHOVEL?" | Reduser setup-kost uten å skape flaskehalser |
| "Bør FG001 produseres på KOD eller KV?" | Sammenlign totalkost per lokasjon inkl. kø og kapasitetsutnyttelse |
| "Hvor mye koster det å kjøre små batcher (mer fleksibelt)?" | Avveining fleksibilitet vs. kostnad |
| "Hvilket work center er flaskehalsen ved ulike batcher?" | Investeringsbeslutninger (ny maskin, ekstra skift) |
| "Hvordan påvirker FG001s batch FG002s ventetid?" | Produktmiks-optimalisering |

---

## 6. Forretningsnytte

### 6.1 Den grunnleggende økonomiske mekanikken

I trelastindustrien er batch-optimalisering spesielt kraftfullt fordi **omstillingstidene er høye** og **kostnadene per time er betydelige**.

Setup-kost per omstilling for en høvel som koster 2 000 kr/t med 25 minutters omstilling:
```
25/60 × 2 000 = 833 kr per changeover
```

Ved 200 000 LM/år og batch = 500:
```
Setup-kost = 400 changeovers × 833 = 333 200 kr/år
```

Ved batch = 2 000:
```
Setup-kost = 100 changeovers × 833 = 83 300 kr/år
Bespartelse = 249 900 kr/år — for ETT produkt på ETT work center
```

**Men** — større batcher har skjulte kostnader: blokkerte maskiner, lengre kø, WIP-kapitalbinding, forsinkede leveranser. Dette fanges kun i en dynamisk SimPy-simulering — ikke i en statisk kalkyle.

### 6.2 Verdier SimPy gir som dagens system ikke kan

| Evne | Dagens CostCalculator | SimPy-simulering |
|---|---|---|
| Beregne setup-kost per enhet for én batch | ✅ | ✅ |
| Forutsi kø når batch A okkuperer maskin mens batch B venter | ❌ | ✅ |
| Vise hvordan én produkts batch påvirker et annet produkts ledetid | ❌ | ✅ |
| Avdekke flaskehalser i produktmiksen | ❌ | ✅ |
| Kvantifisere WIP-kapitalbinding som funksjon av batch | ❌ | ✅ |
| Teste sesongvariasjon | ❌ | ✅ |
| Sammenligne KOD vs KV i kapasitetsperspektiv | ❌ | ✅ |
| "Hvor mye koster fleksibilitet?" (små batcher) | ❌ | ✅ |

### 6.3 Konkrete beslutninger verktøyet støtter

| Beslutning | Årlig potensiell verdi |
|---|---|
| Batch-størrelser per produkt per work center | 250 000–500 000 kr |
| Lokasjonsvalg: KOD vs KV | 100 000–300 000 kr |
| Investering: ny maskin eller ekstra skift | 500 000–2 000 000 kr |
| Leveringstidsløfter til kunder | Kritisk, vanskelig å tallfeste |
| Produktmiks-optimalisering | 100 000–200 000 kr |
| Lagerstrategi: MTO vs MTS | 200 000–500 000 kr |

---

## 7. Forutsetninger for å lykkes

### 7.1 Organisatoriske

| Forutsetning | Risiko ved fravær | Tiltak |
|---|---|---|
| **Forankring hos produksjonsleder og operatører** | Anbefalinger ignoreres; verktøyet brukes ikke | Involver produksjonsleder i UI-design; intervju operatører om praktiske begrensninger (pallestørrelser, tørketider, bemanning) |
| **Eier av verktøyet som vedlikeholder data** | Modellen forvitrer på 6 måneder når data blir utdatert | Utpek én person; lag enkel prosedyre for oppdatering av setup-tider, kostnader og etterspørselsprofiler |
| **Villighet til å eksperimentere** | Simuleringens anbefalinger testes aldri i praksis | Start med pilot på ETT produkt; dokumenter faktisk vs. simulert resultat for å bygge tillit |
| **Forståelse for probabilistiske resultater** | Bruker forventer eksakte tall og mister tillit når to kjøringer gir ulike svar | Kommuniser resultater som "Batch 1 500–2 000 med 95% konfidens", ikke "Batch = 1 732". Vis konfidensintervaller i UI. |

### 7.2 Tekniske

| Forutsetning | Tiltak |
|---|---|
| **SimPy-kompetanse** | Skriv omfattende docstrings på norsk. Lag `batch_simulator_docs.md` med forklaring av alle parametere. |
| **Ytelse ved mange replikasjoner** | Bygg inn `multiprocessing` fra dag 1. Progressbar i Marimo. Mulighet for å avbryte. |
| **Integrasjon med SQLite** | Allerede på plass via `SqliteData` — ingen ny utvikling kreves. |

---

## 8. Data vi må ha stålkontroll på

### 🔴 KRITISK (uten disse gir simuleringen feil svar)

| Data | Hvorfor kritisk | Slik kvalitetssikres det |
|---|---|---|
| **Setup-tider** per operasjon per WC | Hoveddriver for batch-optimalisering. 15 min vs 8 min = dobling av anbefalt batch. | Stoppeklokke på 10–20 reelle omstillinger. Skill intern/ekstern setup. Dokumenter sequence-dependency. |
| **Run-tider** per enhet per operasjon | `run_time × batch_size` = okkupasjonstid. 10% feil = 10% feil i kapasitetsberegning. | Mål gjennomsnitt over minst 50 enheter, med standardavvik. Skill på produktvarianter (21×95 vs 28×200). |
| **Reell tilgjengelig kapasitet** per work center | Setter absolutt øvre grense. 15% feil i kapasitetsestimat = simuleringen er tilsvarende optimistisk/pessimistisk. | Inkluder ALT som stjeler tid: vedlikehold, verktøybytte, bemanning, mikrostopp. `capacity_hours_day` må reflektere faktisk skiftordning. |

### 🟡 VIKTIG (uten disse får vi upresise, men ikke gale svar)

| Data | Hvorfor viktig |
|---|---|
| **Etterspørselsprofil** (volum, variasjon, sesong) | Uten realistisk profil kan ikke simuleringen forutsi når flaskehalser oppstår. Batch optimal i januar kan være katastrofal i mai. |
| **Changeover-tider mellom ulike produkter** (sequence-dependency) | 21×95 glatt → 21×95 profilert er raskere enn 21×95 → 28×200. Uten dette går vi glipp av optimal sekvensiering. |

### 🟢 NYTTIG (uten disse går vi glipp av innsikt)

- Historiske leveringspresisjonsdata (kalibrer forsinkelseskost)
- Nedetidsdata per WC (MTBF/MTTR for maskinhavari)
- Bemanningsdata per skift (noen operasjoner krever 2 operatører)

---

## 9. Anbefalt tilnærming — Tre faser

### Fase 1: Pilot med ETT work center og TO produkter

Start med HOVEDHOVEL på KOD, med to produkter som konkurrerer om maskinen (f.eks. FG001 Utvendig Panel 21×95 og FG002 Innvendig Panel 15×95).

**Mål:** Validere at SimPy-modellen reproduserer observert adferd (ledetider, kølengder, kapasitetsutnyttelse).

**Datakrav:** Kun setup-tider, run-tider og kapasitet for HOVEDHOVEL og to produkter.

### Fase 2: Full modell

Etter validert pilot, utvid til alle work centers og hele produktspekteret. Legg til sequence-dependent changeover.

### Fase 3: Integrasjon med lokasjonsallokering

Kombiner med `ForslagSimulering.md` og `VERDIKJEDE_PLAN.md`: Simuleringen viser ikke bare optimal batch, men også hvilken fabrikk (KOD vs KV) som bør produsere hva.

---

## 10. Trinnvis implementasjonsplan

| Trinn | Fil | Hva som skal bygges | Avhengigheter |
|---|---|---|---|
| **1** | `demand_profile.py` | `DemandProfile`, `DemandGenerator` med 3 metoder | `numpy` |
| **2** | `batch_simulator.py` | `SimulationStats`, `BatchResult`, `OptimizationResult`, `ProductionEnvironment`, `production_order()`, `order_generator()`, `BatchOptimizer` | `simpy`, `demand_profile`, `kostberegning.SqliteData` |
| **3** | `requirements.txt` | Legg til `simpy>=4.1.0` og evt. `numpy` | — |
| **4** | `batch_app.py` | Marimo-grensesnitt: produktvelger, etterspørselskonfig, batch-velger, simulering, resultater, diagrammer, eksport | `batch_simulator`, `marimo`, `plotly` |
| **5** | Test | Kjør pilot: FG001 på HOVEDHOVEL, 8 batch-størrelser, 30 replikasjoner | Alle moduler |
| **6** | Validering | Sammenlign simulerte ledetider med faktiske observasjoner | Produksjonsdata |

### Trinn 1 — `demand_profile.py`

```python
"""
demand_profile.py — Etterspørselsmodellering for batch-simulering.

Definerer hvordan produksjonsordrer ankommer systemet over tid.
Støtter:
- Jevn fordeling over året (fra årsvolum)
- Sesongjustert fordeling (månedsvekter)
- Stokastisk Poisson-ankomst
- Import fra CSV (historiske ordre)
"""

import math
import random
from dataclasses import dataclass
from pathlib import Path

@dataclass
class DemandProfile:
    product_no: str
    orders: list[tuple[float, float]]  # (arrival_time_hours, quantity)
    total_quantity: float = 0.0
    num_orders: int = 0
    description: str = ""

class DemandGenerator:
    WORK_HOURS_PER_DAY = 16.0
    WORK_DAYS_PER_YEAR = 250

    @staticmethod
    def from_annual_volume(
        annual_total: float,
        num_orders: int = 250,
        product_no: str = "",
    ) -> DemandProfile:
        """Jevnt fordelte ordrer over arbeidsdagene i ett år."""
        total_hours = num_orders * DemandGenerator.WORK_HOURS_PER_DAY
        hours_between = total_hours / num_orders if num_orders > 0 else 0
        qty_per_order = annual_total / num_orders if num_orders > 0 else 0

        orders = [
            (i * hours_between, qty_per_order)
            for i in range(num_orders)
        ]
        return DemandProfile(
            product_no=product_no,
            orders=orders,
            total_quantity=annual_total,
            num_orders=num_orders,
            description=f"Jevnt fordelt: {num_orders} ordrer à {qty_per_order:.0f}",
        )

    @staticmethod
    def from_annual_volume_seasonal(
        annual_total: float,
        monthly_weights: dict[int, float],  # {1: 0.05, 2: 0.06, ..., 12: 0.10}
        product_no: str = "",
    ) -> DemandProfile:
        """Sesongjustert fordeling basert på månedsvekter."""
        # Valider at vektene summerer til 1.0
        total_weight = sum(monthly_weights.values())
        normalized = {m: w / total_weight for m, w in monthly_weights.items()}

        orders = []
        hours_per_month = (DemandGenerator.WORK_DAYS_PER_YEAR / 12) * DemandGenerator.WORK_HOURS_PER_DAY

        for month in range(1, 13):
            weight = normalized.get(month, 1.0 / 12)
            month_volume = annual_total * weight
            month_start = (month - 1) * hours_per_month
            # ~21 ordrer per måned (én per arbeidsdag)
            month_orders = max(1, round(weight * 250 / 12))
            qty_per_order = month_volume / month_orders

            for i in range(month_orders):
                arrival = month_start + (i * hours_per_month / month_orders)
                orders.append((arrival, qty_per_order))

        return DemandProfile(
            product_no=product_no,
            orders=sorted(orders, key=lambda x: x[0]),
            total_quantity=annual_total,
            num_orders=len(orders),
            description=f"Sesongjustert: {len(orders)} ordrer over 12 mnd",
        )

    @staticmethod
    def from_poisson_arrivals(
        annual_total: float,
        avg_order_size: float,
        product_no: str = "",
        seed: int = 42,
    ) -> DemandProfile:
        """Stokastisk Poisson-ankomst med eksponensielt fordelte ordrestørrelser."""
        rng = random.Random(seed)
        total_hours = DemandGenerator.WORK_DAYS_PER_YEAR * DemandGenerator.WORK_HOURS_PER_DAY
        avg_orders = annual_total / avg_order_size if avg_order_size > 0 else 250
        # Gjennomsnittlig tid mellom ordrer
        avg_interarrival = total_hours / avg_orders if avg_orders > 0 else 0

        orders = []
        current_time = 0.0

        while current_time < total_hours:
            # Poisson-ankomst: eksponensielt fordelt tid mellom ordrer
            interarrival = rng.expovariate(1.0 / avg_interarrival) if avg_interarrival > 0 else 0
            current_time += interarrival
            if current_time >= total_hours:
                break
            # Ordrestørrelse: eksponensielt fordelt
            order_qty = rng.expovariate(1.0 / avg_order_size) if avg_order_size > 0 else avg_order_size
            orders.append((current_time, order_qty))

        return DemandProfile(
            product_no=product_no,
            orders=orders,
            total_quantity=sum(q for _, q in orders),
            num_orders=len(orders),
            description=f"Poisson: λ={avg_interarrival:.1f}t mellom ordrer, snitt {avg_order_size:.0f}/ordre",
        )
```

### Trinn 2 — `batch_simulator.py`

```python
"""
batch_simulator.py — SimPy-basert batch-optimalisering for trelastproduksjon.

Bruker BOM, Routing og Work Centers fra SqliteData til å bygge en
discrete-event simulering av produksjonsflyten.

Brukes av:
    - batch_app.py (Marimo-grensesnitt)

Avhengigheter:
    - simpy (discrete-event simulation)
    - numpy (statistikk)
    - kostberegning.SqliteData, CostCalculator (lese-avhengigheter)
    - demand_profile.DemandProfile
"""

import math
import multiprocessing as mp
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import simpy

from kostberegning import SqliteData, WorkCenter, RoutingLine
from demand_profile import DemandProfile


# ── Datastrukturer ────────────────────────────────────────────────

@dataclass
class QueueEvent:
    """En kø-hendelse for logging."""
    time: float
    work_center: str
    queue_length: int


@dataclass
class CompletionRecord:
    """Registrering av en ferdigstilt ordre."""
    order_id: str
    product_no: str
    arrival_time: float
    completion_time: float
    batch_qty: float
    lead_time: float


@dataclass
class SimulationStats:
    """Samler alle KPI-er under en SimPy-kjøring."""

    # Rådata
    completions: list[CompletionRecord] = field(default_factory=list)
    queue_events: list[QueueEvent] = field(default_factory=list)
    setup_events: list[tuple[str, float, str]] = field(default_factory=list)  # (wc, time_min, product)
    run_events: list[tuple[str, float, float]] = field(default_factory=list)  # (wc, hours, batch_qty)

    # Akkumulerte verdier
    total_setup_time: dict[str, float] = field(default_factory=dict)  # wc -> timer
    total_run_time: dict[str, float] = field(default_factory=dict)  # wc -> timer
    total_setup_count: dict[str, int] = field(default_factory=dict)  # wc -> antall
    total_quantity_produced: float = 0.0

    def record_queue_time(self, work_center: str, queue_time: float, order_id: str):
        """Registrer at en ordre ventet i kø."""
        self.queue_events.append(QueueEvent(
            time=queue_time,
            work_center=work_center,
            queue_length=0,  # Må oppdateres fra miljøet
        ))

    def record_setup(self, work_center: str, setup_minutes: float, product_no: str):
        """Registrer en omstilling."""
        self.setup_events.append((work_center, setup_minutes, product_no))
        if work_center not in self.total_setup_time:
            self.total_setup_time[work_center] = 0.0
            self.total_setup_count[work_center] = 0
        self.total_setup_time[work_center] += setup_minutes / 60.0
        self.total_setup_count[work_center] += 1

    def record_run(self, work_center: str, run_hours: float, batch_qty: float):
        """Registrer kjøretid."""
        self.run_events.append((work_center, run_hours, batch_qty))
        if work_center not in self.total_run_time:
            self.total_run_time[work_center] = 0.0
        self.total_run_time[work_center] += run_hours
        self.total_quantity_produced += batch_qty

    def record_completion(self, order_id: str, product_no: str,
                          arrival_time: float, completion_time: float,
                          batch_qty: float, lead_time: float):
        """Registrer en ferdigstilt ordre."""
        self.completions.append(CompletionRecord(
            order_id=order_id,
            product_no=product_no,
            arrival_time=arrival_time,
            completion_time=completion_time,
            batch_qty=batch_qty,
            lead_time=lead_time,
        ))

    # ── Beregnede KPI-er ──────────────────────────────────────

    @property
    def avg_lead_time(self) -> float:
        """Gjennomsnittlig ledetid i timer."""
        if not self.completions:
            return 0.0
        return np.mean([c.lead_time for c in self.completions])

    @property
    def avg_lead_time_days(self) -> float:
        """Gjennomsnittlig ledetid i dager (16t/dag)."""
        return self.avg_lead_time / 16.0

    @property
    def lead_time_std(self) -> float:
        """Standardavvik ledetid."""
        if not self.completions:
            return 0.0
        return np.std([c.lead_time for c in self.completions])

    @property
    def total_setup_hours(self) -> float:
        """Total setup-tid i timer."""
        return sum(self.total_setup_time.values())

    @property
    def total_run_hours(self) -> float:
        """Total kjøretid i timer."""
        return sum(self.total_run_time.values())

    @property
    def num_changeovers(self) -> int:
        """Totalt antall omstillinger."""
        return sum(self.total_setup_count.values())


@dataclass
class BatchResult:
    """Resultat for én batch-størrelse (aggregert over N replikasjoner)."""
    batch_size: float
    num_replications: int

    # Gjennomsnitt over replikasjoner
    avg_total_cost_per_unit: float = 0.0
    avg_setup_cost_per_unit: float = 0.0
    avg_operation_cost_per_unit: float = 0.0
    avg_wip_cost_per_unit: float = 0.0
    avg_lead_time_days: float = 0.0
    avg_lead_time_std: float = 0.0
    avg_num_changeovers: float = 0.0
    avg_capacity_utilization: dict[str, float] = field(default_factory=dict)  # wc -> %

    # Konfidensintervall for totalkost
    ci_lower: float = 0.0
    ci_upper: float = 0.0

    # Rådata per replikasjon (for plotting)
    replication_costs: list[float] = field(default_factory=list)
    replication_lead_times: list[float] = field(default_factory=list)


@dataclass
class OptimizationResult:
    """Samling av BatchResult for alle testede batch-størrelser."""
    product_no: str
    location_code: str
    results: list[BatchResult] = field(default_factory=list)

    @property
    def optimal_batch(self) -> tuple[float, float, float]:
        """Returner (batch, kost, ci_lower, ci_upper) for optimum."""
        if not self.results:
            return 0.0, 0.0, 0.0, 0.0
        best = min(self.results, key=lambda r: r.avg_total_cost_per_unit)
        return best.batch_size, best.avg_total_cost_per_unit, best.ci_lower, best.ci_upper

    @property
    def best_result(self) -> Optional[BatchResult]:
        """Returner BatchResult med lavest totalkost."""
        if not self.results:
            return None
        return min(self.results, key=lambda r: r.avg_total_cost_per_unit)


# ── SimPy Prosesfunksjoner ─────────────────────────────────────

def production_order(env, order_id, product_no, batch_qty,
                     routing_lines, work_centers, wc_meta, stats):
    """
    SimPy-prosess for én produksjonsordre.

    1. Ankommer systemet
    2. For hver operasjon i routing:
       a. Be om work center (med prioritet)
       b. Setup-tid (omstilling)
       c. Kjøretid = run_time_minutes * batch_qty (i timer)
    3. Registrer ferdigstillelse
    """
    arrival_time = env.now
    priority = 1  # Standardprioritet

    for rl in routing_lines:
        wc = work_centers[rl.work_center_code]

        # Be om tilgang til work center
        with wc.request(priority=priority) as req:
            yield req

            # Ventetid i kø
            queue_time = env.now - arrival_time
            stats.record_queue_time(rl.work_center_code, queue_time, order_id)

            # Setup
            setup_hours = rl.setup_time_minutes / 60.0
            yield env.timeout(setup_hours)
            stats.record_setup(rl.work_center_code, rl.setup_time_minutes, product_no)

            # Kjøretid
            run_hours = (rl.run_time_minutes / 60.0) * batch_qty
            yield env.timeout(run_hours)
            stats.record_run(rl.work_center_code, run_hours, batch_qty)

    # Ferdigstillelse
    completion_time = env.now
    stats.record_completion(
        order_id=order_id,
        product_no=product_no,
        arrival_time=arrival_time,
        completion_time=completion_time,
        batch_qty=batch_qty,
        lead_time=completion_time - arrival_time,
    )


def order_generator(env, product_no, batch_size, demand_profile,
                    routing_lines, work_centers, wc_meta, stats):
    """Generer produksjonsordrer basert på etterspørselsprofil."""

    def generate_batches(quantity):
        num_batches = math.ceil(quantity / batch_size)
        if num_batches == 0:
            return
        for b in range(num_batches):
            yield min(batch_size, quantity - b * batch_size)

    for i, (arrival_time, quantity) in enumerate(demand_profile.orders):
        if arrival_time > env.now:
            yield env.timeout(arrival_time - env.now)

        for j, batch_qty in enumerate(generate_batches(quantity)):
            order_id = f"{product_no}_ord{i:04d}_b{j:03d}"
            env.process(production_order(
                env, order_id, product_no, batch_qty,
                routing_lines, work_centers, wc_meta, stats
            ))


# ── BatchOptimerer ────────────────────────────────────────────

class BatchOptimizer:
    """
    Orkestrerer SimPy-simulering for batch-optimalisering.

    Bruk:
        optimizer = BatchOptimizer(data, demand_profile)
        result = optimizer.evaluate_batch_sizes("FG001", [100, 500, 1000])
        optimal_batch, optimal_cost = optimizer.find_optimal("FG001")
    """

    def __init__(self, data: SqliteData, demand_profile: DemandProfile):
        self.data = data
        self.demand_profile = demand_profile

        # Bygg work center metadata
        self.wc_meta: dict[str, WorkCenter] = {
            wc.code: wc for wc in data.work_centers if wc.active
        }

    def _get_routing(self, product_no: str, location_code: str = "") -> list[RoutingLine]:
        """Hent routing for et produkt, filtrert på lokasjon."""
        routing = self.data.routing_for(product_no)
        if not routing:
            return []

        if location_code:
            routing = [
                r for r in routing
                if self.data.work_center(r.work_center_code)
                and self.data.work_center(r.work_center_code).location_code == location_code
            ]

        return sorted(routing, key=lambda r: r.operation_no)

    def _run_single_replication(
        self,
        product_no: str,
        batch_size: float,
        routing_lines: list[RoutingLine],
        sim_duration_hours: float,
        warmup_hours: float,
        seed: int,
    ) -> SimulationStats:
        """Kjør én enkelt SimPy-replikasjon."""
        env = simpy.Environment()
        stats = SimulationStats()

        # Bygg work center resources
        work_centers = {
            code: simpy.PriorityResource(env, capacity=1)
            for code in self.wc_meta
            if code in {rl.work_center_code for rl in routing_lines}
        }

        # Start ordregenerator
        env.process(order_generator(
            env, product_no, batch_size, self.demand_profile,
            routing_lines, work_centers, self.wc_meta, stats
        ))

        # Kjør simulering
        env.run(until=sim_duration_hours)

        return stats

    def simulate_batch_size(
        self,
        product_no: str,
        batch_size: float,
        location_code: str = "",
        num_replications: int = 30,
        sim_duration_days: float = 365.0,
        warmup_days: float = 30.0,
    ) -> BatchResult:
        """
        Kjør N replikasjoner av én batch-størrelse.

        Args:
            product_no: Produktnummer (f.eks. "FG001")
            batch_size: Batch-størrelse i output-enheter (f.eks. 500 LM)
            location_code: Fabrikkode (tom = alle)
            num_replications: Antall replikasjoner (30+ anbefales)
            sim_duration_days: Simuleringslengde i dager
            warmup_days: Oppvarmingsperiode (resultater forkastes)

        Returns:
            BatchResult med aggregerte KPI-er
        """
        routing = self._get_routing(product_no, location_code)
        if not routing:
            raise ValueError(f"Ingen routing funnet for {product_no} på {location_code or 'alle'}")

        sim_hours = sim_duration_days * 16.0  # 16 timer per dag
        warmup_hours = warmup_days * 16.0

        # Kjør replikasjoner
        all_replication_stats = []
        for rep in range(num_replications):
            stats = self._run_single_replication(
                product_no, batch_size, routing, sim_hours, warmup_hours, seed=rep
            )
            all_replication_stats.append(stats)

        # Aggreger resultater
        result = BatchResult(
            batch_size=batch_size,
            num_replications=num_replications,
        )

        # Beregn kostnader per replikasjon
        replication_costs = []
        for stats in all_replication_stats:
            total_qty = stats.total_quantity_produced
            if total_qty == 0:
                continue

            # Setup-kost
            setup_cost = sum(
                stats.total_setup_time.get(wc_code, 0.0) * (
                    self.wc_meta[wc_code].total_cost_hour if wc_code in self.wc_meta else 0
                )
                for wc_code in stats.total_setup_time
            )

            # Operasjonskost
            operation_cost = sum(
                stats.total_run_time.get(wc_code, 0.0) * (
                    self.wc_meta[wc_code].total_cost_hour if wc_code in self.wc_meta else 0
                )
                for wc_code in stats.total_run_time
            )

            total_cost = setup_cost + operation_cost
            cost_per_unit = total_cost / total_qty
            replication_costs.append(cost_per_unit)

        if replication_costs:
            costs = np.array(replication_costs)
            result.avg_total_cost_per_unit = float(np.mean(costs))
            result.avg_setup_cost_per_unit = float(np.mean([
                sum(stats.total_setup_time.get(wc, 0) * self.wc_meta.get(wc, WorkCenter("", "", "", 0)).total_cost_hour
                    for wc in stats.total_setup_time) / max(stats.total_quantity_produced, 1)
                for stats in all_replication_stats
            ]))
            result.replication_costs = replication_costs

            # 95% konfidensintervall
            if len(costs) > 1:
                sem = np.std(costs, ddof=1) / np.sqrt(len(costs))
                result.ci_lower = float(np.mean(costs) - 1.96 * sem)
                result.ci_upper = float(np.mean(costs) + 1.96 * sem)

        # Ledetider
        all_lead_times = []
        for stats in all_replication_stats:
            all_lead_times.extend([c.lead_time for c in stats.completions])
        if all_lead_times:
            result.avg_lead_time_days = float(np.mean(all_lead_times)) / 16.0
            result.avg_lead_time_std = float(np.std(all_lead_times))

        # Changeovers
        result.avg_num_changeovers = float(np.mean([
            stats.num_changeovers for stats in all_replication_stats
        ]))

        # Kapasitetsutnyttelse per WC
        total_available_hours = sim_hours  # - warmup_hours (forenklet)
        for wc_code in self.wc_meta:
            wc = self.wc_meta[wc_code]
            total_used = float(np.mean([
                stats.total_run_time.get(wc_code, 0) + stats.total_setup_time.get(wc_code, 0)
                for stats in all_replication_stats
            ]))
            result.avg_capacity_utilization[wc_code] = (
                (total_used / total_available_hours) * 100.0 if total_available_hours > 0 else 0.0
            )

        return result

    def evaluate_batch_sizes(
        self,
        product_no: str,
        batch_sizes: list[float],
        location_code: str = "",
        num_replications: int = 30,
    ) -> OptimizationResult:
        """
        Test flere batch-størrelser.

        Args:
            product_no: Produktnummer
            batch_sizes: Liste med batch-størrelser som skal testes
            location_code: Fabrikkode
            num_replications: Antall replikasjoner per batch

        Returns:
            OptimizationResult med resultater for alle batcher
        """
        results = []
        for i, batch in enumerate(batch_sizes):
            result = self.simulate_batch_size(
                product_no=product_no,
                batch_size=batch,
                location_code=location_code,
                num_replications=num_replications,
            )
            results.append(result)

        return OptimizationResult(
            product_no=product_no,
            location_code=location_code,
            results=results,
        )

    def find_optimal(
        self,
        product_no: str,
        batch_range: tuple[float, float] = (50, 5000),
        steps: int = 20,
        location_code: str = "",
        num_replications: int = 30,
    ) -> OptimizationResult:
        """
        Finn optimal batch-størrelse innenfor et intervall.

        Genererer `steps` jevnt fordelte batch-størrelser i
        logaritmisk skala (bedre dekning av lave verdier).

        Args:
            product_no: Produktnummer
            batch_range: (min_batch, max_batch)
            steps: Antall batch-størrelser som testes
            location_code: Fabrikkode
            num_replications: Antall replikasjoner per batch

        Returns:
            OptimizationResult med optimal batch markert
        """
        # Logaritmisk fordeling for bedre dekning av lave batcher
        log_min = np.log10(batch_range[0])
        log_max = np.log10(batch_range[1])
        batch_sizes = np.logspace(log_min, log_max, steps)
        batch_sizes = [round(b) for b in batch_sizes]

        return self.evaluate_batch_sizes(
            product_no=product_no,
            batch_sizes=batch_sizes,
            location_code=location_code,
            num_replications=num_replications,
        )
```

### Trinn 4 — `batch_app.py` (skisse)

```python
"""
batch_app.py — Marimo-app for batch-optimalisering.

Kjøres på port 8083:
    marimo run batch_app.py -p 8083
"""

import marimo as mo
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np

from batch_simulator import BatchOptimizer
from demand_profile import DemandProfile, DemandGenerator
from kostberegning import SqliteData

# ── CSS (gjenbruk fra STYLING.md) ─────────────────────────────

mo.md("""
<style>
    /* Samme stil som varekost_app.py */
    .fti-card { ... }
    .fti-highlight-green { ... }
    .fti-highlight-red { ... }
</style>
""")

# ── Header ─────────────────────────────────────────────────────

mo.md("""
# 🔧 Batch-optimalisering — SimPy-simulering

Optimaliser batch-størrelser for produksjon på tvers av arbeidssentre.
Simuleringen bruker diskret hendelses-simulering (SimPy) for å modellere
kø, kapasitetsutnyttelse og omstillingskostnader.
""")

# ── Datalasting ────────────────────────────────────────────────

data = SqliteData()

# Produktvelger
product_options = {
    p.item_no: f"{p.item_no} · {p.description} · {p.base_uom}"
    for p in data.products
    if p.item_type in ("Finished Good", "Semi Finished")
}
selected_product = mo.ui.dropdown(
    options=product_options,
    value=list(product_options.keys())[0] if product_options else None,
    label="Velg produkt",
)
mo.md("## 1. Velg produkt").callout()
selected_product

# ── Etterspørselskonfigurasjon ─────────────────────────────────

mo.md("## 2. Konfigurer etterspørsel").callout()

demand_model = mo.ui.radio(
    options={
        "uniform": "Jevnt fordelt over året",
        "seasonal": "Sesongbasert (høy vår/sommer)",
        "poisson": "Stokastisk (Poisson-ankomst)",
    },
    value="uniform",
    label="Etterspørselsmodell",
)

annual_volume = mo.ui.number(
    start=1000, stop=10_000_000, step=10000, value=200000,
    label="Årlig volum (enheter)",
)

num_orders = mo.ui.slider(
    start=50, stop=500, step=10, value=250,
    label="Antall ordrer per år",
)

num_replications = mo.ui.slider(
    start=5, stop=100, step=5, value=30,
    label="Antall replikasjoner per batch",
)

demand_model
annual_volume
num_orders
num_replications

# ── Batch-konfigurasjon ────────────────────────────────────────

mo.md("## 3. Velg batch-størrelser som testes").callout()

batch_sizes_input = mo.ui.text_area(
    value="100, 250, 500, 750, 1000, 1500, 2000, 5000",
    label="Batch-størrelser (kommaseparert)",
)

location_filter = mo.ui.dropdown(
    options={"": "Alle fabrikker", **{l.code: l.name for l in data.locations}},
    value="",
    label="Filtrer på fabrikk",
)

batch_sizes_input
location_filter

# ── Kjør simulering ────────────────────────────────────────────

mo.md("## 4. Kjør simulering").callout()

run_button = mo.ui.run_button(label="🚀 Kjør simulering")
run_button

if run_button.value:
    # Parse batch-størrelser
    try:
        batch_list = [float(b.strip()) for b in batch_sizes_input.value.split(",")]
    except ValueError:
        mo.md("❌ **Feil:** Kunne ikke parse batch-størrelser. Bruk komma-separerte tall.").callout()
        batch_list = []

    if batch_list and selected_product.value:
        product_no = selected_product.value
        product = data.product(product_no)
        uom = product.base_uom if product else ""

        # Bygg DemandProfile
        if demand_model.value == "uniform":
            demand = DemandGenerator.from_annual_volume(
                annual_total=annual_volume.value,
                num_orders=num_orders.value,
                product_no=product_no,
            )
        elif demand_model.value == "seasonal":
            # Standard sesongvekter for byggevarer: høyt mai-august
            seasonal_weights = {
                1: 0.04, 2: 0.04, 3: 0.06, 4: 0.10, 5: 0.14,
                6: 0.15, 7: 0.10, 8: 0.12, 9: 0.10, 10: 0.08,
                11: 0.05, 12: 0.02,
            }
            demand = DemandGenerator.from_annual_volume_seasonal(
                annual_total=annual_volume.value,
                monthly_weights=seasonal_weights,
                product_no=product_no,
            )
        else:
            avg_order_size = annual_volume.value / num_orders.value if num_orders.value > 0 else 800
            demand = DemandGenerator.from_poisson_arrivals(
                annual_total=annual_volume.value,
                avg_order_size=avg_order_size,
                product_no=product_no,
            )

        # Kjør BatchOptimizer
        with mo.output.replace():
            mo.md(f"### ⏳ Kjører simulering for {len(batch_list)} batch-størrelser...").callout()

            optimizer = BatchOptimizer(data, demand)
            opt_result = optimizer.evaluate_batch_sizes(
                product_no=product_no,
                batch_sizes=batch_list,
                location_code=location_filter.value,
                num_replications=num_replications.value,
            )

            mo.md("### ✅ Simulering fullført!").callout()

            # ── Resultater ──────────────────────────────────

            mo.md("## 5. Resultater").callout()

            # Tabell
            results_data = []
            for r in opt_result.results:
                results_data.append({
                    "Batch": f"{r.batch_size:,.0f} {uom}",
                    "Totalkost/enhet": f"{r.avg_total_cost_per_unit:.4f} kr",
                    "Setup-kost/enhet": f"{r.avg_setup_cost_per_unit:.4f} kr",
                    "Ledetid (dager)": f"{r.avg_lead_time_days:.1f}",
                    "Antall omstillinger": f"{r.avg_num_changeovers:.0f}",
                    "95% KI": f"[{r.ci_lower:.4f} – {r.ci_upper:.4f}]",
                })

            results_df = pd.DataFrame(results_data)
            mo.ui.table(data=results_df, label="Simuleringsresultater")

            # Diagram: Kostnadskurve
            batch_vals = [r.batch_size for r in opt_result.results]
            cost_vals = [r.avg_total_cost_per_unit for r in opt_result.results]
            ci_lower_vals = [r.ci_lower for r in opt_result.results]
            ci_upper_vals = [r.ci_upper for r in opt_result.results]

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=batch_vals, y=cost_vals,
                mode='lines+markers',
                name='Totalkost per enhet',
                line=dict(color='#14532D', width=2),
            ))
            fig.add_trace(go.Scatter(
                x=batch_vals + batch_vals[::-1],
                y=ci_upper_vals + ci_lower_vals[::-1],
                fill='toself',
                fillcolor='rgba(20, 83, 45, 0.2)',
                line=dict(color='rgba(255,255,255,0)'),
                name='95% konfidensintervall',
            ))

            optimal = opt_result.optimal_batch
            fig.add_vline(x=optimal[0], line_dash="dash", line_color="#2F855A",
                         annotation_text=f"Optimal: {optimal[0]:,.0f} {uom}")

            fig.update_layout(
                title=f"Kostnadskurve — {product_no}: {product.description if product else ''}",
                xaxis_title=f"Batch-størrelse ({uom})",
                yaxis_title="Totalkost per enhet (kr)",
                template="plotly_white",
            )
            mo.ui.plotly(fig)

            # Anbefaling
            best = opt_result.best_result
            if best:
                mo.md(f"""
                ## 📊 Anbefaling

                | Parameter | Verdi |
                |---|---|
                | **Anbefalt batch-størrelse** | **{optimal[0]:,.0f} {uom}** |
                | Forventet kostnad | {optimal[1]:.4f} kr/{uom} |
                | 95% konfidensintervall | {optimal[2]:.4f} – {optimal[3]:.4f} kr/{uom} |
                | Gj.snittlig ledetid | {best.avg_lead_time_days:.1f} dager |
                | Antall omstillinger/år | {best.avg_num_changeovers:.0f} |
                """).callout()

            # Eksport
            mo.download(
                data=results_df.to_csv(index=False).encode('utf-8'),
                filename=f"batch_optimalisering_{product_no}.csv",
                mimetype="text/csv",
                label="📥 Last ned resultater (CSV)",
            )
```

---

## 11. Tekniske garantier (hva røres IKKE)

| Fil | Status | Merknad |
|---|---|---|
| `kostberegning.py` | **URØRT** | Importeres kun som `SqliteData`, `CostCalculator`, `WorkCenter`, `RoutingLine` |
| `varekost_app.py` | **URØRT** | Egen app på egen port |
| `data_repo.py` | **URØRT** | Ingen nye tabeller i første fase |
| `excel_bridge.py` | **URØRT** | Ingen nye ark å importere |
| `generer_pdf_rapport.py` | **URØRT** | Kan utvides senere for batch-rapporter |
| `generer_excel_rapport.py` | **URØRT** | Kan utvides senere for batch-eksport |

### Nye avhengigheter

```
simpy>=4.1.0    # Discrete-event simulation
numpy>=1.24.0    # Statistikk (allerede implisitt via pandas, men eksplisitt spesifisert)
```

---

## Tre spørsmål ledelsen må svare på før oppstart

1. **Har vi målte setup- og run-tider, eller er dette anslag fra operatørenes hukommelse?** Hvis anslag — start med tidtaking.

2. **Er produksjonsleder villig til å teste anbefalte batch-størrelser i praksis?** Simuleringen har null verdi hvis resultatene ikke brukes.

3. **Finnes det én person som kan eie verktøyet og oppdatere data ved endringer?** Uten dataeier forvitrer modellen på 6 måneder.