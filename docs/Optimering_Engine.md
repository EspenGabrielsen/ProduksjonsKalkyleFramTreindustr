# Optimering_Engine — Produksjonsplanlegging med MILP

## Oversikt

`optimization_engine.py` er en frittstående MILP-motor (PuLP/CBC) som finner den mest lønn­somme produksjonsplanen for en gitt etterspørsel, på tvers av:

- **Flernivå-BOM** — Skurlast → Ubehandlet → Grunnet → Malt
- **Flere lokasjoner** — hvor skal hvert steg produseres?
- **Flere perioder** — når skal det produseres, og hvor mye lager?
- **Transport** — mellom fabrikker
- **Kapasitet** — per arbeidssenter per uke

Eksisterende kode (`kostberegning.py`, `varekost_app.py`, `excel_bridge.py`) endres ALDRI. Motoren importerer og gjenbruker `SqliteData` + `CostCalculator`.

## Kommandolinje

```
python src/optimization_engine.py --db src/produksjonskalkyle.db --output resultat.json --sequence
```

| Parameter | Default | Hva den gjør |
|-----------|---------|--------------|
| `--db` | `src/produksjonskalkyle.db` | Sti til SQLite-database |
| `--output` | — | Skriv resultat til JSON-fil |
| `--days` | 5.0 | Arbeidsdager per periode (kapasitet per uke) |
| `--holding-cost-pct` | 2.0 | Lagerholdskost i % av enhetskost per uke |
| `--forecast-weeks` | 12 | Prognosehorisont fra historisk salg (brukes til min_batch-beregning) |
| `--forecast-max-setup-pct` | 30 | Maks setup-andel av enhetskost for batch-beregning |
| `--time-limit` | — | Maks løsningstid i sekunder |
| `--sequence` | av | Beregn optimal produksjonssekvens (TSP) → JSON |
| `--no-min-batch` | — | Skru AV minste batchstørrelse (tillater mikrobatcher) |
| `--no-transport` | — | Utelat transportvariabler |

## Prinsipper

1. **Kun faktisk demand** — historisk salg genererer IKKE etterspørsel.
2. **Historisk salg styrer batch-størrelse** — via `min_batch`-skranken.
3. **Materialbalanse med lager** — modellen kan produsere tidlig og dekke fremtidig etterspørsel via lager.
4. **Setup-kost amortiseres** — `setup_per_unit = fixed_setup / faktisk_batch`.
5. **Dynamisk Big-M** — per produkt/periode, numerisk stabil.

## Batch-størrelse (`min_batch`)

### Hvordan den beregnes (`_compute_dynamic_min_batch`)

1. Les historisk salg fra `historical_sales`.
2. Kjør **top-down BOM-ekspansjon** (sluttprodukt → råvare) med `quantity_per` og `scrap%`.
3. Aggreger til årlig volum per produkt → ukentlig snitt: `uke_snitt = årlig / 52`.
4. Beregn økonomisk minimumsbatch:

   ```
   min_mengde = fixed_setup / (unit_cost × setup_pct)
   n_uker     = ceil(min_mengde / uke_snitt)
   min_batch  = uke_snitt × n_uker
   ```

5. Kun produkter med **routing** (produserbare varer) får `min_batch` — råvarer styres av materialbalansen.

### Eksempel: JD19098

- Aggregert historisk salg ~200 000 LM/år → `uke_snitt ≈ 3 850 LM`
- `fixed_setup` = 5 002 kr, `unit_cost` = 0.65 kr, `setup_pct` = 30%
- `min_mengde = 5 002 / (0.65 × 0.30) = 25 818 LM` — tilsvarer produksjonslederens anbefaling!

### Skranke i MILP

```
X[p,l,t] ≥ min_batch × Y[p,l,t]
```

Dette tvinger modellen til å produsere **enten 0 eller ≥ min_batch** når den rigger linjen. Overskudd går til lager.

## Kostnadsmodell

### Produksjonsplan (per linje)

| Felt | Forklaring |
|------|------------|
| `quantity` | Produsert mengde LM |
| `actual_batch_size` | Faktisk batch-størrelse (= production når Y=1) |
| `base_unit_cost` | Operasjonskost − biproduktverdi per enhet |
| `setup_per_unit` | `fixed_setup / faktisk_batch` |
| `cost_per_unit` | `base_unit_cost + setup_per_unit` (dynamisk) |
| `holding_cost_per_unit` | `base_unit_cost × hold_pct/100` (per uke) |
| `source` | `demand` (faktisk ordre) eller `bom-avledet` |

### Batch-størrelse vs lagerkostnad

For et produkt med ukentlig salg `S`, batch `B` og lagerkost `h` (%/uke):

```
Uker dekket  T = B / S
Setup/enh    = fixed_setup / B
Hold/enh     = (T × enhetsverdi × h) / 2
Total/enh    = setup/enh + hold/enh
```

**Optimal batch** er der `total/enh` er lavest. Alt som overstiger dette gjør produktet dyrere, ikke billigere.

### Eksempel: JD19148 (ukessalg ~1 200 LM, setup 5 002 kr, verdi 0.70 kr/LM)

| Batch | Uker | Setup/enh | Hold/enh | Total/enh |
|-------|------|-----------|----------|-----------|
| 10 000 | 8 | 0.50 | 0.06 | 0.56 |
| **25 000** | 21 | 0.20 | 0.15 | **0.35** ← optimal |
| **38 411** (modell) | 32 | 0.13 | 0.22 | **0.36** |
| 50 000 | 42 | 0.10 | 0.29 | 0.39 |
| 75 000 (prod.leder) | 63 | 0.07 | 0.44 | **0.51** |

Produksjonslederens 75 000 LM dekker 62 uker med salg — lagerkostnaden spiser opp setup-besparelsen. Modellen velger derfor 38 411, som er 30% billigere per enhet.

## Resultat (JSON)

`resultat.json` inneholder:

- **`status`** — `Optimal` / `Feasible` / annet
- **`total_cost`** — total optimert kost (produksjon + setup + transport + lagerhold)
- **`production_plan`** — hver produksjonslinje med dynamiske kostnader
- **`transport_plan`** — transporterte mengder mellom lokasjoner
- **`batch_decisions`** — hvilke produkt/lokasjon/uke som rigges og setup-kost
- **`inventory_levels`** — lagerbeholdning per produkt/lokasjon/uke
- **`sequence`** — optimal produksjonsrekkefølge per arbeidssenter × uke (med `--sequence`)

## Sekvensering (TSP)

Med `--sequence` beregnes optimal rekkefølge per arbeidssenter × uke:

1. Grupper produkter per familie (FTI-prefiks+siffer).
2. Bruk **greedy nearest-neighbor** til å finne rekkefølgen som minimerer total endringstid.
3. `changeover_minutter` mellom familier beregnes fra `_changeover_minutes()`:
   - Suffiks-bytte (samme familie) → 0 min
   - Bredde-bytte → 22.5 min
   - Tykkelse-bytte → 52.5 min
   - Prefiks-bytte (annen serie) → 90 min

## Analyseverktøy

```
python src/scripts/analyser_batch.py JD19148
```

Viser batch-størrelse vs lagerkostnad-tabell for et produkt — verktøyet for å utfordre/validere produksjonslederens batch-størrelser med matematikk.