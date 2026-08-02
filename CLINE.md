# ProduksjonsKalkyle — Cline-prosjektregler

## Prosjektbeskrivelse

**ProduksjonsKalkyle** er et system for standardkost-beregning og produksjonssimulering for trelastindustrien (Fram Treindustri / Kodal Hovleri). Det beregner materialkost, operasjonskost, setupkost og biproduktverdi, og støtter "what-if"-simuleringer via et Marimo web-grensesnitt.

---

## Arkitektur-oversikt

```
Excel (.xlsx) ── import ──→ SQLite (endringslogg.db) ──→ Python-dataobjekter ──→ Kostnadsberegning
                                   ↑                            ↓
                              Marimo-app (varekost_app.py)    JSON / PDF / Excel-eksport
```

### Filer og deres rolle

## Mappestruktur

```
ProduksjonsKalkyle/
├── src/                    # All Python-kode
│   ├── kostberegning.py    # Kjernelogikk
│   ├── data_repo.py        # SQLite-database
│   ├── excel_bridge.py     # Excel ↔ SQLite
│   ├── generer_pdf_rapport.py
│   ├── generer_excel_rapport.py
│   ├── generer_dokumentasjon.py
│   ├── varekost_app.py     # Marimo web-app
│   ├── app_style.css       # Marimo app-styling (CSS-fil, fra STYLING.md)
│   ├── sheet_i18n.json     # Oversettelser + fanefarger for Excel-eksport
│   ├── __init__.py
│   ├── produksjonskalkyle.db      # SQLite-database (DATA)
│   ├── Produksjonsmodell_Testdata_v3.xlsx  # Testdata-Excel
│   ├── Produksjonsmodell_Mal.xlsx          # Excel-mal
│   └── scripts/            # Hjelpeverktøy
│       ├── lag_testdata_v3.py
│       ├── lag_baseline.py
│       ├── sjekk_diff.py
│       ├── oppdater_mal.py
│       └── sqldatabaseTest.py
├── docs/                   # Dokumentasjon
│   ├── Brukermanual_Produksjonsmodell.md
│   ├── Produksjonsmodell_Dokumentasjon.md
│   ├── STYLING.md
│   └── planer/             # Planleggingsdokumenter
│       ├── ForslagSimulering.md
│       ├── BATCH_SIMULERING_PLAN.md
│       └── VERDIKJEDE_PLAN.md
├── output/                 # Genererte filer
│   ├── baseline_kalkyle.json
│   ├── baseline_simulering.json
│   └── test_log.txt
├── layouts/                # Marimo layouts
├── CLINE.md
├── .gitignore
└── requirements.txt
```

## Filer og deres rolle

| Fil | Formål |
|-----|--------|
| `src/kostberegning.py` | **Kjernelogikk:** ExcelData-parsing, SQLite-datalasting, CostCalculator (kostnadskalkyle), SimulationEngine (what-if), SimulationOverride, export_product_costs_to_json |
| `src/data_repo.py` | SQLite-databasehåndtering: DataRepo-klasse med CRUD, endringslogg, versjonering av opplastede filer. Database: `src/produksjonskalkyle.db` |
| `src/excel_bridge.py` | Import/eksport mellom Excel og SQLite. `validate_excel()` (validering), `import_excel_to_sqlite()` (import), `export_sqlite_to_excel()` (eksport til 11 ark inkl. endringslogg). Støtter engelske og norske ark-navn via `sheet_i18n.json` |
| `src/sheet_i18n.json` | Oversettelser og farger for Excel-faner ved eksport. Norsk navn (`nb`) + tabColor (`color`) per ark. Endres uten å røre Python-kode. `null` = standard Excel-farge |
| `src/varekost_app.py` | **Marimo web-app** med 4 faner: simulering & analyse, dataimport & versjoner, datamodell-innsyn, endringslogg. Benytter `css_file="app_style.css"` for styling og `html_head_file="head.html"` for favicon/app-tittel |
| `src/app_style.css` | Marimo app-styling (CSS) — fargepalett, tabell-hover (grønn `#E6FFFA` i stedet for Marimo sin gule), `.fti-card`, `.fti-header`, KPI-kort m.m. Følger `docs/STYLING.md` |
| `src/generer_pdf_rapport.py` | PDF-rapportgenerering med ReportLab. Inneholder: `generer_rapport()` (simuleringsrapport med ledelsessammendrag, tabeller, kapasitetsdata), `generer_dokumentasjon_pdf()` (Markdown→PDF for dokumentasjonsfiler), `_parse_markdown_to_story()` (Markdown-parser), samt felles styling (farger, fonter, tittelside, header/footer, logo) |
| `src/generer_excel_rapport.py` | Excel-rapportgenerering for simuleringsresultater |
| `src/generer_dokumentasjon.py` | Wrapper-script for å generere stilede PDF-er fra Markdown-dokumentasjon. Støtter enkeltfiler, flere filer og `--all` |
| `src/scripts/lag_testdata_v3.py` | Generer testdata-Excel (`src/Produksjonsmodell_Testdata_v3.xlsx`) |
| `src/scripts/lag_baseline.py` | Beregn baseline-kalkyle (konsoll/JSON til `output/`) |
| `src/scripts/sjekk_diff.py` | Sjekk diff mellom to kjøringer |
| `src/scripts/oppdater_mal.py` | Oppdater Excel-mal (`src/Produksjonsmodell_Mal.xlsx`) |
| `src/scripts/sqldatabaseTest.py` | Test-verktøy for SQLite-databasen |
| `src/scripts/test_transport.py` | Test av transportvare-modulen (5 tester, in-memory DB — berører aldri produksjonsdatabasen) |
| `src/scripts/sjekk_i18n_eksport.py` | Verifiser norske fane-navn og tabColor-farger i Excel-eksporten (fra `sheet_i18n.json`) |
| `output/baseline_kalkyle.json` | Lagret baseline-kalkyle |
| `output/baseline_simulering.json` | Lagret simuleringsresultat |
| `requirements.txt` | Avhengigheter: marimo>=0.23.0, openpyxl>=3.1.0, pandas>=2.0.0, reportlab>=4.0.0 |
| `docs/STYLING.md` | Fargepalett, typografi, CSS-klasser for Marimo-app, PDF og Excel |
| `docs/Brukermanual_Produksjonsmodell.md` | Sluttbrukermanual for Excel-arket og Marimo web-app (for produksjonsledere, økonomi, innkjøp) |
| `docs/Brukermanual_Produksjonsmodell.pdf` | Generert PDF fra Brukermanual_Produksjonsmodell.md — stylet med Fram Treindustri-profil |
| `docs/Produksjonsmodell_Dokumentasjon.md` | Detaljert teknisk dokumentasjon av datamodellen, beregningslogikk, Marimo web-app og import/eksport |
| `docs/Produksjonsmodell_Dokumentasjon.pdf` | Generert PDF fra Produksjonsmodell_Dokumentasjon.md — stylet med Fram Treindustri-profil |
| `docs/planer/` | Planleggingsdokumenter (ForslagSimulering, BATCH_SIMULERING_PLAN, VERDIKJEDE_PLAN) |
| `layouts/varekost_app.grid.json` | Marimo grid-layout |

---

## Dataflyt

1. **Excel import:** Bruker laster opp Excel-fil i Marimo-appen → `validate_excel()` sjekker ark/kolonner/kryssreferanser → `import_excel_to_sqlite()` lagrer i SQLite med endringslogg
2. **SQLite → Python:** `SqliteData`-klassen leser SQLite-tabeller og bygger strukturerte lister (`Product`, `WorkCenter`, `BOMLine`, `RoutingLine`, etc.)
3. **Python-beregning:** `CostCalculator` → Materialkost → Operasjonskost → Setupkost → Brutto → Biproduktverdi → Netto
4. **Simulering:** `SimulationOverride` (overstyrte parametere) → `SimulationEngine.compare_all()` → `SimulationComparison` (original vs simulert)
5. **Eksport:** PDF (ledelsessammendrag + detaljer), Excel (11 ark), JSON

---

## Datamodellen (Excel-ark / SQLite-tabeller)

| # | Ark / Tabell | Innhold | Nøkkelkolonner |
|---|-------------|---------|----------------|
| 1 | **Product Master** / `products` | Vareregister (råvarer, ferdigvarer, biprodukter) | Item No, Description, Item Type, Product Group, Base Unit of Measure |
| 2 | **Locations** / `locations` | Fabrikker og lagre | Location Code, Name, Location Type |
| 3 | **Work Centers** / `work_centers` | Arbeidssentre med kostsatser | Code, Labor/Machine/Overhead Cost per Hour, Capacity Hours/Day, Effective Capacity % |
| 4 | **Operation Master** / `operations` | Standardoperasjoner | Code, Description, Default Work Center |
| 5 | **Item Costs** / `item_costs` | Kostpriser per vare | Item No, Cost Type, Unit Cost, Currency, Effective Date |
| 6 | **BOM** / `bom_lines` | Stykkliste | Parent Item No, Component Item No, Quantity Per, Scrap %, Co-Prod %, Co-Prod Item No |
| 7 | **Routing** / `routing_lines` | Produksjonsflyt (operasjoner, tider) | Item No, Operation No, Work Center Code, Setup Time, Run Time, Batch Size, Changeover Time |
| 8 | **By Product Rules** / `byproduct_rules` | Biprodukter og verdsetting | Parent Item No, By Product Item No, Expected Quantity, Market Value, Allocation Method |
| 9 | **Capacity Calendar** / `capacity_days` | Kapasitetskalender per arbeidssenter | Work Center, Date, Available Hours, Planned Downtime |
| 10 | **Production Scenario** / `production_scenarios` | Forhåndsdefinerte scenarioer | Scenario Name, Product, Planned Quantity |

---

## Kjerneberegninger (i `CostCalculator`)

### Materialkost
```
Materialkost = (Unit Cost / Quantity Per) × (1 + Scrap% / 100)
```
- **Dynamisk cost roll-up:** Hvis en BOM-komponent er en ferdigvare/halvfabrikat som allerede er beregnet, brukes dens dynamiske brutto produksjonskost (fra `calculated_costs`-cachen) i stedet for statisk Item Cost
- **Quantity Per** = output-enheter per input-enhet (f.eks. 400 LM per M3)

### Operasjonskost
```
Operasjonskost = (Run Time Minutes / 60) × Timekost (justert for effektiv kapasitet)
```
- **Effektiv Timepris:** Nominell timepris / (Effective Capacity % / 100)
  - F.eks. 1600 kr/t / 0,85 = 1882,35 kr/t (ved 85% effektivitet)
- **Co-Prod justering:** Run time økes marginalt med (1 + co_pct) for å kompensere for co-produkt

### Setupkost
```
Setupkost = ((Setup Time Minutes / 60) × Timekost) / Batch Size
```
- Setupkost fordeles på batch-størrelse (større batch = lavere kost per enhet)
- Co-Prod justerer IKKE setup-tiden (omstilling er uavhengig av kvalitetsfordeling)

### Biproduktverdi
```
Biproduktverdi = Sum(Expected Quantity × Market Value)
```
- Trekkes fra brutto kostnad for å få netto kostnad

### Co-Produkt (A/B-vare modellering)
- Hvis BOM har `co_product_pct > 0` og `co_product_item_no`:
  - A-vare får justert run-time og materialkost basert på reelt utbytte (quantity_per / (1 - co_pct))
  - B-vare får egen kalkyle med allokert materialkost (samme per LM) og proporsjonal operasjonskost
  - B-varen opereres via `_calc_co_product_results()`

---

## Viktige regler og konvensjoner

### ⚠️ ALDRI endre `kostberegning.py` uten eksplisitt godkjenning
Inneholder all simuleringslogikk, ExcelData-parsing, CostCalculator, SimulationEngine, SimulationOverride — kritisk kjernelogikk.

### Marimo-celleorganisering
```
imports → CSS/header → datalasting → filter → baseline → justeringsparametere → simulering → resultater → eksport
```

### Navnekonvensjoner
- `mo.ui.data_editor`-variabler: `rm_price_df`, `bom_scrap_df`, `wc_cost_df`, `routing_df`
- Filtrerte datasett: `filtered_products`, `filtered_bom_lines`, `filtered_routing_lines` osv.
- Knapper: `run_button`, `export_pdf_button`, `export_excel_button`
- Interne variabler: `_tmp`, `_rows`, `_df` (med underscore-prefiks)

### Concat-konvensjon
```
"Beskrivelse": f"{_p.description} · {_p.base_uom}"  # → "Gran 50x200 US/V · LM"
```
- Separator: ` · ` (mellomrom + midtdot + mellomrom)
- Ved tilbakelesing: `.split(" · ")[0]`

### Filter-logikk (kaskade)
```
vareFilter → filtered_products → filtered_bom_lines → filtered_work_centers
                               → filtered_routing_lines → filtered_operations
                               → filtered_byproduct_rules
```

### Marimo-fallgruver
- `mo.ui.data_editor` = redigerbare tabeller, `mo.ui.table` = skrivebeskyttet
- UI-elementer må stå **sist i cellen** (uten tilordning) for å vises
- `mo.output.replace()` for output etter knappetrykk (ikke `return`)
- `mo.download()` (uten `ui.`) — `mo.ui.download()` finnes ikke i 0.23.x
- `editable_columns` støttes, men `label` og `column_widths` støttes ikke i 0.23.x

### Override-logikk (simulering)
- `persistent_overrides` = dict i Marimo med nøkler: `item_costs`, `bom_scrap`, `bom_co_product`, `work_centers`, `routing`
- Overstyrte verdier lagres i `SimulationOverride`-objektet
- Fjernes automatisk når bruker setter verdi tilbake til original (sjekk: `abs(ny - org) > 0.001`)

### SQLite-database
- DataRepo = context manager (`with DataRepo() as db:`)
- `src/produksjonskalkyle.db` (ved siden av `data_repo.py`)
- Endringer loggføres felt-for-felt i `change_log`-tabellen
- Excel-filer lagres som BLOB i `uploaded_files`-tabellen (versjonering)
- `clear_all_data()` sletter stamdata, **bevarer** change_log

### Style-guide
- Se `STYLING.md` for full fargepalett (skogsgrønn #14532D, lysegrønn #2F855A, etc.)
- CSS-klasser: `.fti-card`, `.fti-highlight-green`, `.fti-highlight-red`
- PDF: ReportLab med A4, DejaVu-fonter (DejaVu/Helvetica/Bold)
- Excel: openpyxl med Calibri, grønn header (#14532D), grønn border (#48BB78)

---

## Kommandolinje-verktøy

```bash
# Full produktkalkyle (fra src/)
python src/kostberegning.py --excel src/Produksjonsmodell_Testdata_v3.xlsx

# Ett spesifikt produkt
python src/kostberegning.py --excel src/Produksjonsmodell_Testdata_v3.xlsx --product FG001

# Innebygget testdata (uten Excel)
python src/kostberegning.py --test

# Generer testdata-Excel (fra src/scripts/)
python src/scripts/lag_testdata_v3.py

# SQLite-administrasjon
python src/data_repo.py --stats      # Vis tabell-statistikk
python src/data_repo.py --changes    # Vis endringslogg
python src/data_repo.py --clear      # Tøm data (bevar logg)

# Excel ↔ SQLite
python src/excel_bridge.py --import data.xlsx
python src/excel_bridge.py --export utdata.xlsx

# Baseline-verktøy
python src/scripts/lag_baseline.py              # Generer baseline
python src/scripts/lag_baseline.py --sammenlign # Sammenlign mot baseline
python src/scripts/sjekk_diff.py                # Sjekk diff

# Transportvarer (fler-høvleri-produksjon)
python src/scripts/test_transport.py            # Kjør 5 tester (in-memory DB)

# Test-modus for Marimo-appen (in-memory DB med testdata, berører ALDRI produksjonsdata)
# PowerShell
$env:PRODUKSJONSKALKYLE_TEST="true"; marimo run src/varekost_app.py
# Bash (Git Bash / WSL)
PRODUKSJONSKALKYLE_TEST=true marimo run src/varekost_app.py
# CMD
set PRODUKSJONSKALKYLE_TEST=true && marimo run src/varekost_app.py
# Normal modus (ekte database)
marimo run src/varekost_app.py

# Start Marimo-app
marimo run src/varekost_app.py

# Generer stilede PDF-er fra Markdown-dokumentasjon
python src/generer_dokumentasjon.py docs/fil.md                     # Én fil
python src/generer_dokumentasjon.py docs/fil1.md docs/fil2.md       # Flere filer
python src/generer_dokumentasjon.py --all                           # Alle dokumentasjonsfiler
python src/generer_dokumentasjon.py --all --output ./rapporter      # Til egen mappe
```

## Marimo-dokumentasjon

- [Marimo API-dokumentasjon (LLM-vennlig)](https://docs.marimo.io/llms.txt) — `llms.txt`-fil med Marimos API, komponenter og beste praksis for AI-assistenter
- [Marimo offisiell dokumentasjon](https://docs.marimo.io)
