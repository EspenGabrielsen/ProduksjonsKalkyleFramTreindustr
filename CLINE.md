# ProduksjonsKalkyle — Cline-prosjektregler

## Marimo-dokumentasjon
Før du jobber med Marimo-relaterte oppgaver i dette prosjektet, hent nyeste dokumentasjon:
https://docs.marimo.io/llms.txt

Se også styling-guide i `STYLING.md` for prosjektets fargepalett, typografi og visuelle retningslinjer.

## Kostberegningslogikk (`kostberegning.py`)
**⚠️ Aldri endre `kostberegning.py` uten eksplisitt godkjenning fra bruker.**
Denne filen inneholder all simuleringslogikk, ExcelData-parsing, CostCalculator, SimulationEngine, SimulationOverride — kritisk kjernelogikk som må håndteres med forsiktighet.

## Prosjektkonvensjoner

### Celle-organisering (Marimo)
Rekkefølge: imports → tittel/header → datalasting → filter → visning av originale data → justeringsparametere → simulering → resultater → eksport (PDF/Excel)

### Navnekonvensjoner
- `mo.ui.data_editor`-variabler: `rm_price_df`, `bom_scrap_df`, `wc_cost_df`, `routing_df`
- Filtrerte datasett: `filtered_products`, `filtered_bom_lines`, `filtered_routing_lines` osv.
- Knapper: `run_button`, `export_pdf_button`, `export_excel_button`
- Interne/private variabler: `_tmp`, `_rows`, `_df` (med underscore-prefiks for å unngå eksport til andre celler)

### Concat-konvensjon for berikede kolonner
For å unngå for mange kolonner i tabeller, berikes eksisterende kolonner med ekstra info:
```python
"Beskrivelse": f"{_p.description} · {_p.base_uom}"  # → "Gran 50x200 US/V · LM"
```
- Separator: ` · ` (mellomrom + midtdot + mellomrom)
- Original verdi bevares **før** ` · `
- Ved lesing tilbake (hvis kolonnen brukes som nøkkel): `.split(" · ")[0]`

### Filter-logikk (kaskade)
```
vareFilter → filtered_products → filtered_bom_lines → filtered_work_centers
                               → filtered_routing_lines → filtered_operations
                               → filtered_byproduct_rules
```

### Marimo-fallgruver (husk)
- `mo.ui.data_editor` for redigerbare tabeller, `mo.ui.table` for skrivebeskyttet
- UI-elementer må stå **sist i cellen** (uten tilordning) for å vises
- `mo.output.replace()` for output som vises etter knappetrykk (ikke `return`)
- `mo.download()` (uten `ui.`) — `mo.ui.download()` finnes ikke i 0.23.x
- Unngå dobbel-rendering: opprett og vis UI-elementer i samme celle