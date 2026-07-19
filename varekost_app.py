import marimo

__generated_with = "0.23.14"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo
    import sys
    import io
    import os
    import re
    import tempfile
    import shutil
    from datetime import datetime, date
    from pathlib import Path
    from typing import Optional

    import pandas as pd

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    elif hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    from kostberegning import (
        ExcelData,
        SqliteData,
        CostCalculator,
        SimulationEngine,
        SimulationOverride,
        SimulationComparison,
        ProductCostResult,
        MaterialCostDetail,
        OperationCostDetail,
        ByProductDetail,
        export_product_costs_to_json,
    )
    from generer_pdf_rapport import generer_rapport, registrer_fonter, _hent_logo
    from generer_excel_rapport import generer_excel_rapport
    from data_repo import DataRepo
    from excel_bridge import import_excel_to_sqlite, export_sqlite_to_excel, validate_excel

    return (
        CostCalculator,
        ExcelData,
        SqliteData,
        SimulationEngine,
        SimulationOverride,
        generer_rapport,
        mo,
        os,
        pd,
        registrer_fonter,
        tempfile,
        _hent_logo,
        generer_excel_rapport,
        DataRepo,
        import_excel_to_sqlite,
        export_sqlite_to_excel,
        validate_excel,
    )


# ═══ KAP 1: CSS + Header ════════════════════════════════════════════════


@app.cell
def _(mo):
    mo.Html("""
    <style>
        body { 
            background-color: #F3F5F2; 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #2C3E2B;
        }
        .marimo-app { max-width: 1200px; margin: 40px auto; padding: 0 24px; }
        
        /* ===== FTI Header ===== */
        .fti-header {
            background: linear-gradient(135deg, #14532D 0%, #1B6E3D 100%);
            border-radius: 12px;
            padding: 20px 28px;
            margin-bottom: 20px;
            box-shadow: 0 4px 16px rgba(20, 83, 45, 0.15);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .fti-header-title {
            font-size: 1.8em;
            font-weight: 700;
            color: #FFFFFF;
            line-height: 1.2;
        }
        .fti-header-subtitle {
            font-size: 0.95em;
            color: #C6E6D0;
            line-height: 1.2;
        }
        
        /* ===== FTI Cards ===== */
        .fti-card {
            background: #F9FBF8; border-radius: 12px; padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(27, 89, 43, 0.08);
            border-left: 6px solid #48BB78; 
            transition: box-shadow 0.2s ease;
        }
        .fti-card:hover {
            box-shadow: 0 6px 20px rgba(27, 89, 43, 0.12);
        }
        .fti-card h2, .fti-card h3 { color: #14532D; margin-top: 0; margin-bottom: 12px; font-weight: 600; }
        .fti-highlight-green { color: #2F855A; background-color: #E6FFFA; padding: 2px 6px; border-radius: 4px; font-weight: 600; }
        .fti-highlight-red { color: #C53030; background-color: #FFF5F5; padding: 2px 6px; border-radius: 4px; font-weight: 600; }
        
        /* ===== FTI Footer ===== */
        .fti-footer {
            background: #14532D;
            color: #C6E6D0;
            font-size: 0.85em;
            text-align: center;
            padding: 12px 24px;
            border-radius: 8px;
            margin-top: 32px;
            margin-bottom: 16px;
        }
        
        /* ===== FTI Sidebar ===== */
        .fti-sidebar-logo {
            padding: 8px 0;
            border-bottom: 1px solid rgba(255,255,255,0.15);
            margin-bottom: 16px;
        }
        .fti-sidebar-section {
            margin-bottom: 20px;
        }
        .fti-sidebar-section h3 {
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: #48BB78;
            margin: 0 0 8px 0;
            font-weight: 600;
        }
        
        /* ===== Styling av Marimo radio-knapper ===== */
        .marimo-radio {
            background: #F9FBF8;
            border-radius: 8px;
            padding: 4px;
            border: 1px solid #D1E0D4;
        }
        .marimo-radio label {
            padding: 8px 14px;
            border-radius: 6px;
            font-weight: 500;
            color: #2C3E2B;
            transition: all 0.15s ease;
        }
        .marimo-radio input[type="radio"]:checked + label {
            background: #14532D;
            color: white;
        }
        .marimo-radio label:hover {
            background: #E6FFFA;
        }
        
        /* ===== Styling av dropdowns ===== */
        .marimo-dropdown {
            border-radius: 6px;
            border: 1px solid #D1E0D4;
            background: #F9FBF8;
        }
        
        /* ===== Outline/TOC styling ===== */
        .marimo-outline {
            background: transparent;
        }
        .marimo-outline a {
            color: #C6E6D0;
            text-decoration: none;
            font-size: 0.9em;
            padding: 4px 0;
            display: block;
            transition: color 0.15s ease;
        }
        .marimo-outline a:hover {
            color: #FFFFFF;
        }
        
        /* ===== Sidebar-specifikke overrides ===== */
        [data-marimo-theme="dark"] .fti-header {
            background: linear-gradient(135deg, #0B2819 0%, #14532D 100%);
        }
        [data-marimo-theme="dark"] .fti-footer {
            background: #0B2819;
        }
    </style>
    """)
    return


@app.cell
def _(mo):
    mo.Html("""
    <div class="fti-header" id="fti-header">
        <div style="display: flex; align-items: center; gap: 16px;">
            <img src="https://framtreindustri.no/wp-content/uploads/2025/08/logo-liggende-2048x512.png"
                 alt="Fram Treindustri" style="height: 48px;" />
            <div>
                <div class="fti-header-title">Produksjonskost Simulator</div>
                <div class="fti-header-subtitle">Fram Treindustri — Standardkostkalkyle, simulering og analyse</div>
            </div>
        </div>
        <div style="text-align: right; color: #C6E6D0; font-size: 0.85em;">
            Last opp Excel-modell · Juster parametere · Se kostnader i sanntid
        </div>
    </div>
    """)
    return


# ═══ 📂 Datagrunnlag ═══════════════════════════════════════════════════


@app.cell
def _(mo):
    mo.md("## 📂 Datagrunnlag")
    return


@app.cell
def _(mo):
    overrides = {
        "item_costs": {},
        "bom_scrap": {},
        "bom_co_product": {},
        "work_centers": {},
        "routing": {},
    }
    # get_reload / set_reload = Marimo state (trigger for reload)
    # Når set_reload kalles i import-cellen, re-kjøres alle celler
    # som leser get_reload() - inkludert data-loading-cellen.
    get_reload, set_reload = mo.state(0)
    return (overrides, get_reload, set_reload)


@app.cell
def _(CostCalculator, DataRepo, SqliteData, SimulationEngine, get_reload, mo, os, overrides, pd, tempfile):
    _db = DataRepo()
    _db.initialize()

    data = None
    engine = None
    baseline = None
    db_stats = None

    # Les trigger for a skape Marimo-avhengighet.
    # Nar set_reload kalles i import-cellen, re-kjorer denne cellen
    # og SqliteData() leser ferske data fra SQLite.
    _reload_verdi = get_reload()

    if _db.is_empty():
        mo.output.replace(
            mo.md("""
            ### Velkommen!

            Du må laste opp en Excel-fil med data for å komme i gang.

            **Slik fungerer det:**
            1. Last ned Excel-malen fra SharePoint
            2. Rediger data i Excel
            3. Last opp filen nedenfor — valideres og importeres automatisk
            4. Kjør simuleringer og analyser kostnader
            5. Last ned oppdatert Excel ved behov

            ---
            """)
        )
    else:
        try:
            data = SqliteData()
            engine = SimulationEngine(data)
            _calculator = CostCalculator(data)
            baseline = _calculator.calculate_all()

            db_stats = _db.stats

            _n_fg = len([p for p in data.products if p.item_type in ('Finished Good', 'Semi Finished')])
            _n_rm = len([p for p in data.products if p.item_type == 'Raw Material'])
            _n_bp = len([p for p in data.products if p.item_type == 'By Product'])

            mo.output.replace(
                mo.md(
                    f"""
                    ### ✅ Data lastet!

                    | Data | Antall |
                    |---|---|
                    | Produkter | {len(data.products)} |
                    | Ferdigvarer/Halvfabrikata | {_n_fg} |
                    | Råvarer | {_n_rm} |
                    | Biprodukter | {_n_bp} |
                    | Arbeidssentre | {len(data.work_centers)} |
                    | BOM-linjer | {len(data.bom_lines)} |
                    | Routing-linjer | {len(data.routing_lines)} |
                    | Biprodukt-regler | {len(data.byproduct_rules)} |
                    | Scenarioer | {len(data.scenarios)} |
                    | Endringer logget | {db_stats.get('change_log', 0)} |
                    """
                )
            )
        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil ved lasting: {_e}"))
            import traceback as _traceback
            _traceback.print_exc()

    return baseline, data, db_stats


# ═══ KAP 3: 📤 Last opp ny Excel-fil ════════════════════════════════════


@app.cell
def _(mo):
    mo.md("### 📤 Last opp Excel")
    return


@app.cell
def _(DataRepo, import_excel_to_sqlite, mo, os, tempfile, validate_excel):
    excel_import_file = mo.ui.file(
        label="📄 Velg Excel-fil",
        filetypes=[".xlsx"],
        multiple=False,
    )
    excel_import_kommentar = mo.ui.text(
        label="Hva er endret? (valgfri kommentar)",
        placeholder="f.eks. 'Oppdaterte råvarepriser Q3' eller 'Ny BOM for Panel'",
    )

    mo.vstack([
        mo.md("Last opp en redigert Excel-fil. Den valideres og importeres automatisk."),
        excel_import_file,
        excel_import_kommentar,
    ])
    return excel_import_file, excel_import_kommentar


@app.cell
def _(DataRepo, excel_import_file, excel_import_kommentar, import_excel_to_sqlite, mo, os, overrides, set_reload, tempfile, validate_excel):
    if excel_import_file.value and excel_import_kommentar.value.strip():
        try:
            _upload = excel_import_file.value[0]
            _fcontent = _upload.contents
            _comment = excel_import_kommentar.value.strip()

            _temp_dir = tempfile.mkdtemp()
            _temp_path = os.path.join(_temp_dir, _upload.name)
            with open(_temp_path, "wb") as _f:
                _f.write(_fcontent)

            _validation = validate_excel(_temp_path)
            _outputs = []

            _outputs.append(mo.md(f"### Kontroll av '{_upload.name}'"))

            if _validation["stats"]:
                _stat_rows = []
                for _sheet, _count in _validation["stats"].items():
                    _stat_rows.append(f"  - {_sheet}: {_count} rader")
                _outputs.append(mo.md("**Antall rader:**\n" + "\n".join(_stat_rows)))

            if _validation["diff"]["nye_rader"] > 0:
                _outputs.append(mo.md(f"**Nye produkter:** {_validation['diff']['nye_rader']}"))
            if _validation["diff"]["slettede_rader"] > 0:
                _outputs.append(mo.md(f"**Fjernede produkter:** {_validation['diff']['slettede_rader']}"))

            if not _validation["valid"]:
                _feil_txt = "\n".join(f"  - {e}" for e in _validation["errors"])
                _outputs.append(mo.md(f"### ❌ Kontrollen fant feil! Ingenting er importert.\n\n{_feil_txt}"))
            else:
                if _validation["warnings"]:
                    _adv_txt = "\n".join(f"  - {w}" for w in _validation["warnings"])
                    _outputs.append(mo.md(f"**Advarsler:**\n{_adv_txt}"))

                _stats = import_excel_to_sqlite(_temp_path, excel_blob=_fcontent, comment=_comment)
                _outputs.append(mo.md("### ✅ Import fullført!"))
                for _sheet, _count in _stats["tables_updated"].items():
                    _outputs.append(mo.md(f"  - {_sheet}: oppdatert {_count} rader"))
                if _stats["errors"]:
                    _err_txt = "\n".join(f"  - {e}" for e in _stats["errors"])
                    _outputs.append(mo.md(f"**Import-advarsler:**\n{_err_txt}"))

                # ═══════════════════════════════════════════════════
                # Nullstill ALLE overstyringer.
                # ═══════════════════════════════════════════════════
                overrides["item_costs"].clear()
                overrides["bom_scrap"].clear()
                overrides["bom_co_product"].clear()
                overrides["work_centers"].clear()
                overrides["routing"].clear()

                # ═══════════════════════════════════════════════════
                # Trigger reload: data-loading-cellen re-kjores og
                # SqliteData() leser ferske data fra SQLite.
                # ═══════════════════════════════════════════════════
                set_reload(lambda v: v + 1)

            mo.output.replace(mo.vstack(_outputs))

        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil: {_e}"))
            import traceback as _traceback
            _traceback.print_exc()


# ═══ KAP 4: 📜 Gå tilbake til tidligere versjon ════════════════════════


@app.cell
def _(mo):
    mo.md("### 📜 Versjonshistorikk")
    return


@app.cell
def _(DataRepo, mo):
    _db = DataRepo()
    _db.initialize()
    _uploads = _db.get_last_uploads(limit=10)

    historikk_valg = None
    upload_id_map = {}
    if len(_uploads) > 0:
        _options = {}
        for _u in _uploads:
            _ts = _u['uploaded_at'][:16]
            _fn = _u['filename']
            _cm = _u.get("comment", "").strip()
            if _cm:
                _label = f"{_ts} – {_fn} ({_cm})"
            else:
                _label = f"{_ts} – {_fn}"
            _options[_label] = _label
            upload_id_map[_label] = _u["id"]

        historikk_valg = mo.ui.dropdown(
            options=_options,
            label="📜 Velg tidligere versjon",
            value=None,
        )
        mo.output.replace(
            mo.vstack([
                mo.md("Velg en tidligere import for å laste den inn på nytt."),
                historikk_valg,
            ])
        )
    return (historikk_valg, upload_id_map)


@app.cell
def _(DataRepo, historikk_valg, import_excel_to_sqlite, mo, os, overrides, set_reload, tempfile, upload_id_map):
    if historikk_valg is not None and historikk_valg.value:
        try:
            _label = historikk_valg.value
            _upload_id = upload_id_map.get(_label)

            if _upload_id is None:
                mo.output.replace(mo.md(f"### ❌ Kunne ikke finne tidligere versjon."))
            else:
                _db = DataRepo()
                _db.initialize()

                _blob = _db.get_upload_blob(_upload_id)
                if not _blob:
                    mo.output.replace(mo.md("### ❌ Kunne ikke hente filen fra database."))
                else:
                    _uploads = _db.get_last_uploads(limit=10)
                    _filename = "historisk_import.xlsx"
                    for _u in _uploads:
                        if _u["id"] == _upload_id:
                            _filename = _u["filename"]
                            break

                    _db.clear_all_data()

                    _temp_dir = tempfile.mkdtemp()
                    _temp_path = os.path.join(_temp_dir, _filename)
                    with open(_temp_path, "wb") as _f:
                        _f.write(_blob)

                    _stats = import_excel_to_sqlite(
                        _temp_path, excel_blob=_blob,
                        comment=f"Gjeninnlasting av versjon {_upload_id}",
                    )

                    # Nullstill overstyringer + trigger reload
                    overrides["item_costs"].clear()
                    overrides["bom_scrap"].clear()
                    overrides["bom_co_product"].clear()
                    overrides["work_centers"].clear()
                    overrides["routing"].clear()
                    set_reload(lambda v: v + 1)

                    mo.output.replace(mo.md("### ✅ Forrige versjon er lastet inn!"))
                    for _sheet, _count in _stats["tables_updated"].items():
                        mo.output.append(mo.md(f"  - {_sheet}: {_count} rader"))

        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil: {_e}"))
            import traceback as _traceback
            _traceback.print_exc()


# ═══ KAP 5: 📥 Last ned komplett datafil ════════════════════════════════


@app.cell
def _(mo):
    mo.md("### 📥 Eksporter data")
    return


@app.cell
def _(mo):
    export_excel_db_button = mo.ui.run_button(label="📥 Last ned komplett datafil")
    export_excel_db_button
    return (export_excel_db_button,)


@app.cell
def _(export_excel_db_button, export_sqlite_to_excel, mo, os, tempfile):
    if export_excel_db_button.value:
        try:
            _output_path = os.path.join(tempfile.gettempdir(), "produksjonsmodell_eksport.xlsx")
            export_sqlite_to_excel(_output_path)

            with open(_output_path, "rb") as _f:
                _excel_content = _f.read()

            mo.output.replace(
                mo.vstack([
                    mo.md("### ✅ Datafil klar!"),
                    mo.download(
                        label="📥 Last ned (11 ark, inkl. endringslogg)",
                        filename="produksjonsmodell_eksport.xlsx",
                        data=_excel_content,
                    ),
                ])
            )
        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil: {_e}"))
    return


# ═══ KAP 6: 🔍 Filtrering (kaskade) ════════════════════════════════════


@app.cell
def _(data, vareFilter):
    filtered_products = []
    filtered_bom_lines = []
    filtered_routing_lines = []
    filtered_byproduct_rules = []
    filtered_item_costs = []
    filtered_scenarios = []
    filtered_work_centers = []
    filtered_locations = []
    filtered_operations = []
    filtered_capacity_days = []
    filter_active = False

    if data:
        _filter_text = vareFilter.value.strip().lower() if vareFilter.value else ""

        if _filter_text:
            filter_active = True

            _fg_set = set()
            for _p in data.products:
                if _p.item_type in ('Finished Good', 'Semi Finished'):
                    if _filter_text in _p.item_no.lower() or _filter_text in _p.description.lower():
                        _fg_set.add(_p.item_no)

            _component_set = set()
            for _bl in data.bom_lines:
                if _bl.parent_item_no in _fg_set:
                    filtered_bom_lines.append(_bl)
                    _component_set.add(_bl.component_item_no)

            _all_product_set = set(_fg_set)

            for _p in data.products:
                if _p.item_type == 'Raw Material' and _p.item_no in _component_set:
                    _all_product_set.add(_p.item_no)

            _bp_set = set()
            for _br in data.byproduct_rules:
                if _br.parent_item_no in _fg_set:
                    _bp_set.add(_br.by_product_item_no)
            for _p in data.products:
                if _p.item_type == 'By Product' and _p.item_no in _bp_set:
                    _all_product_set.add(_p.item_no)

            for _p in data.products:
                if _p.item_no in _all_product_set:
                    filtered_products.append(_p)

            for _br in data.byproduct_rules:
                if _br.parent_item_no in _fg_set:
                    filtered_byproduct_rules.append(_br)

            _routing_wc_set = set()
            for _rl in data.routing_lines:
                if _rl.item_no in _fg_set:
                    filtered_routing_lines.append(_rl)
                    _routing_wc_set.add(_rl.work_center_code)

            for _ic in data.item_costs:
                if _ic.item_no in _all_product_set:
                    filtered_item_costs.append(_ic)

            for _sc in data.scenarios:
                if _sc.product in _fg_set:
                    filtered_scenarios.append(_sc)

            for _wc in data.work_centers:
                if _wc.code in _routing_wc_set:
                    filtered_work_centers.append(_wc)

            _loc_set = {_wc.location_code for _wc in filtered_work_centers}
            for _loc in data.locations:
                if _loc.code in _loc_set:
                    filtered_locations.append(_loc)

            _op_set = {_rl.operation_code for _rl in filtered_routing_lines}
            for _op in data.operations:
                if _op.code in _op_set:
                    filtered_operations.append(_op)

            for _cd in data.capacity_days:
                if _cd.work_center in _routing_wc_set:
                    filtered_capacity_days.append(_cd)
        else:
            filtered_products = list(data.products)
            filtered_bom_lines = list(data.bom_lines)
            filtered_routing_lines = list(data.routing_lines)
            filtered_byproduct_rules = list(data.byproduct_rules)
            filtered_item_costs = list(data.item_costs)
            filtered_scenarios = list(data.scenarios)
            filtered_work_centers = list(data.work_centers)
            filtered_locations = list(data.locations)
            filtered_operations = list(data.operations)
            filtered_capacity_days = list(data.capacity_days)
    return (
        filter_active,
        filtered_bom_lines,
        filtered_byproduct_rules,
        filtered_capacity_days,
        filtered_item_costs,
        filtered_locations,
        filtered_operations,
        filtered_products,
        filtered_routing_lines,
        filtered_scenarios,
        filtered_work_centers,
    )


# ═══ 📊 Kalkyle & Simulering ═══════════════════════════════════════════


@app.cell
def _(mo):
    mo.md("## 📊 Kalkyle & Simulering")
    return


@app.cell
def _(mo):
    mo.md("### 📊 Baseline kostnader")
    return


@app.cell
def _(baseline, mo, pd):
    if baseline:
        _rows = []
        for _r in baseline:
            _rows.append({
                "Lokasjon": _r.location_code,
                "Produkt": _r.product_no,
                "Beskrivelse": _r.product_desc,
                "Materialkost": round(_r.material_cost, 2),
                "Operasjonskost": round(_r.operation_cost, 2),
                "Setupkost": round(_r.setup_cost, 2),
                "Brutto kost": round(_r.gross_production_cost, 2),
                "Biproduktverdi": round(_r.by_product_value, 2),
                "Netto kost": round(_r.net_production_cost, 2),
            })
        _df = pd.DataFrame(_rows)
        mo.output.replace(mo.ui.table(_df, selection=None))
    return


@app.cell
def _(mo):
    mo.md("### 🔧 Parametere")
    return


@app.cell(hide_code=True)
def _(mo):
    mo.md("**🪵 Råvarer**")
    return


@app.cell
def _(data, filtered_products, mo, overrides, pd):
    rm_price_df = None
    if data:
        _rm_items = [p for p in filtered_products if p.item_type == 'Raw Material']
        _rows = []
        for _p in _rm_items:
            _cost = next((c.unit_cost for c in data.item_costs if c.item_no == _p.item_no), 0)
            _ny_cost = overrides["item_costs"].get(_p.item_no, _cost)
            _rows.append({
                "Varenr": _p.item_no,
                "Beskrivelse": f"{_p.description} · {_p.base_uom}",
                "Org. pris": _cost,
                "Ny pris": _ny_cost,
            })
        if _rows:
            _df = pd.DataFrame(_rows)
            rm_price_df = mo.ui.data_editor(_df, editable_columns=["Ny pris"])

    rm_price_df
    return (rm_price_df,)


@app.cell
def _(overrides, rm_price_df):
    if rm_price_df is not None and rm_price_df.value is not None:
        _df = rm_price_df.value
        for _, _row in _df.iterrows():
            _varenr = _row["Varenr"]
            _org = _row["Org. pris"]
            _ny = _row["Ny pris"]
            if abs(_ny - _org) > 0.001:
                overrides["item_costs"][_varenr] = _ny
            elif _varenr in overrides["item_costs"]:
                del overrides["item_costs"][_varenr]
    return


@app.cell
def _(mo):
    mo.md("**🗑️ Svinn- og kapp-prosenter**")
    return


@app.cell
def _(data, filtered_bom_lines, mo, overrides, pd):
    bom_scrap_df = None
    if data:
        _rows = []
        for _bl in filtered_bom_lines:
            _key = (_bl.parent_item_no, _bl.component_item_no)
            _ny_svinn = overrides["bom_scrap"].get(_key, _bl.scrap_pct)
            _ny_co = overrides["bom_co_product"].get(_key, _bl.co_product_pct)
            _komp_desc = next((p.description for p in data.products if p.item_no == _bl.component_item_no), _bl.component_item_no)
            _prod_desc = next((p.description for p in data.products if p.item_no == _bl.parent_item_no), _bl.parent_item_no)
            _rows.append({
                "Komponent": f"{_bl.component_item_no} · {_komp_desc}",
                "Produkt": f"{_bl.parent_item_no} · {_prod_desc}",
                "Org. svinn %": _bl.scrap_pct,
                "Nytt svinn %": _ny_svinn,
                "Org. co-prod %": _bl.co_product_pct,
                "Ny co-prod %": _ny_co,
                "Co-prod vare": _bl.co_product_item_no,
            })
        if _rows:
            _df = pd.DataFrame(_rows)
            bom_scrap_df = mo.ui.data_editor(_df, editable_columns=["Nytt svinn %", "Ny co-prod %"])

    bom_scrap_df
    return (bom_scrap_df,)


@app.cell
def _(bom_scrap_df, overrides):
    if bom_scrap_df is not None and bom_scrap_df.value is not None:
        _df = bom_scrap_df.value
        for _, _row in _df.iterrows():
            _komponent = str(_row["Komponent"]).split(" · ")[0]
            _produkt = str(_row["Produkt"]).split(" · ")[0]
            _key = (_produkt, _komponent)

            _org_svinn = _row["Org. svinn %"]
            _ny_svinn = _row["Nytt svinn %"]
            if abs(_ny_svinn - _org_svinn) > 0.001:
                overrides["bom_scrap"][_key] = _ny_svinn
            elif _key in overrides["bom_scrap"]:
                del overrides["bom_scrap"][_key]

            _org_co = _row["Org. co-prod %"]
            _ny_co = _row["Ny co-prod %"]
            if abs(_ny_co - _org_co) > 0.001:
                overrides["bom_co_product"][_key] = _ny_co
            elif _key in overrides["bom_co_product"]:
                del overrides["bom_co_product"][_key]
    return


@app.cell
def _(mo):
    mo.md("**🏭 Arbeidssentre (timekostnad)**")
    return


@app.cell
def _(data, filtered_work_centers, mo, overrides, pd):
    wc_cost_df = None
    if data:
        _rows = []
        for _wc in filtered_work_centers:
            _saved = overrides["work_centers"].get(_wc.code, {})
            _ny_lønn = _saved.get("labor_cost_hour", _wc.labor_cost_hour)
            _ny_maskin = _saved.get("machine_cost_hour", _wc.machine_cost_hour)
            _ny_overhead = _saved.get("overhead_cost_hour", _wc.overhead_cost_hour)
            _ny_eff = _saved.get("effective_capacity_pct", _wc.effective_capacity_pct)
            _rows.append({
                "Kode": _wc.code,
                "Beskrivelse": f"{_wc.description} · {_wc.location_code}",
                "Org. lønn": _wc.labor_cost_hour,
                "Ny lønn": _ny_lønn,
                "Org. maskin": _wc.machine_cost_hour,
                "Ny maskin": _ny_maskin,
                "Org. overhead": _wc.overhead_cost_hour,
                "Ny overhead": _ny_overhead,
                "Org. eff. %": _wc.effective_capacity_pct,
                "Ny eff. %": _ny_eff,
            })
        if _rows:
            _df = pd.DataFrame(_rows)
            wc_cost_df = mo.ui.data_editor(_df, editable_columns=["Ny lønn", "Ny maskin", "Ny overhead", "Ny eff. %"])

    wc_cost_df
    return (wc_cost_df,)


@app.cell
def _(overrides, wc_cost_df):
    if wc_cost_df is not None and wc_cost_df.value is not None:
        _df = wc_cost_df.value
        for _, _row in _df.iterrows():
            _kode = _row["Kode"]
            _wc_overrides = {}
            if abs(_row["Ny lønn"] - _row["Org. lønn"]) > 0.001:
                _wc_overrides["labor_cost_hour"] = _row["Ny lønn"]
            if abs(_row["Ny maskin"] - _row["Org. maskin"]) > 0.001:
                _wc_overrides["machine_cost_hour"] = _row["Ny maskin"]
            if abs(_row["Ny overhead"] - _row["Org. overhead"]) > 0.001:
                _wc_overrides["overhead_cost_hour"] = _row["Ny overhead"]
            if abs(_row["Ny eff. %"] - _row["Org. eff. %"]) > 0.001:
                _wc_overrides["effective_capacity_pct"] = _row["Ny eff. %"]

            if _wc_overrides:
                overrides["work_centers"][_kode] = _wc_overrides
            elif _kode in overrides["work_centers"]:
                del overrides["work_centers"][_kode]
    return


@app.cell
def _(mo):
    mo.md("**📋 Routing (stykkpris)**")
    return


@app.cell
def _(data, filtered_operations, filtered_routing_lines, filtered_work_centers, mo, overrides, pd):
    routing_df = None
    if data:
        _rows = []
        for _rl in filtered_routing_lines:
            _op = next((o for o in filtered_operations if o.code == _rl.operation_code), None)
            _op_desc = _op.description if _op else _rl.operation_code
            _wc = next((w for w in filtered_work_centers if w.code == _rl.work_center_code), None)
            _loc = _wc.location_code if _wc else ""

            _key = (_rl.item_no, _rl.operation_no, _rl.work_center_code)
            _saved = overrides["routing"].get(_key, {})
            _ny_run_time = _saved.get("run_time_minutes", _rl.run_time_minutes)
            _ny_batch = _saved.get("batch_size", _rl.batch_size)
            _ny_setup = _saved.get("setup_time_minutes", _rl.setup_time_minutes)

            _rows.append({
                "Produkt": _rl.item_no,
                "Operasjon": f"{_rl.operation_code} ({_op_desc})",
                "Arb.senter": _rl.work_center_code,
                "Lokasjon": _loc,
                "Org. run time": _rl.run_time_minutes,
                "Ny run time": _ny_run_time,
                "Org. setup (min)": _rl.setup_time_minutes,
                "Ny setup (min)": _ny_setup,
                "Org. batch": _rl.batch_size,
                "Ny batch": _ny_batch,
                "_operation_no": _rl.operation_no,
            })
        if _rows:
            _df = pd.DataFrame(_rows)
            routing_df = mo.ui.data_editor(_df, editable_columns=["Ny run time", "Ny setup (min)", "Ny batch"])

    routing_df
    return (routing_df,)


@app.cell
def _(overrides, routing_df):
    if routing_df is not None and routing_df.value is not None:
        _df = routing_df.value
        for _, _row in _df.iterrows():
            _produkt = _row["Produkt"]
            _wc = _row["Arb.senter"]
            _op_no = _row["_operation_no"]
            _key = (_produkt, _op_no, _wc)

            _rt_overrides = {}
            if abs(_row["Ny run time"] - _row["Org. run time"]) > 0.001:
                _rt_overrides["run_time_minutes"] = _row["Ny run time"]
            if abs(_row["Ny setup (min)"] - _row["Org. setup (min)"]) > 0.001:
                _rt_overrides["setup_time_minutes"] = _row["Ny setup (min)"]
            if abs(_row["Ny batch"] - _row["Org. batch"]) > 0.001:
                _rt_overrides["batch_size"] = _row["Ny batch"]

            if _rt_overrides:
                overrides["routing"][_key] = _rt_overrides
            elif _key in overrides["routing"]:
                del overrides["routing"][_key]
    return


@app.cell
def _(mo):
    mo.md("**📦 Planlagt kvantum**")
    return


@app.cell
def _(mo):
    mo.md("""
    **Planlagt kvantum** bestemmer hvor mye oppstartskostnad som fordeles per enhet.

    Formel: `Oppstartskost per enhet = Oppstartstid (timer) × timekostnad / kvantum`

    Jo høyere kvantum, desto lavere blir oppstartskostnaden per produsert enhet.
    """)
    return


@app.cell
def _(mo):
    planned_qty = mo.ui.number(
        label="Planlagt kvantum (stk)",
        start=1,
        stop=100000,
        step=1,
        value=1000,
    )
    planned_qty
    return (planned_qty,)


@app.cell
def _(mo):
    mo.md("### 🚀 Resultater")
    return


@app.cell
def _(SimulationEngine, SimulationOverride, data, mo, overrides, pd, planned_qty, run_button):
    sim_results = None
    sim_overrides = None

    if run_button.value:
        if not data:
            mo.output.replace(mo.md("### ❌ Ingen data lastet. Last opp Excel-fil for å komme i gang."))
        else:
            try:
                _overrides = SimulationOverride(planned_quantity=planned_qty.value)

                for _varenr, _ny_pris in overrides["item_costs"].items():
                    _overrides.item_costs[_varenr] = _ny_pris

                for _key, _ny_svinn in overrides["bom_scrap"].items():
                    _overrides.bom_scrap[_key] = _ny_svinn

                for _kode, _wc_overrides in overrides["work_centers"].items():
                    _overrides.work_centers[_kode] = _wc_overrides

                for _key, _rt_overrides in overrides["routing"].items():
                    _overrides.routing[_key] = _rt_overrides

                for _key, _ny_co in overrides["bom_co_product"].items():
                    _overrides.bom_co_product[_key] = _ny_co

                _engine = SimulationEngine(data)
                _comparisons = _engine.compare_all(_overrides)

                if _comparisons:
                    sim_results = _comparisons
                    sim_overrides = _overrides

                    _outputs = []

                    _rows = []
                    for _c in _comparisons:
                        _rows.append({
                            "Lokasjon": _c.location_code,
                            "Produkt": _c.product_no,
                            "Beskrivelse": _c.product_desc,
                            "Org. materialkost": round(_c.original_material_cost, 2),
                            "Sim. materialkost": round(_c.simulated_material_cost, 2),
                            "Diff material": round(_c.material_diff, 2),
                            "Org. operasjonskost": round(_c.original_operation_cost, 2),
                            "Sim. operasjonskost": round(_c.simulated_operation_cost, 2),
                            "Diff operasjon": round(_c.operation_diff, 2),
                            "Org. oppstartkost": round(_c.original_setup_cost, 2),
                            "Sim. oppstartkost": round(_c.simulated_setup_cost, 2),
                            "Diff oppstart": round(_c.setup_diff, 2),
                            "Org. brutto": round(_c.original_gross_cost, 2),
                            "Sim. brutto": round(_c.simulated_gross_cost, 2),
                            "Diff brutto": round(_c.gross_diff, 2),
                            "Org. biprodukt": round(_c.original_byproduct_value, 2),
                            "Sim. biprodukt": round(_c.simulated_byproduct_value, 2),
                            "Diff biprodukt": round(_c.byproduct_diff, 2),
                            "Org. netto": round(_c.original_net_cost, 2),
                            "Sim. netto": round(_c.simulated_net_cost, 2),
                            "Diff netto": round(_c.net_diff, 2),
                        })
                    _df = pd.DataFrame(_rows)
                    _outputs.append(mo.md("**📊 Sammenligning: Opprinnelig vs Simulert**"))
                    _outputs.append(mo.ui.table(_df, selection=None))

                    if _comparisons[0].planned_quantity:
                        _sc_rows = []
                        for _c in _comparisons:
                            _sc_rows.append({
                                "Lokasjon": _c.location_code,
                                "Produkt": _c.product_no,
                                "Kvantum": f"{_c.planned_quantity:.0f}",
                                "Total netto kost": f"{_c.simulated_total_net_cost:,.2f}",
                                "Kost per enhet": f"{_c.simulated_cost_per_unit:.2f}",
                                "Timebehov": f"{_c.simulated_total_hours:.2f} timer",
                            })
                        _sc_df = pd.DataFrame(_sc_rows)
                        _outputs.append(mo.md("**📦 Scenariototaler**"))
                        _outputs.append(mo.ui.table(_sc_df, selection=None))

                    mo.output.replace(mo.vstack(_outputs))
                else:
                    mo.output.replace(mo.md("### Ingen resultater funnet"))

            except Exception as _e:
                mo.output.replace(mo.md(f"### ❌ Feil ved simulering: {_e}"))
                import traceback as _traceback
                _traceback.print_exc()
    return sim_overrides, sim_results


@app.cell
def _(mo):
    mo.md("### 💾 Eksport")
    return


@app.cell
def _(mo):
    export_pdf_button = mo.ui.run_button(label="📄 Eksporter til PDF-rapport")
    pdf_kommentar = mo.ui.text_area(
        label="Ledelsessammendrag / Kommentar (vises øverst i PDF)",
        placeholder="Skriv f.eks. anbefalinger eller en kort analyse av simuleringen her...",
        rows=4,
    )
    pdf_inkluder_detaljer = mo.ui.checkbox(
        label="Inkluder tekniske detaljer (material- og operasjonsdetaljer per produkt)",
        value=False,
    )
    mo.vstack([
        mo.md("### 📊 Tilpass PDF-rapport"),
        pdf_kommentar,
        pdf_inkluder_detaljer,
        export_pdf_button,
    ])
    return export_pdf_button, pdf_inkluder_detaljer, pdf_kommentar


@app.cell
def _(export_pdf_button, generer_rapport, mo, os, pdf_inkluder_detaljer, pdf_kommentar, registrer_fonter, sim_overrides, sim_results, tempfile):
    if export_pdf_button.value:
        try:
            if not sim_results:
                mo.output.replace(mo.md("### ❌ Ingen simuleringsresultater. Kjør en simulering først."))
            else:
                _sammenligninger = []
                for _c in sim_results:
                    _mat_det = []
                    if _c.simulated_material_details:
                        for _md in _c.simulated_material_details:
                            _mat_det.append({"komponent": _md.component, "qty_per": _md.quantity_per, "scrap_pct": _md.scrap_pct, "enhetskost": _md.unit_cost, "total_kost": _md.total_cost})
                    _op_det = []
                    if _c.simulated_operation_details:
                        for _od in _c.simulated_operation_details:
                            _op_det.append({"operasjon": _od.operation_no, "arbeidssenter": _od.work_center, "run_time": _od.run_time_min, "batch": _od.batch_size, "kost_per_time": _od.cost_per_hour, "total_kost": _od.total_cost})
                    _bp_det = []
                    if _c.simulated_byproduct_details:
                        for _bd in _c.simulated_byproduct_details:
                            _bp_det.append({"biprodukt": _bd.item_no, "kvantum": _bd.quantity, "markedsverdi": _bd.market_value, "total_verdi": _bd.total_value})
                    _co_det = []
                    if _c.simulated_co_product_results:
                        for _co in _c.simulated_co_product_results:
                            _co_det.append({"produkt": _co.product_no, "beskrivelse": _co.product_desc, "materialkost": _co.material_cost, "operasjonskost": _co.operation_cost, "setupkost": _co.setup_cost, "brutto": _co.gross_production_cost, "biproduktverdi": _co.by_product_value, "netto": _co.net_production_cost})
                    _sammenligninger.append({
                        "produkt": _c.product_no, "beskrivelse": _c.product_desc,
                        "org_netto": _c.original_net_cost, "sim_netto": _c.simulated_net_cost, "diff_netto": _c.net_diff,
                        "materialdetaljer": _mat_det, "operasjonsdetaljer": _op_det, "biprodukter": _bp_det, "coprodukter": _co_det,
                        "kvantum": _c.planned_quantity, "total_netto": _c.simulated_total_net_cost,
                        "kost_per_enhet": _c.simulated_cost_per_unit, "timebehov": _c.simulated_total_hours,
                    })

                _overrides_dict = {}
                if sim_overrides:
                    _overrides_dict = {
                        "item_costs": getattr(sim_overrides, 'item_costs', {}),
                        "bom_scrap": dict(getattr(sim_overrides, 'bom_scrap', {})),
                        "work_centers": getattr(sim_overrides, 'work_centers', {}),
                        "routing": dict(getattr(sim_overrides, 'routing', {})),
                        "planned_quantity": getattr(sim_overrides, 'planned_quantity', None),
                    }

                _wc_hours = {}
                for _c in sim_results:
                    if _c.simulated_operation_details and _c.planned_quantity:
                        for _od in _c.simulated_operation_details:
                            _wc = _od.work_center
                            _run_hrs = (_od.run_time_min / 60.0) * _c.planned_quantity
                            _setup_hrs = (_od.setup_time_min / 60.0) * (_c.planned_quantity / _od.batch_size) if _od.batch_size else 0.0
                            _wc_hours[_wc] = _wc_hours.get(_wc, 0.0) + (_run_hrs + _setup_hrs)

                registrer_fonter()
                _output_path = os.path.join(tempfile.gettempdir(), "simuleringsrapport.pdf")
                generer_rapport(_sammenligninger, _overrides_dict, _output_path,
                                tittel="Produksjonskost Simulator - Simuleringsrapport",
                                kommentar=pdf_kommentar.value, inkluder_detaljer=pdf_inkluder_detaljer.value, kapasitet_data=_wc_hours)

                with open(_output_path, "rb") as _f:
                    _pdf_content = _f.read()

                mo.output.replace(mo.vstack([
                    mo.md("### ✅ PDF-rapport generert!"),
                    mo.download(label="📥 Last ned PDF-rapport", filename="simuleringsrapport.pdf", data=_pdf_content),
                ]))

        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil ved generering av PDF: {_e}"))
            import traceback as _traceback
            _traceback.print_exc()
    return


@app.cell
def _(mo):
    export_excel_button = mo.ui.run_button(label="📊 Eksporter resultater til Excel", kind="neutral")
    export_excel_button
    return (export_excel_button,)


@app.cell
def _(export_excel_button, generer_excel_rapport, mo, os, sim_results, tempfile):
    if export_excel_button.value:
        if not sim_results:
            mo.output.replace(mo.md("### ❌ Ingen simuleringsresultater. Kjør en simulering først."))
        else:
            try:
                _output_path = os.path.join(tempfile.gettempdir(), "simuleringsresultater.xlsx")
                _excel_data = generer_excel_rapport(sim_results, _output_path)
                mo.output.replace(mo.vstack([
                    mo.md("### ✅ Excel-rapport generert!"),
                    mo.download(label="📥 Last ned Excel-rapport", filename="simuleringsresultater.xlsx", data=_excel_data),
                ]))
            except Exception as _e:
                mo.output.replace(mo.md(f"### ❌ Feil ved generering av Excel: {_e}"))
                import traceback as _traceback
                _traceback.print_exc()
    return


# ═══ 📋 Datamodell ═══════════════════════════════════════════════════


@app.cell
def _(mo):
    mo.md("## 📋 Datamodell")
    return


@app.cell
def _(mo):
    mo.md("""
    Modellen består av **10 ark** i Excel. Hvert ark har en spesifikk rolle:

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
    | 10 | **Production Scenario** | Produksjonsscenarioer | Scenario Name, Product, Planned Quantity |
    """)
    return


@app.cell
def _(data, filtered_bom_lines, filtered_byproduct_rules, filtered_capacity_days, filtered_item_costs, filtered_locations, filtered_operations, filtered_products, filtered_routing_lines, filtered_scenarios, filtered_work_centers, mo, pd):
    if data:
        _outputs = []

        _prod_rows = []
        for _p in filtered_products:
            _prod_rows.append({"Varenr": _p.item_no, "Beskrivelse": _p.description, "Type": _p.item_type, "Gruppe": _p.product_group, "Enhet": _p.base_uom, "Aktiv": "Ja" if _p.active else "Nei"})
        _outputs.append(mo.md("**📦 Product Master (Vareregister)**"))
        _outputs.append(mo.ui.table(pd.DataFrame(_prod_rows), selection=None))

        if filtered_locations:
            _loc_rows = []
            for _loc in filtered_locations:
                _loc_rows.append({"Kode": _loc.code, "Navn": _loc.name, "Type": _loc.location_type})
            _outputs.append(mo.md("**🏭 Locations (Lokasjoner)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_loc_rows), selection=None))

        if filtered_work_centers:
            _wc_rows = []
            for _wc in filtered_work_centers:
                _wc_rows.append({"Arbeidssenter": _wc.code, "Beskrivelse": f"{_wc.description} · {_wc.location_code}", "Lønn/time": _wc.labor_cost_hour, "Maskin/time": _wc.machine_cost_hour, "Overhead/time": _wc.overhead_cost_hour, "Totalt/time": _wc.total_cost_hour, "Timer/dag": _wc.capacity_hours_day, "Eff. %": _wc.effective_capacity_pct})
            _outputs.append(mo.md("**🏭 Work Centers (Arbeidssentre)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_wc_rows), selection=None))

        if filtered_operations:
            _op_rows = []
            for _op in filtered_operations:
                _op_rows.append({"Kode": _op.code, "Beskrivelse": _op.description, "Standard arb.senter": _op.default_work_center, "Enhet": _op.standard_unit})
            _outputs.append(mo.md("**⚙️ Operation Master (Operasjoner)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_op_rows), selection=None))

        if filtered_item_costs:
            _ic_rows = []
            for _ic in filtered_item_costs:
                _ic_rows.append({"Varenr": _ic.item_no, "Kosttype": _ic.cost_type, "Enhetskost": _ic.unit_cost, "Valuta": _ic.currency, "Gyldig fra": _ic.effective_date})
            _outputs.append(mo.md("**💰 Item Costs (Kostpriser)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_ic_rows), selection=None))

        if filtered_bom_lines:
            _bom_rows = []
            for _bl in filtered_bom_lines:
                _bom_rows.append({"Foreldreprodukt": _bl.parent_item_no, "Komponent": _bl.component_item_no, "Qty per": _bl.quantity_per, "Enhet": _bl.uom, "Svinn %": _bl.scrap_pct, "Co-Prod %": _bl.co_product_pct, "Co-Prod Vare": _bl.co_product_item_no})
            _outputs.append(mo.md("**🔗 BOM (Stykkliste)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_bom_rows), selection=None))

        if filtered_routing_lines:
            _rt_rows = []
            for _rl in filtered_routing_lines:
                _op = next((o for o in filtered_operations if o.code == _rl.operation_code), None)
                _op_desc = _op.description if _op else _rl.operation_code
                _wc_c = next((w for w in filtered_work_centers if w.code == _rl.work_center_code), None)
                _loc = _wc_c.location_code if _wc_c else ""
                _rt_rows.append({"Produkt": _rl.item_no, "Op.nr": _rl.operation_no, "Operasjon": _rl.operation_code, "Arbeidssenter": _rl.work_center_code, "Lokasjon": _loc, "Setup (min)": _rl.setup_time_minutes, "Run time (min)": _rl.run_time_minutes, "Batch": _rl.batch_size})
            _outputs.append(mo.md("**📋 Routing (Produksjonsflyt)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_rt_rows), selection=None))

        if filtered_byproduct_rules:
            _bp_rows = []
            for _br in filtered_byproduct_rules:
                _bp_rows.append({"Produkt": _br.parent_item_no, "Biprodukt": _br.by_product_item_no, "Forventet qty": _br.expected_quantity, "Enhet": _br.uom, "Markedsverdi": _br.market_value, "Allokeringsmetode": _br.allocation_method})
            _outputs.append(mo.md("**♻️ By Product Rules (Biproduktregler)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_bp_rows), selection=None))

        if filtered_capacity_days:
            _cap_rows = []
            for _cd in filtered_capacity_days:
                _cap_rows.append({"Arbeidssenter": _cd.work_center, "Dato": _cd.date, "Tilgj. timer": _cd.available_hours, "Planlagt nedetid": _cd.planned_downtime, "Prod. timer": _cd.available_production_hours})
            _outputs.append(mo.md("**📅 Capacity Calendar (Kapasitetskalender)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_cap_rows), selection=None))

        if filtered_scenarios:
            _sc_rows = []
            for _sc in filtered_scenarios:
                _sc_rows.append({"Scenario": _sc.scenario_name, "Produkt": _sc.product, "Planlagt kvantum": _sc.planned_quantity, "Startdato": _sc.start_date, "Sluttdato": _sc.end_date})
            _outputs.append(mo.md("**🎯 Production Scenario (Produksjonsscenarioer)**"))
            _outputs.append(mo.ui.table(pd.DataFrame(_sc_rows), selection=None))

        mo.output.replace(
            mo.accordion({
                "🔍 Vis alle data (BOM, routing, arbeidssentre…)": mo.vstack(_outputs)
            })
        )
    return


# ═══ SIDEBAR: Navigasjon og TOC ═══════════════════════════════════════


@app.cell
def _(mo):
    vareFilter = mo.ui.text(label="🔍 Filtrer på varenr eller beskrivelse")
    run_button = mo.ui.run_button(label="⚡Start simulering", kind="neutral")
    
    mo.sidebar(
        mo.vstack([
            # Handlinger
            mo.Html('<div class="fti-sidebar-section"><h3>Handlinger</h3></div>'),
            run_button,
            vareFilter,
            
            # Outline / TOC
            mo.Html('<div class="fti-sidebar-section"><h3>📑 Innholdsfortegnelse</h3></div>'),
            mo.outline(label=""),
            
            mo.Html('<div style="margin-top: auto; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 0.75em; color: #6B8F7D;">'),
            mo.Html('Fram Treindustri -'),
            mo.Html('v0.23.14 · Marimo'),
        ]),
        width="260px",
    )
    return (vareFilter, run_button)


# ═══ 📋 Endringslogg ════════════════════════════════════════════════


@app.cell
def _(mo):
    mo.md("### 📋 Endringslogg")
    return


@app.cell
def _(DataRepo, mo, pd):
    _db = DataRepo()
    _db.initialize()

    _changes = _db.get_changes(limit=50)
    if _changes:
        _rows = []
        for _c in _changes:
            _rows.append({
                "Tidspunkt": _c["timestamp"],
                "Bruker": _c["user"] or "-",
                "Kilde": _c["source"],
                "Tabell": _c["table_name"],
                "Nøkkel": _c["record_key"],
                "Felt": _c["field_name"],
                "Gammel": (_c["old_value"] or "-")[:30],
                "Ny": (_c["new_value"] or "-")[:30],
            })
        _df = pd.DataFrame(_rows)
        mo.output.replace(mo.ui.table(_df, selection=None))
    else:
        mo.output.replace(mo.md("*(Ingen endringer logget)*"))
    return


# ═══ FOOTER ═════════════════════════════════════════════════════════════


@app.cell
def _(mo):
    mo.Html('<div class="fti-footer">🌲 Fram Treindustri — Produksjonskost Simulator · Standardkostkalkyle & simulering</div>')
    return


if __name__ == "__main__":
    app.run()
