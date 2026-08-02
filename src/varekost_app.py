import marimo

__generated_with = "0.23.14"
app = marimo.App(width="medium", app_title="Produksjonskost Simulator", html_head_file="head.html", css_file="app_style.css")


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
        expand_product_costs_with_transport,
        expand_simulations_with_transport,
    )
    from generer_pdf_rapport import generer_rapport, registrer_fonter, _hent_logo
    from generer_excel_rapport import generer_excel_rapport
    from data_repo import DataRepo
    from excel_bridge import import_excel_to_sqlite, export_sqlite_to_excel, validate_excel, sync_transport_varer
    from testdata_for_app import seed_test_db, er_test_modus, reset_test_db

    return (
        CostCalculator,
        DataRepo,
        SimulationEngine,
        SimulationOverride,
        SqliteData,
        er_test_modus,
        export_sqlite_to_excel,
        generer_excel_rapport,
        generer_rapport,
        import_excel_to_sqlite,
        mo,
        os,
        pd,
        registrer_fonter,
        reset_test_db,
        seed_test_db,
        sync_transport_varer,
        tempfile,
        validate_excel,
        expand_product_costs_with_transport,
        expand_simulations_with_transport,
    )




@app.cell
def _(mo):
    mo.Html("""
    <div class="fti-header" id="fti-header">
        <div style="display: flex; align-items: center; gap: 16px;">
            <img src="https://framtreindustri.no/wp-content/uploads/2025/08/logo-liggende-2048x512.png"
                 alt="FramTre industri" style="height: 48px;" />
            <div>
                <div class="fti-header-title">Produksjonskost Simulator</div>
                <div class="fti-header-subtitle">FramTre industri — Standardkostkalkyle, simulering og analyse</div>
            </div>
        </div>
        <div style="text-align: right; color: #C6E6D0; font-size: 0.85em;">
            Last opp Excel-modell · Juster parametere · Se kostnader i sanntid
        </div>
    </div>
    """)
    return


@app.cell
def _(mo):
    overrides = {
        "item_costs": {},
        "bom_scrap": {},
        "bom_co_product": {},
        "work_centers": {},
        "routing": {},
        "transport_ruter": {},
        "byproduct_values": {},
    }
    get_reload, set_reload = mo.state(0)
    # Test-seed state: sikrer at testdata kun opprettes én gang
    get_test_seeded, set_test_seeded = mo.state(False)
    return get_reload, get_test_seeded, overrides, set_reload, set_test_seeded


@app.cell
def _(CostCalculator, DataRepo, SimulationEngine, SqliteData, expand_product_costs_with_transport, get_reload, get_test_seeded, mo, seed_test_db, er_test_modus, reset_test_db, set_test_seeded):
    _db = DataRepo()
    _db.initialize()
    data = None
    engine = None
    baseline = None
    db_stats = None
    _reload_verdi = get_reload()

    # Test-modus: seed testdata KUN én gang per app-livssyklus.
    # Uten denne staten ville hver ctrl+F5-refresh slette alle
    # transport-flagg/semi-finished brukeren har opprettet.
    if er_test_modus() and not get_test_seeded():
        try:
            _db = reset_test_db()  # ny tom test-DB
            seed_test_db(_db)
            set_test_seeded(True)
        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil ved seeding av testdata: {_e}"))
    # LEGACY (inaktiviert): sync_transport_varer() muterte datamodellen med semi-finished.
    # Den logikken er beholdt i excel_bridge.py for fremtidig Business Central-integrasjon.
    # Transport vises nå KUN i simuleringen via expand_product_costs_with_transport().

    if _db.is_empty():
        mo.output.replace(mo.md("""
        ### Velkommen!
        Du må laste opp en Excel-fil med data for å komme i gang.
        ---
        """))
    else:
        try:
            data = SqliteData()
            engine = SimulationEngine(data)
            _calculator = CostCalculator(data)
            baseline = _calculator.calculate_all()
            # Utvid baseline med fiktive transportrader (KOD→KV, KOD→EIK osv.)
            baseline = expand_product_costs_with_transport(baseline, data, _db)
            db_stats = _db.stats
        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil ved lasting: {_e}"))
            import traceback as _traceback
            _traceback.print_exc()
    return baseline, data


# ═══════════════════════════════════
# DATAIMPORT & VERSJONER
# ═══════════════════════════════════


@app.cell
def _(mo):
    excel_import_file = mo.ui.file(label="📄 Velg Excel-fil", filetypes=[".xlsx"], multiple=False)
    excel_import_kommentar = mo.ui.text(
        label="Hva er endret? (valgfri kommentar)",
        placeholder="f.eks. 'Oppdaterte råvarepriser Q3' eller 'Ny BOM for Panel'",
    )
    return excel_import_file, excel_import_kommentar


@app.cell
def _(
    excel_import_file, excel_import_kommentar, import_excel_to_sqlite, mo, os,
    overrides, set_reload, tempfile, validate_excel,
):
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
            _outputs = [mo.md(f"### Kontroll av '{_upload.name}'")]
            if _validation["stats"]:
                _stat_rows = [f"  - {_sheet}: {_count} rader" for _sheet, _count in _validation["stats"].items()]
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
                overrides["item_costs"].clear()
                overrides["bom_scrap"].clear()
                overrides["bom_co_product"].clear()
                overrides["work_centers"].clear()
                overrides["routing"].clear()
                set_reload(lambda v: v + 1)
            mo.output.replace(mo.vstack(_outputs))
        except Exception as _e:
            mo.output.replace(mo.md(f"### ❌ Feil: {_e}"))
            import traceback as _traceback
            _traceback.print_exc()
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
            _label = f"{_ts} – {_fn} ({_cm})" if _cm else f"{_ts} – {_fn}"
            _options[_label] = _label
            upload_id_map[_label] = _u["id"]
        historikk_valg = mo.ui.dropdown(options=_options, label="📜 Velg tidligere versjon", value=None)
    return historikk_valg, upload_id_map


@app.cell
def _(
    DataRepo, historikk_valg, import_excel_to_sqlite, mo, os,
    overrides, set_reload, tempfile, upload_id_map,
):
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
                    _stats = import_excel_to_sqlite(_temp_path, excel_blob=_blob, comment=f"Gjeninnlasting av versjon {_upload_id}")
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
    return


@app.cell
def _(mo):
    export_excel_db_button = mo.ui.run_button(label="📥 Last ned komplett datafil")
    return (export_excel_db_button,)


@app.cell
def _(export_excel_db_button, export_sqlite_to_excel, mo, os, tempfile):
    db_download = None
    if export_excel_db_button.value:
        try:
            _output_path = os.path.join(tempfile.gettempdir(), "produksjonsmodell_eksport.xlsx")
            export_sqlite_to_excel(_output_path)
            with open(_output_path, "rb") as _f:
                _excel_content = _f.read()
            db_download = mo.download(label="📥 Last ned (11 ark, inkl. endringslogg)", filename="produksjonsmodell_eksport.xlsx", data=_excel_content)
        except Exception as _e:
            pass
    return (db_download,)


# ═══════════════════════════════════
# FILTER
# ═══════════════════════════════════


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
    filter_info = ""
    if data:
        _filter_text = vareFilter.value.strip().lower() if vareFilter.value else ""
        if _filter_text:
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
                if _p.item_no in _all_product_set: filtered_products.append(_p)
            for _br in data.byproduct_rules:
                if _br.parent_item_no in _fg_set: filtered_byproduct_rules.append(_br)
            _routing_wc_set = set()
            for _rl in data.routing_lines:
                if _rl.item_no in _fg_set:
                    filtered_routing_lines.append(_rl)
                    _routing_wc_set.add(_rl.work_center_code)
            for _ic in data.item_costs:
                if _ic.item_no in _all_product_set: filtered_item_costs.append(_ic)
            for _sc in data.scenarios:
                if _sc.product in _fg_set: filtered_scenarios.append(_sc)
            for _wc in data.work_centers:
                if _wc.code in _routing_wc_set: filtered_work_centers.append(_wc)
            for _loc in data.locations:
                if _loc.code in {_wc.location_code for _wc in filtered_work_centers}: filtered_locations.append(_loc)
            for _op in data.operations:
                if _op.code in {_rl.operation_code for _rl in filtered_routing_lines}: filtered_operations.append(_op)
            for _cd in data.capacity_days:
                if _cd.work_center in _routing_wc_set: filtered_capacity_days.append(_cd)
            filter_info = f"ℹ️ Viser {len(filtered_products)} av {len(data.products)} produkter (filtrert på \"{_filter_text}\")"
        else:
            filtered_products = list(data.products); filtered_bom_lines = list(data.bom_lines); filtered_routing_lines = list(data.routing_lines); filtered_byproduct_rules = list(data.byproduct_rules); filtered_item_costs = list(data.item_costs); filtered_scenarios = list(data.scenarios); filtered_work_centers = list(data.work_centers); filtered_locations = list(data.locations); filtered_operations = list(data.operations); filtered_capacity_days = list(data.capacity_days)
    return (filter_info, filtered_bom_lines, filtered_byproduct_rules, filtered_capacity_days, filtered_item_costs, filtered_locations, filtered_operations, filtered_products, filtered_routing_lines, filtered_scenarios, filtered_work_centers)


# ═══════════════════════════════════
# SIMULERING - parametere
# ═══════════════════════════════════


@app.cell
def _(data, filtered_products, mo, overrides, pd):
    rm_price_df = None
    if data:
        _rm_items = [p for p in filtered_products if p.item_type == 'Raw Material']
        _rows = []
        for _p in _rm_items:
            _cost = next((c.unit_cost for c in data.item_costs if c.item_no == _p.item_no), 0)
            _ny_cost = overrides["item_costs"].get(_p.item_no, _cost)
            _rows.append({"Varenr": _p.item_no, "Beskrivelse": f"{_p.description} · {_p.base_uom}", "Org. pris": _cost, "Ny pris": _ny_cost})
        if _rows:
            _df = pd.DataFrame(_rows)
            rm_price_df = mo.ui.data_editor(_df, editable_columns=["Ny pris"])
    return (rm_price_df,)


@app.cell
def _(overrides, rm_price_df):
    if rm_price_df is not None and rm_price_df.value is not None:
        _df = rm_price_df.value
        for _, _row in _df.iterrows():
            _varenr = _row["Varenr"]
            _org = _row["Org. pris"]
            _ny = _row["Ny pris"]
            if abs(_ny - _org) > 0.001: overrides["item_costs"][_varenr] = _ny
            elif _varenr in overrides["item_costs"]: del overrides["item_costs"][_varenr]
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
            _rows.append({"Komponent": f"{_bl.component_item_no} · {_komp_desc}", "Produkt": f"{_bl.parent_item_no} · {_prod_desc}", "Org. svinn %": _bl.scrap_pct, "Nytt svinn %": _ny_svinn, "Org. co-prod %": _bl.co_product_pct, "Ny co-prod %": _ny_co, "Co-prod vare": _bl.co_product_item_no})
        if _rows:
            _df = pd.DataFrame(_rows)
            bom_scrap_df = mo.ui.data_editor(_df, editable_columns=["Nytt svinn %", "Ny co-prod %"])
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
            if abs(_ny_svinn - _org_svinn) > 0.001: overrides["bom_scrap"][_key] = _ny_svinn
            elif _key in overrides["bom_scrap"]: del overrides["bom_scrap"][_key]
            _org_co = _row["Org. co-prod %"]
            _ny_co = _row["Ny co-prod %"]
            if abs(_ny_co - _org_co) > 0.001: overrides["bom_co_product"][_key] = _ny_co
            elif _key in overrides["bom_co_product"]: del overrides["bom_co_product"][_key]
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
            _rows.append({"Kode": _wc.code, "Beskrivelse": f"{_wc.description} · {_wc.location_code}", "Org. lønn": _wc.labor_cost_hour, "Ny lønn": _ny_lønn, "Org. maskin": _wc.machine_cost_hour, "Ny maskin": _ny_maskin, "Org. overhead": _wc.overhead_cost_hour, "Ny overhead": _ny_overhead, "Org. eff. %": _wc.effective_capacity_pct, "Ny eff. %": _ny_eff})
        if _rows:
            _df = pd.DataFrame(_rows)
            wc_cost_df = mo.ui.data_editor(_df, editable_columns=["Ny lønn", "Ny maskin", "Ny overhead", "Ny eff. %"])
    return (wc_cost_df,)


@app.cell
def _(overrides, wc_cost_df):
    if wc_cost_df is not None and wc_cost_df.value is not None:
        _df = wc_cost_df.value
        for _, _row in _df.iterrows():
            _kode = _row["Kode"]
            _wc_overrides = {}
            if abs(_row["Ny lønn"] - _row["Org. lønn"]) > 0.001: _wc_overrides["labor_cost_hour"] = _row["Ny lønn"]
            if abs(_row["Ny maskin"] - _row["Org. maskin"]) > 0.001: _wc_overrides["machine_cost_hour"] = _row["Ny maskin"]
            if abs(_row["Ny overhead"] - _row["Org. overhead"]) > 0.001: _wc_overrides["overhead_cost_hour"] = _row["Ny overhead"]
            if abs(_row["Ny eff. %"] - _row["Org. eff. %"]) > 0.001: _wc_overrides["effective_capacity_pct"] = _row["Ny eff. %"]
            if _wc_overrides: overrides["work_centers"][_kode] = _wc_overrides
            elif _kode in overrides["work_centers"]: del overrides["work_centers"][_kode]
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
            _rows.append({"Produkt": _rl.item_no, "Operasjon": f"{_rl.operation_code} ({_op_desc})", "Arb.senter": _rl.work_center_code, "Lokasjon": _loc, "Org. run time": _rl.run_time_minutes, "Ny run time": _saved.get("run_time_minutes", _rl.run_time_minutes), "Org. setup (min)": _rl.setup_time_minutes, "Ny setup (min)": _saved.get("setup_time_minutes", _rl.setup_time_minutes), "Org. batch": _rl.batch_size, "Ny batch": _saved.get("batch_size", _rl.batch_size), "_operation_no": _rl.operation_no})
        if _rows:
            _df = pd.DataFrame(_rows)
            routing_df = mo.ui.data_editor(_df, editable_columns=["Ny run time", "Ny setup (min)", "Ny batch"])
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
            if abs(_row["Ny run time"] - _row["Org. run time"]) > 0.001: _rt_overrides["run_time_minutes"] = _row["Ny run time"]
            if abs(_row["Ny setup (min)"] - _row["Org. setup (min)"]) > 0.001: _rt_overrides["setup_time_minutes"] = _row["Ny setup (min)"]
            if abs(_row["Ny batch"] - _row["Org. batch"]) > 0.001: _rt_overrides["batch_size"] = _row["Ny batch"]
            if _rt_overrides: overrides["routing"][_key] = _rt_overrides
            elif _key in overrides["routing"]: del overrides["routing"][_key]
    return


@app.cell
def _(DataRepo, mo, overrides, pd):
    transport_df = None
    _db = DataRepo()
    _db.initialize()
    _tr_rows = []
    try:
        _tr_rows_db = _db.conn.execute(
            "SELECT from_loc, to_loc, cost_per_m3, distance_km, hours FROM transport_ruter ORDER BY from_loc, to_loc"
        ).fetchall()
    except Exception:
        _tr_rows_db = []
    for _r in _tr_rows_db:
        _key = (_r["from_loc"], _r["to_loc"])
        _saved = overrides["transport_ruter"].get(_key, {})
        _ny_cost = _saved.get("cost_per_m3", _r["cost_per_m3"])
        _tr_rows.append({"Fra": _r["from_loc"], "Til": _r["to_loc"], "Org. kost/m³": _r["cost_per_m3"], "Ny kost/m³": _ny_cost, "Km": _r["distance_km"], "Timer": _r["hours"]})
    if _tr_rows:
        _df = pd.DataFrame(_tr_rows)
        transport_df = mo.ui.data_editor(_df, editable_columns=["Ny kost/m³"])
    return (transport_df,)


@app.cell
def _(overrides, transport_df):
    if transport_df is not None and transport_df.value is not None:
        _df = transport_df.value
        for _, _row in _df.iterrows():
            _key = (_row["Fra"], _row["Til"])
            _org = _row["Org. kost/m³"]
            _ny = _row["Ny kost/m³"]
            if abs(_ny - _org) > 0.001: overrides["transport_ruter"][_key] = {"cost_per_m3": _ny}
            elif _key in overrides["transport_ruter"]: del overrides["transport_ruter"][_key]
    return


@app.cell
def _(data, filtered_byproduct_rules, mo, overrides, pd):
    byproduct_df = None
    if data and filtered_byproduct_rules:
        _rows = []
        for _br in filtered_byproduct_rules:
            _key = (_br.parent_item_no, _br.by_product_item_no)
            _ny_verdi = overrides["byproduct_values"].get(_key, _br.market_value)
            _bp_desc = next((p.description for p in data.products if p.item_no == _br.by_product_item_no), _br.by_product_item_no)
            _rows.append({"Produkt": _br.parent_item_no, "Biprodukt": f"{_br.by_product_item_no} · {_bp_desc}", "Org. markedsverdi": _br.market_value, "Ny markedsverdi": _ny_verdi, "Forventet qty": _br.expected_quantity, "Enhet": _br.uom})
        if _rows:
            _df = pd.DataFrame(_rows)
            byproduct_df = mo.ui.data_editor(_df, editable_columns=["Ny markedsverdi"])
    return (byproduct_df,)


@app.cell
def _(byproduct_df, overrides):
    if byproduct_df is not None and byproduct_df.value is not None:
        _df = byproduct_df.value
        for _, _row in _df.iterrows():
            _produkt = _row["Produkt"]
            _biprodukt = str(_row["Biprodukt"]).split(" · ")[0]
            _key = (_produkt, _biprodukt)
            _org = _row["Org. markedsverdi"]
            _ny = _row["Ny markedsverdi"]
            if abs(_ny - _org) > 0.001: overrides["byproduct_values"][_key] = _ny
            elif _key in overrides["byproduct_values"]: del overrides["byproduct_values"][_key]
    return


@app.cell
def _(mo):
    planned_qty = mo.ui.number(label="Planlagt kvantum (stk)", start=1, stop=100000, step=1, value=1000)
    return (planned_qty,)


@app.cell
def _(mo):
    run_button = mo.ui.run_button(label="⚡ Start simulering", kind="neutral", full_width=True)
    return (run_button,)


# ── Simuleringslogikk (med spinner) ─────────


@app.cell
def _(SimulationEngine, SimulationOverride, data, expand_simulations_with_transport, mo, overrides, pd, planned_qty, run_button):
    sim_results = None
    sim_overrides = None
    if run_button.value:
        if not data:
            mo.output.replace(mo.md("### ❌ Ingen data lastet. Last opp Excel-fil for å komme i gang."))
        else:
            try:
                _overrides = SimulationOverride(planned_quantity=planned_qty.value)
                for _varenr, _ny_pris in overrides["item_costs"].items(): _overrides.item_costs[_varenr] = _ny_pris
                for _key, _ny_svinn in overrides["bom_scrap"].items(): _overrides.bom_scrap[_key] = _ny_svinn
                for _kode, _wc_overrides in overrides["work_centers"].items(): _overrides.work_centers[_kode] = _wc_overrides
                for _key, _rt_overrides in overrides["routing"].items(): _overrides.routing[_key] = _rt_overrides
                for _key, _ny_co in overrides["bom_co_product"].items(): _overrides.bom_co_product[_key] = _ny_co
                for _key, _ny_verdi in overrides["byproduct_values"].items(): _overrides.byproduct_values[_key] = _ny_verdi
                for _key, _tr_overrides in overrides["transport_ruter"].items(): _overrides.transport_ruter[_key] = _tr_overrides
                _engine = SimulationEngine(data)
                _comparisons = _engine.compare_all(_overrides)
                if _comparisons:
                    # Utvid med fiktive transportrader (KOD→KV, KOD→EIK osv.)
                    _comparisons = expand_simulations_with_transport(
                        _comparisons, data,
                        transport_ruter_overrides=overrides["transport_ruter"],
                    )
                    sim_results = _comparisons
                    sim_overrides = _overrides
            except Exception as _e:
                mo.output.replace(mo.md(f"### ❌ Feil ved simulering: {_e}"))
                import traceback as _traceback
                _traceback.print_exc()
    return sim_overrides, sim_results


# ── Eksport-knapper ─────────────────────


@app.cell
def _(mo):
    export_pdf_button = mo.ui.run_button(label="📄 Eksporter til PDF-rapport")
    pdf_kommentar = mo.ui.text_area(label="Ledelsessammendrag / Kommentar (vises øverst i PDF)", placeholder="Skriv f.eks. anbefalinger eller en kort analyse...", rows=3)
    pdf_inkluder_detaljer = mo.ui.checkbox(label="Inkluder tekniske detaljer (material- og operasjonsdetaljer per produkt)", value=False)
    export_excel_button = mo.ui.run_button(label="📊 Eksporter resultater til Excel", kind="neutral")
    return (export_excel_button, export_pdf_button, pdf_inkluder_detaljer, pdf_kommentar)


# ── SIDEBAR ─────────────────────────


@app.cell
def _(mo):
    vareFilter = mo.ui.text(label="🔍 Filtrer på varenummer og beskrivelse", full_width=True)
    mo.sidebar(mo.vstack([
        mo.Html('<div class="fti-sidebar-section"><h3>Filter</h3></div>'),
        vareFilter,
        mo.Html('<div class="fti-sidebar-section"><h3>ℹ️ Om denne appen</h3></div>'),
        mo.md("""
        Juster råvarepriser, svinnprosenter,
        timekostnader og produksjonstider for
        å simulere kostnadsendringer i sanntid.
        
        Resultater kan eksporteres til **PDF**
        eller **Excel** for videre analyse.
        
        Bruk fanene øverst for å navigere
        mellom simulering, dataimport,
        datamodell og endringslogg.
        """),
        mo.Html('<div style="margin-top: auto; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 0.75em; color: #6B8F7D;">'),
        mo.Html('FramTre industri - marimo v0.23.14'),
    ]), width="260px")
    return (vareFilter,)


# ═══════════════════════════════════
# BRUKERDOKUMENTASJON
# ═══════════════════════════════════


@app.cell
def _(mo, os):
    _doc_path = os.path.join(os.path.dirname(__file__), '..', 'docs', 'Brukermanual_Produksjonsmodell.md')
    with open(_doc_path, "r", encoding="utf-8") as _f:
        brukermanual_md = _f.read()
    return (brukermanual_md,)


# ═══════════════════════════════════
# HOVEDLAYOUT: Faner
# ═══════════════════════════════════


@app.cell
def _(
    DataRepo, baseline, data, db_download, excel_import_file, excel_import_kommentar,
    export_excel_db_button, export_pdf_button, filter_info, historikk_valg,
    import_excel_to_sqlite, mo, overrides, pd, planned_qty, rm_price_df,
    routing_df, run_button, set_reload, validate_excel, wc_cost_df, bom_scrap_df,
    transport_df, byproduct_df,
    os, tempfile, generer_rapport, registrer_fonter, pdf_kommentar,
    pdf_inkluder_detaljer, sim_results, sim_overrides, export_sqlite_to_excel,
    generer_excel_rapport, export_excel_button, vareFilter,
    filtered_bom_lines, filtered_byproduct_rules, filtered_capacity_days,
    filtered_item_costs, filtered_locations, filtered_operations, filtered_products,
    filtered_routing_lines, filtered_scenarios, filtered_work_centers,
    brukermanual_md,
):
    # ── Fane 1: Simulering & Analyse ───────────────────────────
    _sim_parts = []
    if filter_info:
        _sim_parts.append(mo.md(f"> {filter_info}"))

    _sim_parts.append(mo.md("## 🔧 Simuleringsparametere"))

    # Samle data_editor-tabeller i accordion (Trinn 1)
    _parameter_accordions = {}
    if rm_price_df is not None: _parameter_accordions["🪵 Råvarepriser"] = rm_price_df
    if bom_scrap_df is not None: _parameter_accordions["🗑️ Svinn- og kapp-prosenter"] = bom_scrap_df
    if wc_cost_df is not None: _parameter_accordions["🏭 Arbeidssentre (timekostnad)"] = wc_cost_df
    if routing_df is not None: _parameter_accordions["📋 Routing (stykkpris/tider)"] = routing_df
    if transport_df is not None: _parameter_accordions["🚛 Transport (kr/m³)"] = transport_df
    if byproduct_df is not None: _parameter_accordions["♻️ Biproduktverdier"] = byproduct_df
    if _parameter_accordions:
        _sim_parts.append(mo.accordion(_parameter_accordions))

    # planned_qty og run_button rett under accordion
    _sim_parts.append(mo.md("### 📦 Planlagt kvantum"))
    _sim_parts.append(mo.md("Jo høyere kvantum, desto lavere oppstartskostnad per enhet."))
    _sim_parts.append(planned_qty)
    _sim_parts.append(mo.Html('<div class="fti-run-button"></div>'))
    _sim_parts.append(mo.hstack([mo.Html('<div style="flex:1"></div>'), run_button, mo.Html('<div style="flex:1"></div>')], justify="center"))

    if sim_results:
        _total_diff = sum(_c.net_diff for _c in sim_results)
        _total_org = sum(_c.original_net_cost for _c in sim_results)
        _pct_diff = (_total_diff / _total_org * 100) if _total_org else 0
        _total_hours = sum(_c.simulated_total_hours for _c in sim_results)
        _total_products = len(sim_results)

        _sim_parts.append(mo.md("## 🚀 Simuleringsresultater"))
        _sim_parts.append(mo.hstack([
            mo.Html(f'<div class="fti-kpi-card"><div class="fti-kpi-title">Totalt simulerte produkter</div><div class="fti-kpi-value">{_total_products} stk</div></div>'),
            mo.Html(f'<div class="fti-kpi-card"><div class="fti-kpi-title">Samlet kostnadsendring</div><div class="fti-kpi-value">{_total_diff:,.0f} NOK</div><div class="fti-kpi-delta">{_pct_diff:+.1f}%</div></div>'),
            mo.Html(f'<div class="fti-kpi-card"><div class="fti-kpi-title">Totalt timebehov</div><div class="fti-kpi-value">{_total_hours:.1f} t</div></div>'),
        ], justify="start"))

        _rows = []
        for _c in sim_results:
            _rows.append({"Lokasjon": _c.location_code, "Produkt": _c.product_no, "Beskrivelse": _c.product_desc, "Org. materialkost": round(_c.original_material_cost, 2), "Sim. materialkost": round(_c.simulated_material_cost, 2), "Diff material": round(_c.material_diff, 2), "Org. operasjonskost": round(_c.original_operation_cost, 2), "Sim. operasjonskost": round(_c.simulated_operation_cost, 2), "Diff operasjon": round(_c.operation_diff, 2), "Org. oppstartkost": round(_c.original_setup_cost, 2), "Sim. oppstartkost": round(_c.simulated_setup_cost, 2), "Diff oppstart": round(_c.setup_diff, 2), "Org. brutto": round(_c.original_gross_cost, 2), "Sim. brutto": round(_c.simulated_gross_cost, 2), "Diff brutto": round(_c.gross_diff, 2), "Org. biprodukt": round(_c.original_byproduct_value, 2), "Sim. biprodukt": round(_c.simulated_byproduct_value, 2), "Diff biprodukt": round(_c.byproduct_diff, 2), "Org. netto": round(_c.original_net_cost, 2), "Sim. netto": round(_c.simulated_net_cost, 2), "Diff netto": round(_c.net_diff, 2)})
        _sim_parts.append(mo.md("**📊 Sammenligning: Opprinnelig vs Simulert**"))
        _sim_parts.append(mo.ui.table(pd.DataFrame(_rows), selection=None))

        if sim_results[0].planned_quantity:
            _sc_rows = []
            for _c in sim_results:
                _sc_rows.append({"Lokasjon": _c.location_code, "Produkt": _c.product_no, "Kvantum": f"{_c.planned_quantity:.0f}", "Total netto kost": f"{_c.simulated_total_net_cost:,.2f}", "Kost per enhet": f"{_c.simulated_cost_per_unit:.2f}", "Timebehov": f"{_c.simulated_total_hours:.2f} timer"})
            _sim_parts.append(mo.md("**📦 Scenariototaler**"))
            _sim_parts.append(mo.ui.table(pd.DataFrame(_sc_rows), selection=None))

        # Eksportpanel med spinner (Trinn 3)
        _eksport_items = [export_pdf_button, pdf_kommentar, pdf_inkluder_detaljer, mo.Html('<hr style="margin: 12px 0; border-color: #E2E8F0;">'), export_excel_button]

        # Bestem hvilke produkter som skal inkluderes i PDF (respekter filteret)
        _filter_text = vareFilter.value.strip().lower() if vareFilter.value else ""
        if _filter_text:
            _pdf_filtered_fg_set = {_p.item_no for _p in filtered_products if _p.item_type in ('Finished Good', 'Semi Finished')}
            _pdf_sim_results = [_c for _c in sim_results if _c.product_no in _pdf_filtered_fg_set]
        else:
            _pdf_sim_results = list(sim_results)
        _pdf_info = f"ℹ️ PDF inkluderer {len(_pdf_sim_results)} av {len(sim_results)} produkter" + (f" (filtrert på \"{_filter_text}\")" if _filter_text else "")

        if export_pdf_button.value:
            try:
                _sammenligninger = []
                for _c in _pdf_sim_results:
                    _mat_det = [{"komponent": _md.component, "qty_per": _md.quantity_per, "scrap_pct": _md.scrap_pct, "enhetskost": _md.unit_cost, "total_kost": _md.total_cost} for _md in (_c.simulated_material_details or [])]
                    _op_det = [{"operasjon": _od.operation_no, "arbeidssenter": _od.work_center, "run_time": _od.run_time_min, "batch": _od.batch_size, "kost_per_time": _od.cost_per_hour, "total_kost": _od.total_cost} for _od in (_c.simulated_operation_details or [])]
                    _bp_det = [{"biprodukt": _bd.item_no, "kvantum": _bd.quantity, "markedsverdi": _bd.market_value, "total_verdi": _bd.total_value} for _bd in (_c.simulated_byproduct_details or [])]
                    _co_det = [{"produkt": _co.product_no, "beskrivelse": _co.product_desc, "materialkost": _co.material_cost, "operasjonskost": _co.operation_cost, "setupkost": _co.setup_cost, "brutto": _co.gross_production_cost, "biproduktverdi": _co.by_product_value, "netto": _co.net_production_cost} for _co in (_c.simulated_co_product_results or [])]
                    _sammenligninger.append({"produkt": _c.product_no, "beskrivelse": _c.product_desc, "location_code": _c.location_code, "location_name": _c.location_name, "org_netto": _c.original_net_cost, "sim_netto": _c.simulated_net_cost, "diff_netto": _c.net_diff, "materialdetaljer": _mat_det, "operasjonsdetaljer": _op_det, "biprodukter": _bp_det, "coprodukter": _co_det, "kvantum": _c.planned_quantity, "total_netto": _c.simulated_total_net_cost, "kost_per_enhet": _c.simulated_cost_per_unit, "timebehov": _c.simulated_total_hours})
                _overrides_dict = {"item_costs": getattr(sim_overrides, 'item_costs', {}), "bom_scrap": dict(getattr(sim_overrides, 'bom_scrap', {})), "work_centers": getattr(sim_overrides, 'work_centers', {}), "routing": dict(getattr(sim_overrides, 'routing', {})), "planned_quantity": getattr(sim_overrides, 'planned_quantity', None)} if sim_overrides else {}
                _wc_hours = {}
                for _c in _pdf_sim_results:
                    if _c.simulated_operation_details and _c.planned_quantity:
                        for _od in _c.simulated_operation_details:
                            _wc = _od.work_center
                            _wc_hours[_wc] = _wc_hours.get(_wc, 0.0) + ((_od.run_time_min / 60.0) * _c.planned_quantity) + ((_od.setup_time_min / 60.0) * (_c.planned_quantity / _od.batch_size) if _od.batch_size else 0.0)
                registrer_fonter()
                _output_path = os.path.join(tempfile.gettempdir(), "simuleringsrapport.pdf")
                generer_rapport(_sammenligninger, _overrides_dict, _output_path,
                                tittel="Produksjonskost Simulator - Simuleringsrapport",
                                kommentar=pdf_kommentar.value, inkluder_detaljer=pdf_inkluder_detaljer.value, kapasitet_data=_wc_hours)
                with open(_output_path, "rb") as _f:
                    _pdf_content = _f.read()
                _eksport_items.append(mo.download(label="📄 Last ned PDF-rapport", filename="simuleringsrapport.pdf", data=_pdf_content))
            except Exception as _e:
                import traceback as _traceback
                _traceback.print_exc()

        if export_excel_button.value:
            try:
                _output_path = os.path.join(tempfile.gettempdir(), "simuleringsresultater.xlsx")
                _excel_data = generer_excel_rapport(sim_results, _output_path)
                _eksport_items.append(mo.download(label="📊 Last ned Excel-rapport", filename="simuleringsresultater.xlsx", data=_excel_data))
            except Exception as _e:
                import traceback as _traceback
                _traceback.print_exc()

        _sim_parts.append(mo.md("### 💾 Eksport"))
        if _filter_text:
            _eksport_items.insert(0, mo.md(f"> {_pdf_info}"))
        _sim_parts.append(mo.accordion({"📥 Last ned rapporter (PDF / Excel)": mo.vstack(_eksport_items)}))

    _tab_simulering = mo.vstack(_sim_parts)

    # ── Fane 2: Dataimport & Versjoner ─────────────────────────
    _import_parts = []
    _import_parts.append(mo.md("**📤 Last opp Excel**"))
    _import_parts.append(mo.md("Valideres og importeres automatisk."))
    _import_parts.append(excel_import_file)
    _import_parts.append(excel_import_kommentar)
    _import_parts.append(mo.md("**📜 Versjonshistorikk**"))
    if historikk_valg is not None:
        _import_parts.append(mo.md("Velg en tidligere import for å laste den inn på nytt."))
        _import_parts.append(historikk_valg)
    _import_parts.append(mo.md("**📥 Eksporter komplett datafil**"))
    _import_parts.append(export_excel_db_button)
    if db_download is not None:
        _import_parts.append(db_download)
    _tab_import = mo.vstack(_import_parts)

    # ── Fane 3: Datamodell (Innsyn) ────────────────────────────
    _modell_parts = [mo.md("Modellen består av **10 stamdata-ark + 2 transport-ark** i Excel.\n\n| # | Arknavn | Innhold |\n|---|---------|---------|\n| 1 | **Product Master** | Vareregister |\n| 2 | **Locations** | Fabrikker og lagre |\n| 3 | **Work Centers** | Arbeidssentre med kostsatser |\n| 4 | **Operation Master** | Standardoperasjoner |\n| 5 | **Item Costs** | Kostpriser per vare |\n| 6 | **BOM** | Stykkliste |\n| 7 | **Routing** | Produksjonsflyt |\n| 8 | **By Product Rules** | Biprodukter og verdsetting |\n| 9 | **Capacity Calendar** | Kapasitetskalender *(legacy)* |\n| 10 | **Production Scenario** | Produksjonsscenarioer *(legacy – kvantum styres nå fra UI)* |\n| 11 | **Transport Flagg** | Transportvarer mellom høvlerier |\n| 12 | **Transport Ruter** | Fraktkost per M3 mellom lokasjoner |")]
    if data:
        _data_outputs = []
        _data_outputs.append(mo.md("**📦 Product Master**"))
        _data_outputs.append(mo.ui.table(pd.DataFrame([{"Varenr": _p.item_no, "Beskrivelse": _p.description, "Type": _p.item_type, "Enhet": _p.base_uom, "Aktiv": "Ja" if _p.active else "Nei"} for _p in filtered_products]), selection=None))
        if filtered_locations:
            _data_outputs.append(mo.md("**🏭 Locations**"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Kode": _loc.code, "Navn": _loc.name, "Type": _loc.location_type} for _loc in filtered_locations]), selection=None))
        if filtered_work_centers:
            _data_outputs.append(mo.md("**🏭 Work Centers**"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Kode": _wc.code, "Beskrivelse": f"{_wc.description} · {_wc.location_code}", "Lønn/time": _wc.labor_cost_hour, "Maskin/time": _wc.machine_cost_hour, "Overhead/time": _wc.overhead_cost_hour, "Totalt/time": _wc.total_cost_hour, "Eff. %": _wc.effective_capacity_pct} for _wc in filtered_work_centers]), selection=None))
        if filtered_operations:
            _data_outputs.append(mo.md("**⚙️ Operations**"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Kode": _op.code, "Beskrivelse": _op.description} for _op in filtered_operations]), selection=None))
        if filtered_item_costs:
            _data_outputs.append(mo.md("**💰 Item Costs**"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Varenr": _ic.item_no, "Enhetskost": _ic.unit_cost, "Valuta": _ic.currency} for _ic in filtered_item_costs]), selection=None))
        if filtered_bom_lines:
            _data_outputs.append(mo.md("**🔗 BOM**"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Foreldreprodukt": _bl.parent_item_no, "Komponent": _bl.component_item_no, "Qty per": _bl.quantity_per, "Svinn %": _bl.scrap_pct} for _bl in filtered_bom_lines]), selection=None))
        if filtered_routing_lines:
            _data_outputs.append(mo.md("**📋 Routing**"))
            _rt_list = []
            for _rl in filtered_routing_lines:
                _op = next((o for o in filtered_operations if o.code == _rl.operation_code), None)
                _op_desc = _op.description if _op else _rl.operation_code
                _rt_list.append({"Produkt": _rl.item_no, "Operasjon": f"{_rl.operation_code} ({_op_desc})", "Arb.senter": _rl.work_center_code, "Setup (min)": _rl.setup_time_minutes, "Run time (min)": _rl.run_time_minutes, "Batch": _rl.batch_size})
            _data_outputs.append(mo.ui.table(pd.DataFrame(_rt_list), selection=None))
        if filtered_byproduct_rules:
            _data_outputs.append(mo.md("**♻️ By Product Rules**"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Produkt": _br.parent_item_no, "Biprodukt": _br.by_product_item_no, "Forventet qty": _br.expected_quantity, "Markedsverdi": _br.market_value} for _br in filtered_byproduct_rules]), selection=None))
        if filtered_capacity_days:
            _data_outputs.append(mo.md("**📅 Capacity Calendar**"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Arbeidssenter": _cd.work_center, "Dato": _cd.date, "Prod. timer": _cd.available_production_hours} for _cd in filtered_capacity_days]), selection=None))
        if filtered_scenarios:
            _data_outputs.append(mo.md("**🎯 Production Scenario** *(legacy – ikke i aktiv bruk)*"))
            _data_outputs.append(mo.ui.table(pd.DataFrame([{"Scenario": _sc.scenario_name, "Produkt": _sc.product, "Kvantum": _sc.planned_quantity} for _sc in filtered_scenarios]), selection=None))
        _db_model = DataRepo()
        _db_model.initialize()
        try:
            _tr_ruter = _db_model.conn.execute(
                "SELECT from_loc, to_loc, cost_per_m3, distance_km, hours FROM transport_ruter ORDER BY from_loc, to_loc"
            ).fetchall()
            if _tr_ruter:
                _data_outputs.append(mo.md("**🚛 Transport Ruter**"))
                _data_outputs.append(mo.ui.table(pd.DataFrame([{"Fra": _r["from_loc"], "Til": _r["to_loc"], "Kost/m³": _r["cost_per_m3"], "Km": _r["distance_km"], "Timer": _r["hours"]} for _r in _tr_ruter]), selection=None))
            _tr_flagg = _db_model.conn.execute(
                "SELECT item_no, is_transport, updated_at FROM transport_flagg ORDER BY item_no"
            ).fetchall()
            if _tr_flagg:
                _data_outputs.append(mo.md("**🏷️ Transport Flagg**"))
                _data_outputs.append(mo.ui.table(pd.DataFrame([{"Varenr": _f["item_no"], "Transport": "Ja" if _f["is_transport"] else "Nei", "Oppdatert": _f["updated_at"]} for _f in _tr_flagg]), selection=None))
        except Exception:
            pass
        _modell_parts.append(mo.accordion({"🔍 Vis alle data": mo.vstack(_data_outputs)}))
    _tab_modell = mo.vstack(_modell_parts)

    # ── Fane 4: Endringslogg ───────────────────────────────────
    _db = DataRepo()
    _db.initialize()
    _changes = _db.get_changes(limit=50)
    if _changes:
        _ch_rows = []
        for _c in _changes:
            _ch_rows.append({"Tidspunkt": _c["timestamp"], "Bruker": _c["user"] or "-", "Kilde": _c["source"], "Tabell": _c["table_name"], "Nøkkel": _c["record_key"], "Felt": _c["field_name"], "Gammel": (_c["old_value"] or "-")[:30], "Ny": (_c["new_value"] or "-")[:30]})
        _tab_changelog = mo.vstack([mo.md("## 📋 Endringslogg"), mo.ui.table(pd.DataFrame(_ch_rows), selection=None)])
    else:
        _tab_changelog = mo.vstack([mo.md("## 📋 Endringslogg"), mo.md("*(Ingen endringer logget)*")])

    # ── Fane 5: Brukerdokumentasjon ────────────────────────────
    _tab_brukermanual = mo.vstack([mo.md(brukermanual_md)])

    # ── Samle i faner ──────────────────────────────────────────
    mo.output.replace(mo.ui.tabs({
        "📊 Simulering & Analyse": _tab_simulering,
        "📁 Dataimport & Versjoner": _tab_import,
        "🔍 Datamodell (Innsyn)": _tab_modell,
        "📜 Endringslogg": _tab_changelog,
        "📖 Brukerdokumentasjon": _tab_brukermanual,
    }))
    return


@app.cell
def _(mo):
    mo.Html('<div class="fti-footer">🌲 FramTre industri — Produksjonskost Simulator · Standardkostkalkyle & simulering</div>')
    return


if __name__ == "__main__":
    app.run()