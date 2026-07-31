"""
lag_testdata_v3.py - Oppretter Produksjonsmodell_Testdata_v3.xlsx med testdata,
kolonnebeskrivelser og statiske dropdowns.

Data leses DYNAMISK fra "Produsjonskalkyle med overflate behandling.xlsx"
saa artikkelnumre, batch/kolonne J, og raavare-referanser alltid er korrekte.

Bruker effective_capacity_pct mekanismen i kostberegning.py for aa handtere statid.

Kjor:
  python lag_testdata_v3.py
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date
import os
import re
import sys


# ── Stildefinisjoner ──────────────────────────────────────────────
HEADER_FONT = Font(bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin'),
)
MAX_ROW = 2000

# ── Kolonnebeskrivelser (vises som kommentarer på header) ─────────
KOLONNER = {
    "Product Master": {
        "Item No": "Unik identifikator for varen (artikkelnr fra ERP).",
        "Description": "Beskrivende navn pa varen.",
        "Item Type": "Type vare: Raw Material, Semi Finished, Finished Good, By Product, Trading Item",
        "Product Group": "Gruppering av varer.",
        "Base Unit of Measure": "Standard maleenhet. Eksempel: LM, M3, KG, LTR",
        "Active": "Er varen aktiv? Ja / Nei",
    },
    "Locations": {
        "Location Code": "Unik kode for lokasjonen. Eksempel: KOD",
        "Location Name": "Navn pa lokasjonen.",
        "Location Type": "Type lokasjon: Factory, Warehouse, Distribution Center, Sales Office",
        "Active": "Er lokasjonen aktiv? Ja / Nei",
    },
    "Work Centers": {
        "Work Center Code": "Unik identifikator for arbeidssenteret.",
        "Description": "Beskrivende navn pa arbeidssenteret",
        "Location Code": "Fabrikken arbeidssenteret tilhorer",
        "Labor Cost per Hour": "Arbeidskostnad per time (lonn, avgift, pensjon, ferie)",
        "Machine Cost per Hour": "Maskinkostnad per time.",
        "Overhead Cost per Hour": "Indirekte produksjonskostnader",
        "Capacity Hours per Day": "Tilgjengelige timer per dag.",
        "Effective Capacity %": "Hvor stor del av tiden som kan brukes til produksjon (statid inkludert)",
        "Active": "Er arbeidssenteret aktivt? Ja / Nei",
    },
    "Operation Master": {
        "Operation Code": "Unik kode for operasjonen.",
        "Description": "Beskrivelse av operasjonen",
        "Default Work Center": "Anbefalt arbeidssenter for operasjonen",
        "Standard Unit": "Maleenhet for produksjonstid. Eksempel: Minutes, Hours",
        "Active": "Er operasjonen aktiv? Ja / Nei",
    },
    "Item Costs": {
        "Item No": "Referanse til varen (Item No fra Product Master)",
        "Cost Type": "Type kostpris: Standard Cost, Last Direct Cost, Forecast Cost, Budget Cost",
        "Unit Cost": "Kostpris per enhet.",
        "Currency": "Valuta. Eksempel: NOK, EUR",
        "Effective Date": "Dato kostprisen gjelder fra",
    },
    "BOM": {
        "Parent Item No": "Produktet som produseres (Item No)",
        "Component Item No": "Komponenten som forbrukes (Item No)",
        "Quantity Per": "Antall output-enheter per input-enhet.",
        "Unit of Measure": "Maleenhet. Eksempel: LM",
        "Scrap %": "Forventet materialsvinn i prosent.",
        "Co-Prod %": "Andel samprodukt (co-product).",
        "Co-Prod Item No": "Varenummer for samproduktet.",
        "Valid From": "Gyldig fra dato",
        "Valid To": "Gyldig til dato (tom = alltid)",
    },
    "Routing": {
        "Item No": "Produkt som produseres (Item No)",
        "Operation No": "Sekvensnummer.",
        "Operation Code": "Operasjon (ref. Operation Master)",
        "Work Center Code": "Arbeidssenter (ref. Work Centers)",
        "Setup Time Minutes": "Klargjoringstid i minutter.",
        "Run Time Minutes": "Produksjonstid per enhet i minutter.",
        "Batch Size": "Normal ordrestorrelse.",
        "Valid From": "Gyldig fra dato",
        "Valid To": "Gyldig til dato (tom = alltid)",
    },
    "By Product Rules": {
        "Parent Item No": "Produktet (ferdigvaren) som skaper biproduktet",
        "By Product Item No": "Biproduktet (Item No).",
        "Expected Quantity": "Forventet mengde biprodukt per enhet ferdigvare",
        "Unit of Measure": "Maleenhet.",
        "Market Value": "Forventet markedspris per enhet.",
        "Allocation Method": "Reduce Main Product Cost, Separate Profit Center, Informational Only",
    },
    "Capacity Calendar": {
        "Work Center": "Arbeidssenter (ref. Work Centers)",
        "Date": "Dato",
        "Available Hours": "Tilgjengelige timer for dagen",
        "Planned Downtime": "Planlagte stopp i timer (vedlikehold, ferie, ombygging)",
        "Available Production Hours": "Timer til produksjon = Available - Planned",
    },
    "Production Scenario": {
        "Scenario Name": "Navn pa scenario.",
        "Product": "Produktet som simuleres (Item No)",
        "Planned Quantity": "Planlagt produksjonsmengde.",
        "Start Date": "Startdato for scenario",
        "End Date": "Sluttdato for scenario",
    },
}

# ── Statiske dropdowns ────────────────────────────────────────────
STATISKE_DROPDOWNS = [
    ("Product Master", "C", '"Raw Material,Semi Finished,Finished Good,By Product,Trading Item"'),
    ("Product Master", "F", '"Ja,Nei"'),
    ("Locations", "C", '"Factory,Warehouse,Distribution Center,Sales Office"'),
    ("Locations", "D", '"Ja,Nei"'),
    ("Work Centers", "I", '"Ja,Nei"'),
    ("Operation Master", "D", '"Minutes,Hours"'),
    ("Operation Master", "E", '"Ja,Nei"'),
    ("Item Costs", "B", '"Standard Cost,Last Direct Cost,Forecast Cost,Budget Cost"'),
    ("Item Costs", "D", '"NOK,EUR,USD,SEK,DKK"'),
    ("By Product Rules", "F", '"Reduce Main Product Cost,Separate Profit Center,Informational Only"'),
]

# ═══════════════════════════════════════════════════════════════════════
#  DATA FRA PRODUSJONSKALKYLE MED OVERFLATE BEHANDLING
# ═══════════════════════════════════════════════════════════════════════

KILDE_FIL = r"C:\Users\EspenGabrielsen\Documents\vareKost\Produsjonskalkyle med overflate behandling.xlsx"

# Råstoffdefinisjoner (fra Inndata-arket)
# (kode, dimensjon, kvalitet, treslag, pris_m3, meter_pr_m3, splitt2, splitt3, splitt4, produktrruppe)
RAASTOFF = [
    # Gran US/V (rad 30-44)
    ("RM001", "38x125 US/V Gran", "", "", 2984, 210.35, "Skrulast"),
    ("RM002", "44x100 US/V Gran", "", "", 2568, 227.27, "Skrulast"),
    ("RM003", "44x125 US/V Gran", "", "", 3360, 181.82, "Skrulast"),
    ("RM004", "44x150 US/V Gran", "", "", 3448, 151.52, "Skrulast"),
    ("RM005", "50x75 US/V Gran", "", "", 2831, 266.67, "Skrulast"),
    ("RM006", "48x100 US/V Gran", "", "", 2568, 208.33, "Skrulast"),
    ("RM007", "50x125 US/V Gran", "", "", 3064, 160.00, "Skrulast"),
    ("RM008", "50x150 US/V Gran", "", "", 3461, 133.33, "Skrulast"),
    ("RM009", "50x175 US/V Gran", "", "", 3390, 114.29, "Skrulast"),
    ("RM010", "50x200 US/V Gran", "", "", 3390, 100.00, "Skrulast"),
    ("RM011", "50x225 US/V Gran", "", "", 3010, 88.89, "Skrulast"),
    ("RM012", "63x150 US/V Gran", "", "", 2600, 105.82, "Skrulast"),
    ("RM013", "63x175 US/V Gran", "", "", 2640, 90.70, "Skrulast"),
    ("RM014", "66x150 US/V Gran", "", "", 3390, 101.10, "Skrulast"),
    ("RM015", "66x123 US/V Furu", "", "", 3000, 121.00, "Skrulast"),
    # Furu K90 (rad 48-54)
    ("RM016", "38x150 K90 Furu", "", "", 4532, 175.00, "Skrulast"),
    ("RM017", "50x100 K90 Furu", "", "", 4539, 200.00, "Skrulast"),
    ("RM018", "50x125 K90 Furu", "", "", 4526, 160.00, "Skrulast"),
    ("RM019", "44x150 K90 Furu", "", "", 4105, 151.52, "Skrulast"),
    ("RM020", "50x150 K90 Furu", "", "", 4156, 133.33, "Skrulast"),
    ("RM021", "50x175 K90 Furu", "", "", 3600, 114.29, "Skrulast"),
    ("RM022", "50x200 K90 Furu", "", "", 3960, 100.00, "Skrulast"),
    # Thermo (rad 58-60)
    ("RM023", "25x150 THERMO", "", "", 5886, 266.00, "Skrulast"),
    ("RM024", "50x150 Thermo", "", "", 5668, 133.00, "Skrulast"),
    ("RM025", "50x100 Thermo", "", "", 5668, 200.00, "Skrulast"),
    # C24 (rad 63-66)
    ("RM026", "50x100 C24", "", "", 2900, 200.00, "Skrulast"),
    ("RM027", "50x125 C24", "", "", 2900, 160.00, "Skrulast"),
    ("RM028", "50x150 C24", "", "", 2900, 133.33, "Skrulast"),
    ("RM029", "50x200 C24", "", "", 2900, 100.00, "Skrulast"),
    # Impregnert
    ("RM030", "66x123 Impregnert", "", "", 3000, 121.00, "Skrulast"),
    # Maling
    ("RM031", "Maling - Opaque", "", "", 42.22, None, "Maling"),
    ("RM032", "Maling - Visir", "", "", 73.10, None, "Maling"),
    ("RM033", "Maling - Power/C.EX", "", "", 62.44, None, "Maling"),
    ("RM034", "Maling - Trebitt", "", "", 69.14, None, "Maling"),
    ("RM035", "Maling - Extreem", "", "", 64.00, None, "Maling"),
    ("RM036", "Jernvitrol", "", "", 47.86, None, "Maling"),
    ("RM037", "Flokuleringsmiddel", "", "", 50.00, None, "Kjemi"),
]

# Inndata-rad -> RM-kode
RAD_TIL_RM = {
    30: "RM001", 31: "RM002", 32: "RM003", 33: "RM004", 34: "RM005",
    35: "RM006", 36: "RM007", 37: "RM008", 38: "RM009", 39: "RM010",
    40: "RM011", 41: "RM012", 42: "RM013", 43: "RM014", 44: "RM015",
    48: "RM016", 49: "RM017", 50: "RM018", 51: "RM019", 52: "RM020",
    53: "RM021", 54: "RM022", 58: "RM023", 59: "RM024", 60: "RM025",
    63: "RM026", 64: "RM027", 65: "RM028", 66: "RM029",
}

def ekstraher_inndata_rad(formel):
    if formel and 'D' in str(formel):
        m = re.search(r'D(\d+)', str(formel))
        if m:
            return int(m.group(1))
    return None

# Hjelpemapping: Inndata-rad -> pris (slik at vi kan matche kolonne V)
RAD_TIL_PRIS = {
    30: 2984, 31: 2568, 32: 3360, 33: 3448, 34: 2831,
    35: 2568, 36: 3064, 37: 3461, 38: 3390, 39: 3390,
    40: 3010, 41: 2600, 42: 2640, 43: 3390, 44: 3000,
    48: 4532, 49: 4539, 50: 4526, 51: 4105, 52: 4156,
    53: 3600, 54: 3960, 58: 5886, 59: 5668, 60: 5668,
    63: 2900, 64: 2900, 65: 2900, 66: 2900,
}
PRIS_TIL_RAD = {v: k for k, v in RAD_TIL_PRIS.items()}

def last_produkter_fra_excel():
    """Les alle produkter dynamisk fra Excel-filen.
    Kolonne V inneholder råvarepris (cached value).
    Vi matcher prisen mot Inndata-priser for å finne hvilket råstoff som brukes.
    """
    wb = openpyxl.load_workbook(KILDE_FIL, data_only=True)
    ws = wb['Prod kalkyle med behandling']

    produkter = []
    for r in range(4, 170):
        c = ws.cell(r, 3).value
        if c is None:
            continue
        # Artikkelnr: konverter float->int for K90-produkter
        a_raw = ws.cell(r, 1).value
        if isinstance(a_raw, float):
            a = int(a_raw)
        else:
            a = a_raw

        e = ws.cell(r, 5).value   # Avdeling
        f = ws.cell(r, 6).value   # Statid
        g = ws.cell(r, 7).value   # 2.sort %
        h = ws.cell(r, 8).value   # Hastighet m/min
        j = ws.cell(r, 10).value  # Batch/storrelse (kolonne J)
        v_pris = ws.cell(r, 22).value  # Råvarepris (kolonne V)
        
        # Match prisen mot Inndata for å finne riktig rad
        raav_rad = PRIS_TIL_RAD.get(v_pris)
        
        produkter.append({
            'artnr': a,
            'beskr': str(c).strip(),
            'avdeling': str(e).strip() if e else None,
            'statid': f,
            'sort2': g,
            'hastighet': h,
            'raavare_rad': raav_rad,
            'batch': j,
        })
    wb.close()
    return produkter

# ── Hjelpefunksjoner ──────────────────────────────────────────────

def finn_produktgruppe(beskrivelse, kode):
    b = beskrivelse.upper()
    if str(kode) in ("JD16073", "JD16073-B"): return "Panel"
    if "TERRASSE" in b: return "Terrasse"
    if "KONSTRUKSJON" in b or "C24" in b: return "Konstruksjon"
    if "THERMO" in b: return "Terrasse"
    if "GULV" in b or "GULVTAVLE" in b or "TREGULV" in b or "LAMINERT" in b: return "Gulv"
    if "VEGGTAVLE" in b: return "Gulv"
    if "PANEL" in b or "DOBBELT" in b or "FASPANEL" in b: return "Panel"
    if "KONGSPANEL" in b or "DRONNINGPANEL" in b: return "Panel"
    if "SVEITSERKL" in b or "KRAGER" in b: return "Panel"
    if "BAROKK" in b or "EMPIRE" in b or "TUNSBERG" in b: return "Kledning"
    if "VANNBRETT" in b or "FENDER" in b or "PERLE" in b: return "Kledning"
    if "RUSTIKK" in b or "FINSAG" in b or "VESTL" in b: return "Kledning"
    if "JUST" in b or "D.FALS" in b: return "Kledning"
    if "RAFT" in b or "HASAS" in b or "HASÅS" in b: return "Kledning"
    if "KLEDN" in b: return "Kledning"
    return "Diverse"

def finn_maling_rm(beskrivelse):
    b = beskrivelse.upper()
    if "VISIR" in b: return "RM032"
    if "OPAQUE" in b or "OPAK" in b: return "RM031"
    if "POWER" in b or "C.EX" in b or "C/EX" in b: return "RM033"
    if "TREBITT" in b: return "RM034"
    if "EXTREEM" in b: return "RM035"
    if "JERNVITROL" in b: return "RM036"
    return None

def legg_til_dropdowns(ws, ark_navn):
    for sheet, col, liste in STATISKE_DROPDOWNS:
        if sheet != ark_navn:
            continue
        dv = DataValidation(type="list", formula1=liste, allow_blank=True,
            showErrorMessage=True, errorTitle="Ugyldig verdi",
            error=f"Verdien ma vare en av: {liste.replace(chr(34), '')}")
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{MAX_ROW}")

def lag_ark(wb, navn, headers):
    ws = wb.create_sheet(navn)
    ws.append(headers)
    desc = KOLONNER.get(navn, {})
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = THIN_BORDER
        if header in desc:
            cell.comment = Comment(desc[header], "System")
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(len(str(headers[col-1])) + 4, 35)
    legg_til_dropdowns(ws, navn)
    return ws

def style_auto(ws, num_rader, num_cols):
    for row in range(2, num_rader + 1):
        for col in range(1, num_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(wrap_text=True)
    # Auto-bredde
    for col in range(1, num_cols + 1):
        max_len = 0
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=True):
            for val in row:
                if val is not None:
                    max_len = max(max_len, len(str(val)))
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(max_len + 3, 40)


# ══════════════════════════════════════════════════════════════════
#  HOVEDPROGRAM
# ══════════════════════════════════════════════════════════════════

wb = openpyxl.Workbook()
ARBEIDSSENTER_KOST = {
    'SPESIALHOVEL': {'labor': 600, 'machine': 700, 'overhead': 305, 'timer': 16, 'eff': 94},
    'HOVEDHOVEL':   {'labor': 1800, 'machine': 2000, 'overhead': 803, 'timer': 16, 'eff': 92},
    'MALINGSLINJE': {'labor': 900, 'machine': 900, 'overhead': 560, 'timer': 16, 'eff': 90},
    'PAKKELINJE':   {'labor': 400, 'machine': 300, 'overhead': 100, 'timer': 8, 'eff': 95},
    'KVHOVEL':      {'labor': 500, 'machine': 200, 'overhead': 100, 'timer': 16, 'eff': 92},
}

AVD_WC = {'Hovedhøvel': 'HOVEDHOVEL', 'Spesialhøvel': 'SPESIALHOVEL'}
DEFAULT_BATCH = 500

# Last produkter
PRODUKTER = last_produkter_fra_excel()

print(f"Lest {len(PRODUKTER)} produkter fra {KILDE_FIL}")
ant_med_batch = sum(1 for p in PRODUKTER if p['batch'] is not None)
print(f"  {ant_med_batch} produkter har batch-verdi")

# ══════════════════════════════════════════════════════════════════
# 1. Product Master
# ══════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Product Master"
headers = ["Item No", "Description", "Item Type", "Product Group", "Base Unit of Measure", "Active"]
ws.append(headers)
desc = KOLONNER["Product Master"]
for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.font = HEADER_FONT; cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
    cell.border = THIN_BORDER
    if header in desc: cell.comment = Comment(desc[header], "System")
for col in range(1, len(headers) + 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(len(headers[col-1]) + 4, 35)
legg_til_dropdowns(ws, "Product Master")

pm_data = []
for rm in RAASTOFF:
    pm_data.append((rm[0], rm[1], "Raw Material", rm[6], "M3", "Ja"))
for p in PRODUKTER:
    grp = finn_produktgruppe(p['beskr'], p['artnr'] or '')
    pm_data.append((p['artnr'], p['beskr'], "Finished Good", grp, "LM", "Ja"))
pm_data.append(("JD16073", "16x73 Just. Kledn. ca. 1250 m/pk", "Finished Good", "Panel", "LM", "Ja"))
pm_data.append(("JD16073-B", "16x73 Just. Kledn. B-vare", "Semi Finished", "Panel", "LM", "Ja"))
pm_data.append(("BP001", "Hovelspon", "By Product", "Spon", "KG", "Ja"))
pm_data.append(("BP002", "Flis", "By Product", "Spon", "KG", "Ja"))
pm_data.append(("BP003", "Bark", "By Product", "Spon", "KG", "Ja"))

num_cols = len(headers)
for row in pm_data:
    ws.append(row)
style_auto(ws, len(pm_data) + 1, num_cols)

# ══════════════════════════════════════════════════════════════════
# 2. Locations
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "Locations", ["Location Code", "Location Name", "Location Type", "Active"])
for row in [("KOD", "Kodal Fabrikk", "Factory", "Ja"),
            ("KV", "Kvaas", "Factory", "Ja"),
            ("SKI", "Skien Lager", "Warehouse", "Ja")]:
    ws.append(row)
style_auto(ws, 4, 4)

# ══════════════════════════════════════════════════════════════════
# 3. Work Centers
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "Work Centers", [
    "Work Center Code", "Description", "Location Code",
    "Labor Cost per Hour", "Machine Cost per Hour", "Overhead Cost per Hour",
    "Capacity Hours per Day", "Effective Capacity %", "Active"])
for code, desc, loc in [("HOVEDHOVEL", "Hovedhovel (HH)", "KOD"),
                          ("SPESIALHOVEL", "Spesialhovel (SH)", "KOD"),
                          ("MALINGSLINJE", "Malingslinje (M1)", "KOD"),
                          ("PAKKELINJE", "Pakkelinje", "KOD"),
                          ("KVHOVEL", "Kvaas hovel", "KV")]:
    k = ARBEIDSSENTER_KOST[code]
    ws.append((code, desc, loc, k['labor'], k['machine'], k['overhead'], k['timer'], k['eff'], "Ja"))
style_auto(ws, 6, 9)

# ══════════════════════════════════════════════════════════════════
# 4. Operation Master
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "Operation Master", ["Operation Code", "Description", "Default Work Center", "Standard Unit", "Active"])
for row in [("HOVLING", "Hovling (oppdeling + hovling + profilering)", "HOVEDHOVEL", "Minutes", "Ja"),
            ("MALING", "Overflatebehandling/maling", "MALINGSLINJE", "Minutes", "Ja"),
            ("PACKING", "Pakking", "PAKKELINJE", "Minutes", "Ja")]:
    ws.append(row)
style_auto(ws, 4, 5)

# ══════════════════════════════════════════════════════════════════
# 5. Item Costs
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "Item Costs", ["Item No", "Cost Type", "Unit Cost", "Currency", "Effective Date"])
ic_data = []
for rm in RAASTOFF:
    ic_data.append((rm[0], "Standard Cost", rm[4], "NOK", date(2026, 1, 1)))
for p in PRODUKTER:
    ic_data.append((p['artnr'], "Standard Cost", 0.0, "NOK", date(2026, 1, 1)))
ic_data.append(("JD16073", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)))
ic_data.append(("JD16073-B", "Standard Cost", 3.00, "NOK", date(2026, 1, 1)))
ic_data.append(("BP001", "Standard Cost", 1.50, "NOK", date(2026, 1, 1)))
ic_data.append(("BP002", "Standard Cost", 0.80, "NOK", date(2026, 1, 1)))
ic_data.append(("BP003", "Standard Cost", 0.50, "NOK", date(2026, 1, 1)))
for row in ic_data:
    ws.append(row)
style_auto(ws, len(ic_data) + 1, 5)

# ══════════════════════════════════════════════════════════════════
# 6. BOM
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "BOM", ["Parent Item No", "Component Item No", "Quantity Per",
    "Unit of Measure", "Scrap %", "Co-Prod %", "Co-Prod Item No", "Valid From", "Valid To"])
bom_data = []

RM_METR_PR_M3 = {rm[0]: rm[5] for rm in RAASTOFF if rm[5] is not None}

for p in PRODUKTER:
    if p['artnr'] is None:
        continue
    rm_kode = RAD_TIL_RM.get(p['raavare_rad'])
    if rm_kode is None:
        continue
    mp3 = RM_METR_PR_M3.get(rm_kode, 1.0)
    scrap_pct = (p['statid'] or 0) * 100
    bom_data.append((p['artnr'], rm_kode, round(mp3, 6), "LM", scrap_pct, 0.0, "",
                     date(2026, 1, 1), date(2026, 12, 31)))
    maling_rm = finn_maling_rm(p['beskr'])
    if maling_rm:
        bom_data.append((p['artnr'], maling_rm, 0.12, "LTR", 2.0, 0.0, "",
                         date(2026, 1, 1), date(2026, 12, 31)))

# JD16073 co-prod B-vare
bom_data.append(("JD16073", "RM005", 266.67, "LM", 0.0, 0.5, "JD16073-B",
                 date(2026, 1, 1), date(2026, 12, 31)))

for row in bom_data:
    ws.append(row)
style_auto(ws, len(bom_data) + 1, 9)

# ══════════════════════════════════════════════════════════════════
# 7. Routing (kun HOVLING + evt. MALING for overflatebehandling)
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "Routing", ["Item No", "Operation No", "Operation Code", "Work Center Code",
    "Setup Time Minutes", "Run Time Minutes", "Batch Size", "Valid From", "Valid To"])
routing_data = []

for p in PRODUKTER:
    if p['artnr'] is None:
        continue
    avd = p['avdeling'] or ""
    wc = AVD_WC.get(avd, "HOVEDHOVEL")
    run_time = (1.0 / p['hastighet']) if p['hastighet'] and p['hastighet'] > 0 else 0.15
    bs = p['batch'] if p['batch'] is not None else DEFAULT_BATCH

    routing_data.append((p['artnr'], 10, "HOVLING", wc, 15.0, round(run_time, 6), bs,
                         date(2026, 1, 1), date(2026, 12, 31)))
    if finn_maling_rm(p['beskr']):
        routing_data.append((p['artnr'], 20, "MALING", "MALINGSLINJE", 20.0, 0.10, bs,
                             date(2026, 1, 1), date(2026, 12, 31)))

# JD16073
routing_data.append(("JD16073", 10, "HOVLING", "SPESIALHOVEL", 30.0, 0.028571, 3200.0,
                      date(2026, 1, 1), date(2026, 12, 31)))

for row in routing_data:
    ws.append(row)
style_auto(ws, len(routing_data) + 1, 9)

# ══════════════════════════════════════════════════════════════════
# 8. By Product Rules
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "By Product Rules", ["Parent Item No", "By Product Item No", "Expected Quantity",
    "Unit of Measure", "Market Value", "Allocation Method"])
byprod_data = []
for p in PRODUKTER:
    if p['artnr'] is not None:
        byprod_data.append((p['artnr'], "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"))
byprod_data.append(("JD16073", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"))
for row in byprod_data:
    ws.append(row)
style_auto(ws, len(byprod_data) + 1, 6)

# ══════════════════════════════════════════════════════════════════
# 9. Capacity Calendar
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "Capacity Calendar", ["Work Center", "Date", "Available Hours",
    "Planned Downtime", "Available Production Hours"])
kal_data = []
for wc in ["HOVEDHOVEL", "MALINGSLINJE", "SPESIALHOVEL", "PAKKELINJE"]:
    timer = ARBEIDSSENTER_KOST.get(wc, {}).get('timer', 16)
    for dag in range(1, 32):
        try:
            d = date(2026, 1, dag)
            if d.weekday() < 5:
                nedetid = 4 if dag == 19 else 0
                kal_data.append((wc, d, timer, nedetid, timer - nedetid))
        except:
            pass
for row in kal_data:
    ws.append(row)
style_auto(ws, len(kal_data) + 1, 5)

# ══════════════════════════════════════════════════════════════════
# 10. Production Scenario
# ══════════════════════════════════════════════════════════════════
ws = lag_ark(wb, "Production Scenario", ["Scenario Name", "Product", "Planned Quantity",
    "Start Date", "End Date"])

# Utvalg av topp-produkter
topp_prod = PRODUKTER[:12]  # Forste 12
scenario_data = []
for p in topp_prod:
    qty = min(50000, (p['batch'] or 500) * 10)
    scenario_data.append(("Normal Produksjon", p['artnr'], qty, date(2026, 1, 1), date(2026, 12, 31)))
scenario_data.append(("Normal Produksjon", "JD16073", 3200, date(2026, 1, 1), date(2026, 12, 31)))

# Full kapasitet (2x for de mest populære)
if len(PRODUKTER) > 0:
    p = PRODUKTER[0]
    scenario_data.append(("Full Kapasitet", p['artnr'], 80000, date(2026, 1, 1), date(2026, 12, 31)))
if len(PRODUKTER) > 6:
    p = PRODUKTER[6]
    scenario_data.append(("Full Kapasitet", p['artnr'], 60000, date(2026, 1, 1), date(2026, 12, 31)))

for row in scenario_data:
    ws.append(row)
style_auto(ws, len(scenario_data) + 1, 5)

# ── Lagre ────────────────────────────────────────────────────────
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]

# Lagre til src/ (der kostberegning.py forventer filen)
_SRC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_FILE = os.path.join(_SRC_DIR, "Produksjonsmodell_Testdata_v3.xlsx")
wb.save(OUTPUT_FILE)
print(f"OK: Testdata lagret til {OUTPUT_FILE}")
print(f"   {len(wb.sheetnames)} ark:")
for i, name in enumerate(wb.sheetnames, 1):
    ws = wb[name]
    antall = ws.max_row - 1 if ws.max_row > 1 else 0
    antall_val = len(ws.data_validations.dataValidation) if ws.data_validations else 0
    print(f"   {i}. {name}: {antall} rader, {antall_val} dropdowns")