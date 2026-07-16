"""
lag_testdata_v3.py - Oppretter Produksjonsmodell_Testdata_v3.xlsx med testdata OG beskrivelser (kommentarer) fra malen

Kjor:
  python lag_testdata_v3.py

Dette lager en Excel-fil med testdata som kan brukes med kostberegning.py
og inkluderer beskrivende kommentarer pa header-radene fra Produksjonsmodell_Mal.xlsx
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import date

# ── Stildefinisjoner ──────────────────────────────────────────────
HEADER_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT_WHITE = Font(bold=True, size=11, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin'),
)


def style_header(ws, num_cols):
    """Styl overskriftraden."""
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = THIN_BORDER


def style_data(ws, start_row, end_row, num_cols):
    """Styl datarader."""
    for row in range(start_row, end_row + 1):
        for col in range(1, num_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(wrap_text=True)


def auto_width(ws, num_cols):
    """Auto-juster kolonnebredde."""
    for col in range(1, num_cols + 1):
        max_len = 0
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=True):
            for cell_val in row:
                if cell_val is not None:
                    max_len = max(max_len, len(str(cell_val)))
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(max_len + 3, 30)


def add_descriptions(wb_dest):
    """Legg til beskrivende kommentarer pa header-radene basert pa ForslagExcelModel.md."""
    from openpyxl.comments import Comment

    descriptions = {
        "Product Master": {
            "Item No": "Unik identifikator for varen. Eksempel: RM001, FG001, BP001",
            "Description": "Beskrivende navn pa varen. Eksempel: Skrulast 48x198, Utvendig Panel 21x95",
            "Item Type": "Type vare: Raw Material, Semi Finished, Finished Good, By Product, Trading Item",
            "Product Group": "Gruppering av varer. Eksempel: Skrulast, Panel, Kledning, Spon",
            "Base Unit of Measure": "Standard maleenhet. Eksempel: LM, M3, KG, PCS",
            "Active": "Angir om varen er aktiv. Verdier: Ja / Nei",
        },
        "Locations": {
            "Location Code": "Unik kode for lokasjonen. Eksempel: KOD",
            "Location Name": "Navn pa lokasjonen. Eksempel: Kodal Fabrikk",
            "Location Type": "Type lokasjon: Factory, Warehouse, Distribution Center, Sales Office",
            "Active": "Angir om lokasjonen er aktiv. Verdier: Ja / Nei",
        },
        "Work Centers": {
            "Work Center Code": "Unik identifikator for arbeidssenteret. Eksempel: HOVEDHOVEL",
            "Description": "Beskrivende navn pa arbeidssenteret",
            "Location Code": "Fabrikken arbeidssenteret tilhorer",
            "Labor Cost per Hour": "Arbeidskostnad per time (lonn, arbeidsgiveravgift, pensjon, feriepenger). Eksempel: 550 NOK/time",
            "Machine Cost per Hour": "Maskinkostnad per time (avskrivninger, service, leasing, vedlikehold, energi). Eksempel: 900 NOK/time",
            "Overhead Cost per Hour": "Indirekte produksjonskostnader (produksjonsledelse, kvalitet, vedlikeholdsadm., intern logistikk). Eksempel: 150 NOK/time",
            "Capacity Hours per Day": "Tilgjengelige timer per dag. Eksempel: 16 timer",
            "Effective Capacity %": "Hvor stor del av tilgjengelig tid som faktisk kan brukes til produksjon. Tar hensyn til stopp, vedlikehold, feil, justeringer, omskiftninger. Eksempel: 85 %",
            "Active": "Angir om arbeidssenteret er aktivt. Verdier: Ja / Nei",
        },
        "Operation Master": {
            "Operation Code": "Unik operasjonskode. Eksempel: HOVLING, MALING, PACKING",
            "Description": "Beskrivelse av operasjonen. Eksempel: Hovling (oppdeling + høvling + profilering), Maling, Pakking",
            "Default Work Center": "Anbefalt arbeidssenter for operasjonen",
            "Standard Unit": "Maleenhet for produksjonstid. Eksempel: Minutes, Hours",
            "Active": "Angir om operasjonen er aktiv. Verdier: Ja / Nei",
        },
        "Item Costs": {
            "Item No": "Referanse til varen (Item No fra Product Master)",
            "Cost Type": "Type kostpris: Standard Cost, Last Direct Cost, Forecast Cost, Budget Cost",
            "Unit Cost": "Kostpris per enhet. Eksempel: 12,00 NOK per lm",
            "Currency": "Valuta. Eksempel: NOK, EUR",
            "Effective Date": "Dato kostprisen gjelder fra",
        },
        "BOM": {
            "Parent Item No": "Produktet som produseres (Item No)",
            "Component Item No": "Komponenten som forbrukes (Item No)",
            "Quantity Per": "Antall output-enheter per input-enhet. Forbruk per output = 1 / Quantity Per. Eksempel: 400 LM per M3",
            "Unit of Measure": "Maleenhet for forholdet",
            "Scrap %": "Forventet materialsvinn i prosent. Eksempel: 5 %",
            "Co-Prod %": "Andel av produksjonen som blir samprodukt (co-product). Eksempel: 6% B-vare",
            "Co-Prod Item No": "Varenummer for samproduktet (co-product). Eksempel: JD16073-B",
            "Valid From": "Gyldig fra dato",
            "Valid To": "Gyldig til dato",
        },
        "Routing": {
            "Item No": "Produkt som produseres (Item No)",
            "Operation No": "Sekvensnummer for operasjonen. Eksempel: 10, 20, 30",
            "Operation Code": "Hvilken operasjon som utfores (ref. Operation Master)",
            "Work Center Code": "Arbeidssenter som utforer operasjonen (ref. Work Centers)",
            "Setup Time Minutes": "Tid brukt til klargjoring (omstilling, knivbytte, innkjoring, kontrollmaling). Eksempel: 15 minutter",
            "Run Time Minutes": "Produksjonstid per enhet. Eksempel: 0,15 minutter per lm",
            "Batch Size": "Normal ordrestorrelse. Brukes til a fordele setupkostnad. Eksempel: 500 lm",
            "Changeover Time Minutes": "Tid det tar a stille om FRA dette produktet til et annet pa samme maskin. Brukes i scenario-simulering med flere produkter. Eksempel: 45 minutter",
            "Valid From": "Gyldig fra dato",
            "Valid To": "Gyldig til dato",
        },

        "By Product Rules": {
            "Parent Item No": "Produktet (ferdigvaren) som skaper biproduktet",
            "By Product Item No": "Biproduktet (Item No)",
            "Expected Quantity": "Forventet mengde biprodukt per enhet hovedprodukt",
            "Unit of Measure": "Maleenhet for biproduktet",
            "Market Value": "Forventet markedspris per enhet. Eksempel: 1,50 NOK/kg",
            "Allocation Method": "Hvordan verdien skal handteres: Reduce Main Product Cost, Separate Profit Center, Informational Only",
        },
        "Capacity Calendar": {
            "Work Center": "Arbeidssenter (ref. Work Centers)",
            "Date": "Dato for kapasitetsplanen",
            "Available Hours": "Tilgjengelige timer denne dagen",
            "Planned Downtime": "Planlagte stopp (vedlikehold, ferie, ombygging)",
            "Available Production Hours": "Timer som faktisk er tilgjengelige for produksjon = Available Hours - Planned Downtime",
        },
        "Production Scenario": {
            "Scenario Name": "Navn pa scenario. Eksempel: Budsjett 2027, Full Kapasitet, Normal Produksjon",
            "Product": "Produktet som simuleres (Item No)",
            "Planned Quantity": "Planlagt produksjonsmengde. Eksempel: 100 000 lm",
            "Start Date": "Startdato for scenarioet",
            "End Date": "Sluttdato for scenarioet",
        },
    }

    for sheet_name, cols in descriptions.items():
        if sheet_name in wb_dest.sheetnames:
            ws = wb_dest[sheet_name]
            for col_idx in range(1, ws.max_column + 1):
                header_cell = ws.cell(row=1, column=col_idx)
                if header_cell.value and header_cell.value in cols:
                    comment_text = cols[header_cell.value]
                    comment = Comment(comment_text, "System")
                    comment.width = 300
                    comment.height = 100
                    header_cell.comment = comment



# ── Opprett workbook ──────────────────────────────────────────────
wb = openpyxl.Workbook()

# ══════════════════════════════════════════════════════════════════
# 1. Product Master
# ══════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Product Master"
headers = ["Item No", "Description", "Item Type", "Product Group", "Base Unit of Measure", "Active"]
ws.append(headers)
data = [
    # ── Råvarer ────────────────────────────────────────────────────
    # Gran US/V (fra kalkylen Inndata: 2568-3461 kr/M3)
    ("RM001", "Gran 50x200 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM002", "Gran 44x150 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM003", "Gran 50x125 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM004", "Gran 50x100 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM005", "Gran 44x100 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM006", "Gran 44x125 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM007", "Gran 50x75 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM008", "Gran 50x150 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM009", "Gran 50x175 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM010", "Gran 63x150 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM011", "Gran 63x175 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM012", "Gran 38x125 US/V", "Raw Material", "Skrulast", "M3", "Ja"),
    # Furu K90 (fra kalkylen: 3600-4539 kr/M3)
    ("RM013", "Furu 50x100 K90", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM014", "Furu 44x150 K90", "Raw Material", "Skrulast", "M3", "Ja"),
    # Thermo Furu (fra kalkylen: 5668-5886 kr/M3)
    ("RM015", "Thermo Furu 25x150", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM016", "Thermo Furu 50x150", "Raw Material", "Skrulast", "M3", "Ja"),
    # C24-konstruksjon (fra kalkylen: 2900 kr/M3)
    ("RM017", "Gran C24 50x200", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM018", "Gran C24 50x150", "Raw Material", "Skrulast", "M3", "Ja"),
    ("RM019", "Gran C24 50x100", "Raw Material", "Skrulast", "M3", "Ja"),
    # Impregnert (fra kalkylen: 3000 kr/M3 for 66x125)
    ("RM020", "Impregnert 66x125", "Raw Material", "Skrulast", "M3", "Ja"),
    # Maling (fra kalkylen: 42-73 kr/L)
    ("RM021", "Maling - Opaque", "Raw Material", "Maling", "LTR", "Ja"),
    ("RM022", "Maling - Visir", "Raw Material", "Maling", "LTR", "Ja"),
    ("RM023", "Maling - Power/C.EX", "Raw Material", "Maling", "LTR", "Ja"),
    ("RM024", "Maling - Trebitt", "Raw Material", "Maling", "LTR", "Ja"),
    ("RM025", "Maling - Extreem", "Raw Material", "Maling", "LTR", "Ja"),
    ("RM026", "Jernvitrol", "Raw Material", "Maling", "LTR", "Ja"),
    # Flokuleringsmiddel (fra kalkylen)
    ("RM027", "Flokuleringsmiddel", "Raw Material", "Kjemi", "KG", "Ja"),

    # ── Ferdigvarer ────────────────────────────────────────────────
    # Panel 21x148 ubehandlet - JD19148 fra kalkylen
    # Råstoff: 44x150 US/V, 303,04 lm/M3, kost 12,41 kr/m
    ("FG001", "Utvendig Panel 21x148", "Finished Good", "Panel", "LM", "Ja"),
    # Panel 21x148 med Opaque - kost 14,82 kr/m (12,41 + 1,07 + flokulering)
    ("FG002", "Utvendig Panel 21x148 - Opaque", "Finished Good", "Panel", "LM", "Ja"),
    # Panel 21x148 med Visir - kost 15,32 kr/m (12,41 + 1,57 + flokulering)
    ("FG003", "Utvendig Panel 21x148 - Visir", "Finished Good", "Panel", "LM", "Ja"),
    # Panel 21x148 med Power/C.EX - kost 15,52 kr/m
    ("FG004", "Utvendig Panel 21x148 - Power/C.EX", "Finished Good", "Panel", "LM", "Ja"),
    # Panel 21x148 med Jernvitrol - kost 14,66 kr/m
    ("FG005", "Utvendig Panel 21x148 - Jernvitrol", "Finished Good", "Panel", "LM", "Ja"),
    # Kledning 19x123 ubehandlet - JD19123 fra kalkylen
    # Råstoff: 44x125 US/V, 363,64 lm/M3, kost 10,26 kr/m
    ("FG006", "Kledning 19x123", "Finished Good", "Kledning", "LM", "Ja"),
    # Kledning 19x123 med Opaque - kost 12,54 kr/m
    ("FG007", "Kledning 19x123 - Opaque", "Finished Good", "Kledning", "LM", "Ja"),
    # Kledning 19x123 med Visir - kost 12,98 kr/m
    ("FG008", "Kledning 19x123 - Visir", "Finished Good", "Kledning", "LM", "Ja"),
    # Kledning 19x123 med Trebitt - kost 13,21 kr/m
    ("FG009", "Kledning 19x123 - Trebitt", "Finished Good", "Kledning", "LM", "Ja"),
    # Kledning 22x123 ubehandlet - JD22123 fra kalkylen
    # Råstoff: 50x125 US/V, 320 lm/M3, kost 10,69 kr/m
    ("FG010", "Kledning 22x123", "Finished Good", "Kledning", "LM", "Ja"),
    # Kledning 22x123 med Opaque - kost 12,97 kr/m
    ("FG011", "Kledning 22x123 - Opaque", "Finished Good", "Kledning", "LM", "Ja"),
    # Kledning 22x123 med Visir - kost 13,41 kr/m
    ("FG012", "Kledning 22x123 - Visir", "Finished Good", "Kledning", "LM", "Ja"),
    # Terrassebord 28x120 impregnert - IG28120 fra kalkylen
    # Råstoff: 66x125 Imp, 242 lm/M3, kost 13,35 kr/m
    ("FG013", "Terrassebord 28x120 Impregnert", "Finished Good", "Terrasse", "LM", "Ja"),
    # Konstruksjon 48x198 C24 - JA48198G fra kalkylen
    # Råstoff: 50x200 C24, 100 lm/M3, kost 31,79 kr/m
    ("FG014", "Konstruksjon 48x198 C24", "Finished Good", "Konstruksjon", "LM", "Ja"),
    # Konstruksjon 48x148 C24 - JA48148G fra kalkylen
    # Råstoff: 50x150 C24, 133,33 lm/M3, kost 23,82 kr/m
    ("FG015", "Konstruksjon 48x148 C24", "Finished Good", "Konstruksjon", "LM", "Ja"),
    # Panel 19x148 dobbeltfals - JK19148 fra kalkylen
    # Råstoff: 44x150 US/V, 303,04 lm/M3, kost 12,36 kr/m
    ("FG016", "Panel 19x148 Dobbeltfals", "Finished Good", "Panel", "LM", "Ja"),
    # Panel 19x148 dobbeltfals med Opaque - kost 14,77 kr/m
    ("FG017", "Panel 19x148 Dobbeltfals - Opaque", "Finished Good", "Panel", "LM", "Ja"),
    # Panel 19x148 dobbeltfals med Visir - kost 15,26 kr/m
    ("FG018", "Panel 19x148 Dobbeltfals - Visir", "Finished Good", "Panel", "LM", "Ja"),
    # Kledning 29x148 justert - JD29148 fra kalkylen
    # Råstoff: 63x150 US/V, 211,64 lm/M3, kost 13,23 kr/m
    ("FG019", "Kledning 29x148", "Finished Good", "Kledning", "LM", "Ja"),
    # Kledning 29x148 med Opaque - kost 15,64 kr/m
    ("FG020", "Kledning 29x148 - Opaque", "Finished Good", "Kledning", "LM", "Ja"),

    # ── Vare JD16073 med 2. sortering (A- og B-vare) ─────────────────
    # A-vare: 16x73 JUST. KLEDN. (94% andel av produksjonen)
    ("JD16073", "16x73  JUST. KLEDN.         CA. 1250 M/PK", "Finished Good", "Panel", "LM", "Ja"),
    # B-vare: 16x73 JUST. KLEDN. B-vare (6% andel av produksjonen)
    # B-varen er et samprodukt (co-product), ikke et biprodukt
    ("JD16073-B", "16x73  JUST. KLEDN. B-vare", "Semi Finished", "Panel", "LM", "Ja"),

    # ── Biprodukter ────────────────────────────────────────────────
    ("BP001", "Hovelspon", "By Product", "Spon", "KG", "Ja"),
    ("BP002", "Flis", "By Product", "Spon", "KG", "Ja"),
    ("BP003", "Bark", "By Product", "Spon", "KG", "Ja"),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 2. Locations
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Locations")
headers = ["Location Code", "Location Name", "Location Type", "Active"]
ws.append(headers)
data = [
    ("KOD", "Kodal Fabrikk", "Factory", "Ja"),
    ("KV", "Kvås", "Factory", "Ja"),
    ("SKI", "Skien Lager", "Warehouse", "Ja"),
]
for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 3. Work Centers
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Work Centers")
headers = [
    "Work Center Code", "Description", "Location Code",
    "Labor Cost per Hour", "Machine Cost per Hour", "Overhead Cost per Hour",
    "Capacity Hours per Day", "Effective Capacity %", "Active",
]
ws.append(headers)
data = [
    ("HOVEDHOVEL", "Hovedhovel", "KOD", 550, 900, 150, 16, 85, "Ja"),
    ("SPESIALHOVEL", "Spesialhovel", "KOD", 550, 950, 150, 16, 85, "Ja"),
    ("MALINGSLINJE", "Malingslinje", "KOD", 500, 400, 120, 16, 80, "Ja"),
    ("PAKKELINJE", "Pakkelinje", "KOD", 450, 300, 100, 8, 90, "Ja"),
    ("KVHOVEL", "Kvås høvel", "KV", 500, 200, 100, 16, 85, "Ja"),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 4. Operation Master
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Operation Master")
headers = ["Operation Code", "Description", "Default Work Center", "Standard Unit", "Active"]
ws.append(headers)
data = [
    ("HOVLING", "Hovling (oppdeling + høvling + profilering)", "HOVEDHOVEL", "Minutes", "Ja"),
    ("MALING", "Maling", "MALINGSLINJE", "Minutes", "Ja"),
    ("PACKING", "Pakking", "PAKKELINJE", "Minutes", "Ja"),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 5. Item Costs
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Item Costs")
headers = ["Item No", "Cost Type", "Unit Cost", "Currency", "Effective Date"]
ws.append(headers)
data = [
    # ── Råvarepriser (fra kalkylen "Inndata" sheet) ───────────────
    # Gran US/V
    ("RM001", "Standard Cost", 3390.00, "NOK", date(2026, 1, 1)),  # 50x200 US/V = 3390 kr/M3
    ("RM002", "Standard Cost", 3448.00, "NOK", date(2026, 1, 1)),  # 44x150 US/V = 3448 kr/M3
    ("RM003", "Standard Cost", 3064.00, "NOK", date(2026, 1, 1)),  # 50x125 US/V = 3064 kr/M3
    ("RM004", "Standard Cost", 2568.00, "NOK", date(2026, 1, 1)),  # 50x100 US/V = 2568 kr/M3
    ("RM005", "Standard Cost", 2568.00, "NOK", date(2026, 1, 1)),  # 44x100 US/V = 2568 kr/M3
    ("RM006", "Standard Cost", 3360.00, "NOK", date(2026, 1, 1)),  # 44x125 US/V = 3360 kr/M3
    ("RM007", "Standard Cost", 2831.00, "NOK", date(2026, 1, 1)),  # 50x75 US/V = 2831 kr/M3
    ("RM008", "Standard Cost", 3461.00, "NOK", date(2026, 1, 1)),  # 50x150 US/V = 3461 kr/M3
    ("RM009", "Standard Cost", 3390.00, "NOK", date(2026, 1, 1)),  # 50x175 US/V = 3390 kr/M3
    ("RM010", "Standard Cost", 2600.00, "NOK", date(2026, 1, 1)),  # 63x150 US/V = 2600 kr/M3
    ("RM011", "Standard Cost", 2640.00, "NOK", date(2026, 1, 1)),  # 63x175 US/V = 2640 kr/M3
    ("RM012", "Standard Cost", 2984.00, "NOK", date(2026, 1, 1)),  # 38x125 US/V = 2984 kr/M3
    # Furu K90
    ("RM013", "Standard Cost", 4539.00, "NOK", date(2026, 1, 1)),  # 50x100 K90 = 4539 kr/M3
    ("RM014", "Standard Cost", 4105.00, "NOK", date(2026, 1, 1)),  # 44x150 K90 = 4105 kr/M3
    # Thermo Furu
    ("RM015", "Standard Cost", 5886.00, "NOK", date(2026, 1, 1)),  # 25x150 Thermo = 5886 kr/M3
    ("RM016", "Standard Cost", 5668.00, "NOK", date(2026, 1, 1)),  # 50x150 Thermo = 5668 kr/M3
    # C24
    ("RM017", "Standard Cost", 2900.00, "NOK", date(2026, 1, 1)),  # 50x200 C24 = 2900 kr/M3
    ("RM018", "Standard Cost", 2900.00, "NOK", date(2026, 1, 1)),  # 50x150 C24 = 2900 kr/M3
    ("RM019", "Standard Cost", 2900.00, "NOK", date(2026, 1, 1)),  # 50x100 C24 = 2900 kr/M3
    # Impregnert
    ("RM020", "Standard Cost", 3000.00, "NOK", date(2026, 1, 1)),  # 66x125 Imp = 3000 kr/M3
    # Maling (fra kalkylen: kr pr L)
    ("RM021", "Standard Cost", 42.22, "NOK", date(2026, 1, 1)),    # Opaque = 42,22 kr/L
    ("RM022", "Standard Cost", 73.10, "NOK", date(2026, 1, 1)),    # Visir = 73,10 kr/L
    ("RM023", "Standard Cost", 62.44, "NOK", date(2026, 1, 1)),    # Power/C.EX = 62,44 kr/L
    ("RM024", "Standard Cost", 69.14, "NOK", date(2026, 1, 1)),    # Trebitt = 69,14 kr/L
    ("RM025", "Standard Cost", 64.00, "NOK", date(2026, 1, 1)),    # Extreem = 64,00 kr/L
    ("RM026", "Standard Cost", 47.86, "NOK", date(2026, 1, 1)),    # Jernvitrol = 47,86 kr/L
    ("RM027", "Standard Cost", 50.00, "NOK", date(2026, 1, 1)),    # Flokuleringsmiddel

    # ── Ferdigvarer (0 = beregnes dynamisk) ───────────────────────
    ("FG001", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 21x148 ubehandlet
    ("FG002", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 21x148 Opaque
    ("FG003", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 21x148 Visir
    ("FG004", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 21x148 Power/C.EX
    ("FG005", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 21x148 Jernvitrol
    ("FG006", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 19x123 ubehandlet
    ("FG007", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 19x123 Opaque
    ("FG008", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 19x123 Visir
    ("FG009", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 19x123 Trebitt
    ("FG010", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 22x123 ubehandlet
    ("FG011", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 22x123 Opaque
    ("FG012", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 22x123 Visir
    ("FG013", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Terrassebord 28x120 Impregnert
    ("FG014", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Konstruksjon 48x198 C24
    ("FG015", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Konstruksjon 48x148 C24
    ("FG016", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 19x148 Dobbeltfals
    ("FG017", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 19x148 Dobbeltfals Opaque
    ("FG018", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Panel 19x148 Dobbeltfals Visir
    ("FG019", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 29x148
    ("FG020", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),  # Kledning 29x148 Opaque

    # ── Vare JD16073 med B-vare (0.0 = beregnes, B-vare = 3.00 kr) ──
    ("JD16073", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),
    ("JD16073-B", "Standard Cost", 3.00, "NOK", date(2026, 1, 1)),

    # ── Biprodukter ───────────────────────────────────────────────
    ("BP001", "Standard Cost", 1.50, "NOK", date(2026, 1, 1)),  # Hovelspon
    ("BP002", "Standard Cost", 0.80, "NOK", date(2026, 1, 1)),  # Flis
    ("BP003", "Standard Cost", 0.50, "NOK", date(2026, 1, 1)),  # Bark
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 6. BOM
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("BOM")
headers = ["Parent Item No", "Component Item No", "Quantity Per", "Unit of Measure", "Scrap %", "Co-Prod %", "Co-Prod Item No", "Valid From", "Valid To"]
ws.append(headers)
data = [
    # ── Panel 21x148 (ubehandlet) ← 44x150 US/V ──────────────────
    # 303,04 lm/M3, Quantity Per = 303.04
    ("FG001", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 21x148 (Opaque) ← 44x150 US/V + Opaque maling ─────
    ("FG002", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG002", "RM021", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),  # 0,12 L/m
    # ── Panel 21x148 (Visir) ─────────────────────────────────────
    ("FG003", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG003", "RM022", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 21x148 (Power/C.EX) ────────────────────────────────
    ("FG004", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG004", "RM023", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 21x148 (Jernvitrol) ────────────────────────────────
    ("FG005", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG005", "RM026", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (ubehandlet) ← 44x125 US/V ──────────────
    # 363,64 lm/M3
    ("FG006", "RM006", 363.64, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (Opaque) ─────────────────────────────────
    ("FG007", "RM006", 363.64, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG007", "RM021", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (Visir) ──────────────────────────────────
    ("FG008", "RM006", 363.64, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG008", "RM022", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (Trebitt) ────────────────────────────────
    ("FG009", "RM006", 363.64, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG009", "RM024", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 22x123 (ubehandlet) ← 50x125 US/V ──────────────
    # 320 lm/M3
    ("FG010", "RM003", 320.0, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 22x123 (Opaque) ─────────────────────────────────
    ("FG011", "RM003", 320.0, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG011", "RM021", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 22x123 (Visir) ──────────────────────────────────
    ("FG012", "RM003", 320.0, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG012", "RM022", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Terrassebord 28x120 Impregnert ← 66x125 Imp ─────────────
    # 242 lm/M3
    ("FG013", "RM020", 242.0, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Konstruksjon 48x198 C24 ← 50x200 C24 ────────────────────
    # 100 lm/M3
    ("FG014", "RM017", 100.0, "LM", 3.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Konstruksjon 48x148 C24 ← 50x150 C24 ────────────────────
    # 133,33 lm/M3
    ("FG015", "RM018", 133.33, "LM", 3.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 19x148 Dobbeltfals (ubehandlet) ← 44x150 US/V ─────
    # 303,04 lm/M3
    ("FG016", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 19x148 Dobbeltfals (Opaque) ────────────────────────
    ("FG017", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG017", "RM021", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 19x148 Dobbeltfals (Visir) ─────────────────────────
    ("FG018", "RM002", 303.04, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG018", "RM022", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 29x148 (ubehandlet) ← 63x150 US/V ──────────────
    # 211,64 lm/M3
    ("FG019", "RM010", 211.64, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 29x148 (Opaque) ─────────────────────────────────
    ("FG020", "RM010", 211.64, "LM", 5.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),
    ("FG020", "RM021", 0.12, "LTR", 2.0, 0.0, "", date(2026, 1, 1), date(2026, 12, 31)),

    # ── Vare JD16073 (A-vare) ← RM001 (Gran 50x200 US/V) ──────────
    # Teoretisk utbytte er 533.34 meter per m3.
    # Co-Prod: 0.5% blir B-vare (JD16073-B)
    ("JD16073", "RM001", 533.34, "LM", 0.0, 0.5, "JD16073-B", date(2026, 1, 1), date(2026, 12, 31)),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 7. Routing
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Routing")
headers = ["Item No", "Operation No", "Operation Code", "Work Center Code", "Setup Time Minutes", "Run Time Minutes", "Batch Size", "Valid From", "Valid To"]
ws.append(headers)
data = [
    # ── Panel 21x148 (ubehandlet) - produseres på KOD (HOVEDHOVEL) eller KV (KVHOVEL) ──
    ("FG001", 10, "HOVLING", "HOVEDHOVEL", 15, 0.03, 50000, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG001", 10, "HOVLING", "KVHOVEL", 15, 0.03, 50000, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 21x148 (Opaque) ────────────────────────────────────
    ("FG002", 10, "HOVLING", "HOVEDHOVEL", 15, 0.15, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG002", 20, "MALING", "MALINGSLINJE", 20, 0.10, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG002", 30, "PACKING", "PAKKELINJE", 5, 0.05, 500, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 21x148 (Visir) ─────────────────────────────────────
    ("FG003", 10, "HOVLING", "HOVEDHOVEL", 15, 0.15, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG003", 20, "MALING", "MALINGSLINJE", 20, 0.10, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG003", 30, "PACKING", "PAKKELINJE", 5, 0.05, 500, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 21x148 (Power/C.EX) ────────────────────────────────
    ("FG004", 10, "HOVLING", "HOVEDHOVEL", 15, 0.15, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG004", 20, "MALING", "MALINGSLINJE", 20, 0.10, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG004", 30, "PACKING", "PAKKELINJE", 5, 0.05, 500, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 21x148 (Jernvitrol) ────────────────────────────────
    ("FG005", 10, "HOVLING", "HOVEDHOVEL", 15, 0.15, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG005", 20, "MALING", "MALINGSLINJE", 20, 0.10, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG005", 30, "PACKING", "PAKKELINJE", 5, 0.05, 500, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (ubehandlet) ─────────────────────────────
    ("FG006", 10, "HOVLING", "HOVEDHOVEL", 15, 0.12, 600, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG006", 20, "PACKING", "PAKKELINJE", 5, 0.04, 600, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (Opaque) ─────────────────────────────────
    ("FG007", 10, "HOVLING", "HOVEDHOVEL", 15, 0.12, 600, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG007", 20, "MALING", "MALINGSLINJE", 20, 0.10, 600, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG007", 30, "PACKING", "PAKKELINJE", 5, 0.04, 600, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (Visir) ──────────────────────────────────
    ("FG008", 10, "HOVLING", "HOVEDHOVEL", 15, 0.12, 600, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG008", 20, "MALING", "MALINGSLINJE", 20, 0.10, 600, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG008", 30, "PACKING", "PAKKELINJE", 5, 0.04, 600, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 19x123 (Trebitt) ────────────────────────────────
    ("FG009", 10, "HOVLING", "HOVEDHOVEL", 15, 0.12, 600, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG009", 20, "MALING", "MALINGSLINJE", 20, 0.10, 600, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG009", 30, "PACKING", "PAKKELINJE", 5, 0.04, 600, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 22x123 (ubehandlet) ─────────────────────────────
    ("FG010", 10, "HOVLING", "HOVEDHOVEL", 15, 0.13, 550, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG010", 20, "PACKING", "PAKKELINJE", 5, 0.04, 550, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 22x123 (Opaque) ─────────────────────────────────
    ("FG011", 10, "HOVLING", "HOVEDHOVEL", 15, 0.13, 550, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG011", 20, "MALING", "MALINGSLINJE", 20, 0.10, 550, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG011", 30, "PACKING", "PAKKELINJE", 5, 0.04, 550, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 22x123 (Visir) ──────────────────────────────────
    ("FG012", 10, "HOVLING", "HOVEDHOVEL", 15, 0.13, 550, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG012", 20, "MALING", "MALINGSLINJE", 20, 0.10, 550, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG012", 30, "PACKING", "PAKKELINJE", 5, 0.04, 550, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Terrassebord 28x120 Impregnert ───────────────────────────
    ("FG013", 10, "HOVLING", "SPESIALHOVEL", 20, 0.18, 400, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG013", 20, "PACKING", "PAKKELINJE", 5, 0.05, 400, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Konstruksjon 48x198 C24 ──────────────────────────────────
    ("FG014", 10, "HOVLING", "HOVEDHOVEL", 10, 0.10, 300, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG014", 20, "PACKING", "PAKKELINJE", 5, 0.03, 300, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Konstruksjon 48x148 C24 ──────────────────────────────────
    ("FG015", 10, "HOVLING", "HOVEDHOVEL", 10, 0.10, 300, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG015", 20, "PACKING", "PAKKELINJE", 5, 0.03, 300, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 19x148 Dobbeltfals (ubehandlet) ────────────────────
    ("FG016", 10, "HOVLING", "HOVEDHOVEL", 15, 0.15, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG016", 20, "PACKING", "PAKKELINJE", 5, 0.05, 500, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 19x148 Dobbeltfals (Opaque) ────────────────────────
    ("FG017", 10, "HOVLING", "HOVEDHOVEL", 15, 0.15, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG017", 20, "MALING", "MALINGSLINJE", 20, 0.10, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG017", 30, "PACKING", "PAKKELINJE", 5, 0.05, 500, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Panel 19x148 Dobbeltfals (Visir) ─────────────────────────
    ("FG018", 10, "HOVLING", "HOVEDHOVEL", 15, 0.15, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG018", 20, "MALING", "MALINGSLINJE", 20, 0.10, 500, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG018", 30, "PACKING", "PAKKELINJE", 5, 0.05, 500, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 29x148 (ubehandlet) ─────────────────────────────
    ("FG019", 10, "HOVLING", "SPESIALHOVEL", 15, 0.16, 450, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG019", 20, "PACKING", "PAKKELINJE", 5, 0.05, 450, date(2026, 1, 1), date(2026, 12, 31)),
    # ── Kledning 29x148 (Opaque) ─────────────────────────────────
    ("FG020", 10, "HOVLING", "SPESIALHOVEL", 15, 0.16, 450, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG020", 20, "MALING", "MALINGSLINJE", 20, 0.10, 450, date(2026, 1, 1), date(2026, 12, 31)),
    ("FG020", 30, "PACKING", "PAKKELINJE", 5, 0.05, 450, date(2026, 1, 1), date(2026, 12, 31)),

    # ── Vare JD16073 ─────────────────────────────────────────────
    # Kjører på SPESIALHOVEL. Hastighet 35 m/min. Setup: 30 min. Run: 1/35 min. Batch: 3200.
    ("JD16073", 10, "HOVLING", "SPESIALHOVEL", 30.0, 0.0285714285714286, 3200.0, date(2026, 1, 1), date(2026, 12, 31)),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 8. By Product Rules
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("By Product Rules")
headers = ["Parent Item No", "By Product Item No", "Expected Quantity", "Unit of Measure", "Market Value", "Allocation Method"]
ws.append(headers)
data = [
    # Alle produkter som hovles gir høvelspon
    ("FG001", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG002", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG003", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG004", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG005", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG006", "BP001", 0.12, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG007", "BP001", 0.12, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG008", "BP001", 0.12, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG009", "BP001", 0.12, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG010", "BP001", 0.13, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG011", "BP001", 0.13, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG012", "BP001", 0.13, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG013", "BP001", 0.18, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG014", "BP001", 0.10, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG015", "BP001", 0.10, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG016", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG017", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG018", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG019", "BP001", 0.16, "KG", 1.50, "Reduce Main Product Cost"),
    ("FG020", "BP001", 0.16, "KG", 1.50, "Reduce Main Product Cost"),

    # ── Vare JD16073 ─────────────────────────────────────────────
    # B-vare er flyttet til Co-Prod % i BOM-arket.
    # B-varen har sin egen kalkyle som samprodukt.
    # Gir også hovelspon
    ("JD16073", "BP001", 0.15, "KG", 1.50, "Reduce Main Product Cost"),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 9. Capacity Calendar
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Capacity Calendar")
headers = ["Work Center", "Date", "Available Hours", "Planned Downtime", "Available Production Hours"]
ws.append(headers)
data = [
    # Hovedhovel - januar 2026 (typisk måned)
    ("HOVEDHOVEL", date(2026, 1, 5), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 6), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 7), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 8), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 9), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 12), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 13), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 14), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 15), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 16), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 19), 16, 4, 12),  # 4t planlagt vedlikehold
    ("HOVEDHOVEL", date(2026, 1, 20), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 21), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 22), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 23), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 26), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 27), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 28), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 29), 16, 0, 16),
    ("HOVEDHOVEL", date(2026, 1, 30), 16, 0, 16),
    # Malingslinje - januar 2026
    ("MALINGSLINJE", date(2026, 1, 5), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 6), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 7), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 8), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 9), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 12), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 13), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 14), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 15), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 16), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 19), 16, 2, 14),  # 2t planlagt vedlikehold
    ("MALINGSLINJE", date(2026, 1, 20), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 21), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 22), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 23), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 26), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 27), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 28), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 29), 16, 0, 16),
    ("MALINGSLINJE", date(2026, 1, 30), 16, 0, 16),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ══════════════════════════════════════════════════════════════════
# 10. Production Scenario
# ══════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Production Scenario")
headers = ["Scenario Name", "Product", "Planned Quantity", "Start Date", "End Date"]
ws.append(headers)
data = [
    ("Normal Produksjon", "FG001", 50000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG002", 30000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG003", 20000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG006", 40000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG007", 25000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG010", 35000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG011", 20000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG013", 15000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG014", 10000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG015", 10000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG016", 25000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "FG019", 15000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Normal Produksjon", "JD16073", 3200, date(2026, 1, 1), date(2026, 12, 31)),
    ("Full Kapasitet", "FG001", 80000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Full Kapasitet", "FG002", 50000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Full Kapasitet", "FG006", 60000, date(2026, 1, 1), date(2026, 12, 31)),
    ("Full Kapasitet", "FG010", 50000, date(2026, 1, 1), date(2026, 12, 31)),
]

for row in data:
    ws.append(row)
style_header(ws, len(headers))
style_data(ws, 2, len(data) + 1, len(headers))
auto_width(ws, len(headers))

# ── Legg til beskrivelser og lagre ────────────────────────────────
add_descriptions(wb)

OUTPUT_FILE = "Produksjonsmodell_Testdata_v3.xlsx"
wb.save(OUTPUT_FILE)
print(f"OK: Testdata lagret til {OUTPUT_FILE}")
print(f"   Ark: {wb.sheetnames}")
