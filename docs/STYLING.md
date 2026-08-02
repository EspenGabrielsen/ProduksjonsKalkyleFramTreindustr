# FramTre industri — Styling-referanse

## Fargepalett

| Farge | Hex | Bruksområde |
|-------|-----|-------------|
| Dyp skogsgrønn | `#14532D` | Overskrifter (App-h2/h3, PDF-h1/h3, Excel-tittel), logo-aksent |
| Lysegrønn | `#2F855A` | Highlight-tekst, PDF-h2, Excel subheader-fill, grønn font for positive diff |
| Frisk lysegrønn | `#48BB78` | App-kort venstre border, Excel thin border |
| Mørk skoggrønn | `#2C3E2B` | Brødtekst (App body, PDF body) |
| Lys grågrønn | `#F3F5F2` | App body background, PDF `LIGHT_BG` |
| Kortbakgrunn | `#F9FBF8` | App `.fti-card` background |
| Subtil grønn | `#E6FFFA` | App highlight-bakgrunn, PDF `LIGHT_TEXT` |
| Mørkeste grønn | `#0B2819` | PDF tittelside bakgrunn (`DARK_BG`) |
| Dempet grønn | `#6B8F7D` | PDF småtekst (`MUTED_TEXT`) |
| Rød | `#C53030` | App `.fti-highlight-red`, PDF RED, Excel rød for positiv diff |
| Mørkeblå | `#1B3A5C` | App header-tittel "Produksjonskost Simulator" |
| Mellomblå | `#2C5F8A` | App header undertittel, `fti-card` p-tekst |
| Hvit | `#FFFFFF` | PDF header-tekst, Excel header-font |

## Typografi

### App (Marimo)
- **Fonte (prioritert):** `Segoe UI`, `Tahoma`, `Geneva`, `Verdana`, sans-serif
- **Brødtekstfarge:** `#2C3E2B`
- **Overskrifter:** `#14532D`, `font-weight: 600`

### PDF (ReportLab)
- **Brødtekst:** DejaVu/Helvetica, 10pt, leading 14
- **Overskrifter:** DejaVu-Bold/Helvetica-Bold
  - H1: 16pt, `#14532D`
  - H2: 13pt, `#2F855A`
  - H3: 11pt, `#14532D`
- **Kode:** DejaVuMono/Courier
- **Sidestørrelse:** A4

### Excel (openpyxl)
- **Overskrifter:** Calibri bold, 11pt, hvit (`#FFFFFF`) på grønn bakgrunn
- **Data:** Calibri, 10pt
- **Seksjonstitler:** Calibri bold, 14pt (tittel) / 12pt (seksjon) / 10pt (data)

## App (Marimo) — CSS-klasser

### `.fti-card`
```css
background: #F9FBF8;
border-radius: 12px;
padding: 24px;
margin-bottom: 20px;
box-shadow: 0 4px 12px rgba(27, 89, 43, 0.04);
border-left: 5px solid #48BB78;
```

### `.fti-card h2, .fti-card h3`
```css
color: #14532D;
margin-top: 0;
margin-bottom: 12px;
font-weight: 600;
```

### `.fti-highlight-green`
```css
color: #2F855A;
background-color: #E6FFFA;
padding: 2px 6px;
border-radius: 4px;
font-weight: 600;
```

### `.fti-highlight-red`
```css
color: #C53030;
background-color: #FFF5F5;
padding: 2px 6px;
border-radius: 4px;
font-weight: 600;
```

### App-layout
- `mo.hstack` for header (logo + tittel), `justify="space-between", align="center"`
- `mo.vstack` for vertikal stabling av elementer
- Accordion for "Vis detaljerte stamdata"
- `mo.Html` for global CSS-injeksjon

## App (Marimo) — Tabeller

### `mo.ui.data_editor` (redigerbar)
- Brukes for: råvarepriser, svinn/kapp, arbeidssentre, routing
- `editable_columns` styrer hvilke kolonner som kan redigeres
- `editable_columns=["Ny pris"]`, `["Nytt svinn %", "Ny co-prod %"]` osv.
- Støtter IKKE `label`, `column_widths` i 0.23.x

### `mo.ui.table` (skrivebeskyttet)
- Brukes for: baseline kostnader, alle accordion-tabeller, simuleringsresultater
- `selection=None` for å slå av radvalg

### Concat-konvensjon
Berikelse av eksisterende kolonner: `f"{verdi} · {ekstra}"`
- Separator: ` · ` (mellomrom + midtdot + mellomrom)
- Ved tilbakelesing som nøkkel: `.split(" · ")[0]`

Applikasjoner:
- `Beskrivelse` → `f"{p.description} · {p.base_uom}"` (råvarer)
- `Beskrivelse` → `f"{wc.description} · {wc.location_code}"` (arbeidssentre)
- `Komponent` → `f"{bl.component_item_no} · {komp_desc}"` (svinn/kapp)
- `Produkt` → `f"{bl.parent_item_no} · {prod_desc}"` (svinn/kapp)

## PDF (ReportLab) — Stiler

### Fargekonstanter
```python
PRIMARY = HexColor("#14532D")
SECONDARY = HexColor("#2F855A")
ACCENT = HexColor("#48BB78")
TEXT_COLOR = HexColor("#2C3E2B")
LIGHT_BG = HexColor("#F3F5F2")
DARK_BG = HexColor("#0B2819")
LIGHT_TEXT = HexColor("#E6FFFA")
MUTED_TEXT = HexColor("#6B8F7D")
GREEN = HexColor("#2F855A")
RED = HexColor("#C53030")
CARD_BG = HexColor("#F9FBF8")
```

### ParagraphStyles
- `h1`: 16pt, bold, PRIMARY, spaceBefore=10mm, spaceAfter=4mm
- `h2`: 13pt, bold, SECONDARY, spaceBefore=6mm, spaceAfter=3mm
- `h3`: 11pt, bold, PRIMARY, spaceBefore=4mm, spaceAfter=2mm
- `body`: 10pt, TEXT_COLOR, leading 14, TA_JUSTIFY
- `body_bold`: 10pt, bold, SECONDARY
- `code`: 9pt, MUTED_TEXT, mono-font
- `small`: 8pt, MUTED_TEXT

### Tabellstiler (PDF)
- Header-fill: PRIMARY (`#14532D`), font: hvit
- Subheader-fill: SECONDARY (`#2F855A`)
- Data-align: left (tekst) / right (tall)
- Tallformat: `#,##0.00`
- Prosentformat: `0.0%`

## Excel (openpyxl) — Stiler

### Farger
```python
header_fill = PatternFill(start_color='14532D', end_color='14532D', fill_type='solid')
subheader_fill = PatternFill(start_color='2F855A', end_color='2F855A', fill_type='solid')
```

### Fonter
```python
header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
title_font = Font(name='Calibri', bold=True, size=14, color='14532D')
section_font = Font(name='Calibri', bold=True, size=12, color='2F855A')
data_font = Font(name='Calibri', size=10)
green_font = Font(name='Calibri', size=10, color='2F855A')
red_font = Font(name='Calibri', size=10, color='C53030')
```

### Borders
```python
thin_border = Border(
    left=Side(style='thin', color='48BB78'),
    right=Side(style='thin', color='48BB78'),
    top=Side(style='thin', color='48BB78'),
    bottom=Side(style='thin', color='48BB78'),
)
```

### Tallformater
- `'#,##0.00'` for desimaltall
- `'0.0%'` for prosent

### Justering
- `Alignment(horizontal='center', vertical='center')` for headers
- `Alignment(horizontal='left', vertical='center')` for tekst
- `Alignment(horizontal='right', vertical='center')` for tall