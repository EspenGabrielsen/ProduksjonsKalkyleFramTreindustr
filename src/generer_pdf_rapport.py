#!/usr/bin/env python3
"""
generer_simuleringsrapport.py - Genererer en PDF-rapport fra simuleringsresultater.

Brukes av varekost_app.py for å eksportere resultater til PDF.

Kjører frittstående:
    python generer_simuleringsrapport.py --help
"""

import argparse
import json
import os
import re
import sys
import urllib.request
import tempfile
from pathlib import Path
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Preformatted, HRFlowable, Spacer, Table, TableStyle, Image,
    KeepTogether, PageBreak
)
from reportlab.platypus.flowables import CondPageBreak
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- Farger (Fram Treindustri-profil) ---
PRIMARY = HexColor("#14532D")          # Dyp skogsgrønn - overskrifter
SECONDARY = HexColor("#2F855A")        # Lysegrønn - highlight-tekst
ACCENT = HexColor("#48BB78")           # Frisk lysegrønn - aksenter/kant
TEXT_COLOR = HexColor("#2C3E2B")       # Mørk skoggrønn - brødtekst
LIGHT_BG = HexColor("#F3F5F2")         # Veldig lys grågrønn - bakgrunn
DARK_BG = HexColor("#0B2819")          # Mørkeste grønn - tittelside bakgrunn
LIGHT_TEXT = HexColor("#E6FFFA")       # Subtil grønn - highlight-bakgrunn
MUTED_TEXT = HexColor("#6B8F7D")       # Dempet grønn - småtekst
GREEN = HexColor("#2F855A")            # Grønn highlight (samme som SECONDARY)
RED = HexColor("#C53030")              # Rød (fra .fti-highlight-red)
CARD_BG = HexColor("#F9FBF8")          # Kortbakgrunn (.fti-card background)


def registrer_fonter():
    """Registrer Unicode-fonter fra Windows hvis tilgjengelig."""
    windows_fonts = r"C:\Windows\Fonts"
    fonter = [
        ("DejaVu", "DejaVuSans.ttf"),
        ("DejaVu-Bold", "DejaVuSans-Bold.ttf"),
        ("DejaVu-Italic", "DejaVuSans-Oblique.ttf"),
        ("DejaVuMono", "DejaVuSansMono.ttf"),
    ]
    for font_name, font_file in fonter:
        font_path = os.path.join(windows_fonts, font_file)
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
            except Exception:
                pass


def bygg_stiler():
    """Bygg ParagraphStyles for rapporten."""
    try:
        registered = [f.fontName for f in pdfmetrics.registeredFonts()]
    except (AttributeError, TypeError):
        registered = []
    body_font = 'DejaVu' if 'DejaVu' in registered else 'Helvetica'
    body_bold_font = 'DejaVu-Bold' if 'DejaVu-Bold' in registered else 'Helvetica-Bold'
    code_font = 'DejaVuMono' if 'DejaVuMono' in registered else 'Courier'

    return {
        'h1': ParagraphStyle('H1', fontName=body_bold_font, fontSize=16, leading=22,
                              textColor=PRIMARY, spaceBefore=10*mm, spaceAfter=4*mm),
        'h2': ParagraphStyle('H2', fontName=body_bold_font, fontSize=13, leading=18,
                              textColor=SECONDARY, spaceBefore=6*mm, spaceAfter=3*mm),
        'h3': ParagraphStyle('H3', fontName=body_bold_font, fontSize=11, leading=15,
                              textColor=PRIMARY, spaceBefore=4*mm, spaceAfter=2*mm),
        'body': ParagraphStyle('Body', fontName=body_font, fontSize=10, leading=14,
                                textColor=TEXT_COLOR, spaceAfter=3*mm, alignment=TA_JUSTIFY),
        'body_bold': ParagraphStyle('BodyBold', fontName=body_bold_font, fontSize=10, leading=14,
                                     textColor=SECONDARY, spaceAfter=2*mm),
        'bullet': ParagraphStyle('Bullet', fontName=body_font, fontSize=10, leading=14,
                                  textColor=TEXT_COLOR, leftIndent=8*mm, spaceAfter=1*mm),
        'code': ParagraphStyle('Code', fontName=code_font, fontSize=7.5, leading=10,
                                textColor=TEXT_COLOR, leftIndent=4*mm, spaceAfter=3*mm,
                                backColor=LIGHT_BG, borderPadding=4),
        'small': ParagraphStyle('Small', fontName=body_font, fontSize=8, leading=11,
                                 textColor=MUTED_TEXT, spaceAfter=2*mm),
        'table_header': ParagraphStyle('TableHeader', fontName=body_bold_font, fontSize=8, leading=10,
                                        textColor=white, alignment=TA_CENTER),
        'table_cell': ParagraphStyle('TableCell', fontName=body_font, fontSize=8, leading=10,
                                      textColor=TEXT_COLOR, alignment=TA_RIGHT),
        'table_cell_left': ParagraphStyle('TableCellLeft', fontName=body_font, fontSize=8, leading=10,
                                           textColor=TEXT_COLOR, alignment=TA_LEFT),
        'comment_box': ParagraphStyle('CommentBox', fontName=body_font, fontSize=9.5, leading=14,
                                      textColor=TEXT_COLOR, spaceAfter=4*mm,
                                      backColor=CARD_BG, borderPadding=8, borderWidth=0.5, borderColor=ACCENT),
    }


# Last ned Fram Treindustri-logo og cache i temp-mappe
_LOGO_URL = "https://framtreindustri.no/wp-content/uploads/2025/08/logo-liggende-2048x512.png"
_LOGO_PATH = None


def _hent_logo():
    """Last ned logo én gang, returner cachet sti."""
    global _LOGO_PATH
    if _LOGO_PATH is not None and os.path.exists(_LOGO_PATH):
        return _LOGO_PATH
    try:
        _tmp = tempfile.mkdtemp()
        _LOGO_PATH = os.path.join(_tmp, "framtre_logo.png")
        urllib.request.urlretrieve(_LOGO_URL, _LOGO_PATH)
    except Exception:
        _LOGO_PATH = None
    return _LOGO_PATH


def lag_tittelside(canvas, doc, tittel, undertittel, dato):
    """Tegn tittelsiden med myk gradientovergang og logo."""

    # Myk gradient fra PRIMARY (topp) til DARK_BG (bunn) – 200 striper
    striper = 200
    stripe_hoyde = A4[1] / striper
    r1, g1, b1 = PRIMARY.red, PRIMARY.green, PRIMARY.blue
    r2, g2, b2 = DARK_BG.red, DARK_BG.green, DARK_BG.blue
    for i in range(striper):
        t = i / striper
        r = r1 + (r2 - r1) * t
        g = g1 + (g2 - g1) * t
        b = b1 + (b2 - b1) * t
        farge = HexColor(f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}")
        canvas.setFillColor(farge)
        canvas.rect(0, i * stripe_hoyde, A4[0], stripe_hoyde + 1, fill=1, stroke=0)

    # Logo med hvit bakgrunnsstripe (slik at grønn logo synes mot grønn gradient)
    logo_sti = _hent_logo()
    if logo_sti:
        logo_bredde = 200
        logo_hoyde = 50  # 2048:512 = 4:1 forhold
        logo_x = (A4[0] - logo_bredde) / 2
        logo_y = A4[1] * 0.70
        # Hvit bakgrunnsstripe bak logoen
        canvas.setFillColor(HexColor("#F3F5F2"))
        padding = 8
        canvas.roundRect(
            logo_x - padding, logo_y - padding,
            logo_bredde + 2 * padding, logo_hoyde + 2 * padding,
            6, fill=1, stroke=0
        )
        canvas.drawImage(logo_sti, logo_x, logo_y, width=logo_bredde, height=logo_hoyde,
                         preserveAspectRatio=True, mask='auto')

    canvas.setStrokeColor(ACCENT)
    canvas.setLineWidth(1.5)
    canvas.line(A4[0]/2, A4[1]*0.65, A4[0]/2, A4[1]*0.58)
    canvas.setLineWidth(0.8)
    canvas.line(A4[0]*0.3, A4[1]*0.56, A4[0]*0.7, A4[1]*0.56)

    canvas.setFillColor(white)
    canvas.setFont('Helvetica-Bold', 24)
    titler = tittel.split('\n') if '\n' in tittel else [tittel]
    y = A4[1] * 0.50
    for t in titler:
        canvas.drawCentredString(A4[0]/2, y, t.strip())
        y -= 32

    canvas.setStrokeColor(ACCENT)
    canvas.setLineWidth(0.5)
    canvas.line(A4[0]*0.35, y - 8, A4[0]*0.65, y - 8)

    if undertittel:
        canvas.setFillColor(LIGHT_TEXT)
        canvas.setFont('Helvetica', 12)
        canvas.drawCentredString(A4[0]/2, y - 28, undertittel)

    if dato:
        canvas.setFillColor(LIGHT_TEXT)
        canvas.setFont('Helvetica-Oblique', 9)
        canvas.drawCentredString(A4[0]/2, y - 48, dato)

    canvas.setStrokeColor(ACCENT)
    canvas.setLineWidth(0.5)
    canvas.line(A4[0]*0.35, A4[1]*0.12, A4[0]*0.65, A4[1]*0.12)

    canvas.showPage()


def lag_header_footer(canvas, doc, header_tekst):
    """Tegn header og footer."""
    canvas.saveState()
    if doc.page > 1:
        canvas.setFont('Helvetica-Oblique', 7)
        canvas.setFillColor(SECONDARY)
        canvas.drawCentredString(A4[0] / 2, A4[1] - 15*mm, header_tekst)
        canvas.setStrokeColor(ACCENT)
        canvas.setLineWidth(0.3)
        canvas.line(20*mm, A4[1] - 18*mm, A4[0] - 20*mm, A4[1] - 18*mm)
        canvas.line(20*mm, 18*mm, A4[0] - 20*mm, 18*mm)
        canvas.setFont('Helvetica-Oblique', 8)
        canvas.setFillColor(MUTED_TEXT)
        canvas.drawCentredString(A4[0] / 2, 12*mm, f'Side {doc.page}')
    canvas.restoreState()


def lag_tabell(data, styles, kolonnebredder=None):
    """Lag en pen tabell fra en liste med dicts."""
    if not data:
        return Spacer(1, 2*mm)

    kolonner = list(data[0].keys())
    headers = [Paragraph(k, styles['table_header']) for k in kolonner]

    rows = [headers]
    for rad in data:
        celle_paragraf = []
        for k in kolonner:
            verdi = rad[k]
            if isinstance(verdi, (int, float)):
                celle_paragraf.append(Paragraph(f"{verdi:,.2f}", styles['table_cell']))
            else:
                celle_paragraf.append(Paragraph(str(verdi), styles['table_cell_left']))
        rows.append(celle_paragraf)

    if kolonnebredder is None:
        kolonnebredder = [A4[0] / len(kolonner) - 15*mm] * len(kolonner)

    tbl = Table(rows, colWidths=kolonnebredder, repeatRows=1)
    tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), LIGHT_BG),
        ('GRID', (0, 0), (-1, -1), 0.5, ACCENT),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 1), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 3),
    ]))
    return tbl


def lag_ledelsessammendrag(sammenligninger, styles):
    """
    Generer et automatisert ledelsessammendrag basert på simuleringsresultater.

    Identifiserer topp 5 produkter med størst endring (i kroner og prosent),
    samt eventuelle kostnadsreduksjoner.
    """
    if not sammenligninger:
        return []

    story = []

    # Beregn totaler
    total_org = sum(c.get('org_netto', 0) for c in sammenligninger)
    total_sim = sum(c.get('sim_netto', 0) for c in sammenligninger)
    total_diff = total_sim - total_org
    total_pct = (total_diff / total_org * 100) if total_org != 0 else 0.0

    # Sorter for å finne topp/bunn
    sortert_kr = sorted(sammenligninger, key=lambda c: abs(c.get('diff_netto', 0)), reverse=True)
    sortert_pct = sorted(
        [c for c in sammenligninger if c.get('org_netto', 0) > 0],
        key=lambda c: abs(c.get('diff_netto', 0) / c['org_netto'] * 100),
        reverse=True
    )

    antall_produkter = len(sammenligninger)
    antall_okning = sum(1 for c in sammenligninger if c.get('diff_netto', 0) > 0.001)
    antall_reduksjon = sum(1 for c in sammenligninger if c.get('diff_netto', 0) < -0.001)
    antall_uendret = antall_produkter - antall_okning - antall_reduksjon

    # Seksjon: Hovedfunn
    retning = "kostnadsøkning" if total_diff >= 0 else "besparelse"
    story.append(Paragraph("Hovedfunn", styles['h2']))

    sammendrag_tekst = (
        f"Simuleringen viser en samlet {retning} på <b>{abs(total_diff):+,.2f}</b> "
        f"(<b>{total_pct:+.2f} %</b>) for porteføljen på {antall_produkter} produkter. "
    )
    if antall_okning > 0:
        sammendrag_tekst += f"{antall_okning} produkter har økt i kostnad. "
    if antall_reduksjon > 0:
        sammendrag_tekst += f"{antall_reduksjon} produkter har redusert kostnad. "
    if antall_uendret > 0:
        sammendrag_tekst += f"{antall_uendret} produkter er uendret."

    story.append(Paragraph(sammendrag_tekst, styles['comment_box']))
    story.append(Spacer(1, 3*mm))

    # Seksjon: Mest berørt (størst økning i kroner) – topp 5
    _top_kr = [c for c in sortert_kr if c.get('diff_netto', 0) > 0.001]
    if _top_kr:
        story.append(Paragraph("Mest berørt (størst økning i kroner)", styles['h3']))
        for i, c in enumerate(_top_kr[:5], 1):
            diff_netto = c.get('diff_netto', 0)
            org_netto = c.get('org_netto', 0)
            pct = (diff_netto / org_netto * 100) if org_netto > 0 else 0.0
            prod = f"{c.get('produkt', '')} – {c.get('beskrivelse', '')}"
            story.append(Paragraph(
                f"{i}. {prod}: <b>+{diff_netto:+,.2f}</b> (+{pct:.1f} %)",
                styles['bullet']
            ))
        story.append(Spacer(1, 2*mm))

    # Seksjon: Størst relativ økning (topp 5)
    _top_pct = [c for c in sortert_pct if c.get('diff_netto', 0) > 0.001]
    if _top_pct:
        story.append(Paragraph("Størst relativ økning", styles['h3']))
        for i, c in enumerate(_top_pct[:5], 1):
            diff_netto = c.get('diff_netto', 0)
            org_netto = c.get('org_netto', 0)
            pct = (diff_netto / org_netto * 100) if org_netto > 0 else 0.0
            prod = f"{c.get('produkt', '')} – {c.get('beskrivelse', '')}"
            story.append(Paragraph(
                f"{i}. {prod}: <b>+{pct:.1f} %</b> (+{diff_netto:+,.2f})",
                styles['bullet']
            ))
        story.append(Spacer(1, 2*mm))

    # Seksjon: Kostnadsreduksjoner (hvis noen)
    _reduksjoner = [c for c in sortert_kr if c.get('diff_netto', 0) < -0.001]
    if _reduksjoner:
        story.append(Paragraph("Kostnadsreduksjoner", styles['h3']))
        for i, c in enumerate(_reduksjoner[:5], 1):
            diff_netto = c.get('diff_netto', 0)
            org_netto = c.get('org_netto', 0)
            pct = (diff_netto / org_netto * 100) if org_netto > 0 else 0.0
            prod = f"{c.get('produkt', '')} – {c.get('beskrivelse', '')}"
            story.append(Paragraph(
                f"{i}. {prod}: <b>{diff_netto:+,.2f}</b> ({pct:+.1f} %)",
                styles['bullet']
            ))

    story.append(Spacer(1, 4*mm))
    return story


def _grupper_per_lokasjon(sammenligninger):
    """Grupper sammenligninger per lokasjon.

    Returnerer en liste med tupler:
        (lokasjonskode, reelle_rader, transport_rader)

    - Reelle lokasjoner (uten '->') sorteres alfabetisk.
    - Simulerte transport-lokasjoner (f.eks. 'KOD->EIK') legges under sin
      destinasjon (f.eks. 'EIK') i den rekkefølgen de dukker opp i dataene.
    """
    grupper: dict[str, dict] = {}

    for c in sammenligninger:
        loc_code = c.get('location_code', '') or ''
        if '->' in loc_code:
            dest = loc_code.split('->')[1].strip()
            grupper.setdefault(dest, {"reelle": [], "transport": []})
            grupper[dest]["transport"].append(c)
        else:
            grupper.setdefault(loc_code, {"reelle": [], "transport": []})
            grupper[loc_code]["reelle"].append(c)

    return [(loc, g["reelle"], g["transport"]) for loc, g in sorted(grupper.items())]


def generer_rapport(sammenligninger, overrides, output_path, tittel="Simuleringsrapport", kommentar=None, inkluder_detaljer=False, kapasitet_data=None):
    """
    Generer en PDF-rapport fra simuleringsresultater.

    Args:
        sammenligninger: Liste med dicts med simuleringsdata
        overrides: Dict med hvilke parametere som ble endret
        output_path: Sti til output PDF
        tittel: Tittel på rapporten
        kommentar: Valgfri ledelseskommentar/sammendrag
        inkluder_detaljer: Om tekniske produktdetaljer skal inkluderes
        kapasitet_data: Dict med arbeidssenter-timer
    """
    styles = bygg_stiler()
    dato_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        topMargin=22*mm, bottomMargin=22*mm,
        leftMargin=22*mm, rightMargin=22*mm,
    )

    story = []

    # ── Tittel ────────────────────────────────────────────────────
    story.append(Paragraph("Simuleringsrapport", styles['h1']))
    story.append(Paragraph(f"Generert: {dato_str}", styles['small']))
    story.append(HRFlowable(width="100%", thickness=0.5, color=ACCENT, spaceAfter=4*mm))

    # ── Manuell ledelseskommentar ─────────────────────────────────
    if kommentar:
        story.append(Paragraph("Sammendrag / Ledelseskommentar", styles['h2']))
        story.append(Paragraph(kommentar.replace("\n", "<br/>"), styles['comment_box']))
        story.append(Spacer(1, 4*mm))

    # ── Automatisk ledelsessammendrag ─────────────────────────────
    if sammenligninger:
        story.extend(lag_ledelsessammendrag(sammenligninger, styles))

    # ── Oppsummering av endringer ─────────────────────────────────
    story.append(Paragraph("Endringer som er simulert", styles['h2']))

    endringer = []
    if overrides.get('item_costs'):
        for vare, pris in overrides['item_costs'].items():
            endringer.append(f"• Råvarepris {vare}: satt til {pris:,.2f}")
    if overrides.get('bom_scrap'):
        for _key, scrap in overrides['bom_scrap'].items():
            # Nøkkelen kan være en tuple (parent, comp) eller en streng
            if isinstance(_key, tuple) and len(_key) == 2:
                parent, comp = _key
            else:
                parent, comp = str(_key), ""
            endringer.append(f"• Svinn {comp} → {parent}: satt til {scrap}%")
    if overrides.get('work_centers'):
        for wc, wc_endr in overrides['work_centers'].items():
            for felt, verdi in wc_endr.items():
                felt_navn = {"labor_cost_hour": "Lønn/time", "machine_cost_hour": "Maskin/time",
                             "overhead_cost_hour": "Overhead/time"}.get(felt, felt)
                endringer.append(f"• Arbeidssenter {wc} - {felt_navn}: satt til {verdi:,.0f}")
    if overrides.get('routing'):
        for _key, rt_endr in overrides['routing'].items():
            # Nøkkelen kan være en tuple (item, op, wc) eller en streng
            if isinstance(_key, tuple) and len(_key) == 3:
                item, op, wc = _key
            else:
                item, op, wc = str(_key), "", ""
            for felt, verdi in rt_endr.items():
                felt_navn = {"run_time_minutes": "Run time (min)", "batch_size": "Batch"}.get(felt, felt)
                endringer.append(f"• Routing {item} op.{op} @ {wc} - {felt_navn}: satt til {verdi}")
    if overrides.get('planned_quantity'):
        endringer.append(f"• Planlagt kvantum: {overrides['planned_quantity']:,.0f} stk")

    if endringer:
        for e in endringer:
            story.append(Paragraph(e, styles['bullet']))
    else:
        story.append(Paragraph("Ingen parametere ble endret (sammenligner baseline).", styles['body']))

    story.append(Spacer(1, 4*mm))

    # ── Sammenligningstabell ──────────────────────────────────────
    story.append(Paragraph("Sammenligning: Baseline vs Simulert", styles['h2']))

    if sammenligninger:
        # Grupper per lokasjon: KOD, EIK, KV osv. hver i sin egen tabell.
        # Simulerte transport-lokasjoner (f.eks. KOD->EIK) legges under
        # destinasjonen de går til (f.eks. EIK), i den rekkefølgen de
        # dukker opp i dataene.
        lokasjonsgrupper = _grupper_per_lokasjon(sammenligninger)
        kol_bredder = [28*mm, 45*mm, 30*mm, 30*mm, 27*mm, 25*mm]

        for lok, reelle_rader, transport_rader in lokasjonsgrupper:
            lok_tittel = lok
            _navn = ""
            if reelle_rader:
                _navn = reelle_rader[0].get('location_name', '') or ""
            elif transport_rader:
                _navn = transport_rader[0].get('location_name', '') or ""
            if _navn and _navn != lok:
                lok_tittel = f"{lok} — {_navn}"
            # Start ny side hvis det er mindre enn 33% igjen, slik at
            # tabellen ikke blir delt over et sideskift
            story.append(CondPageBreak(A4[1] * 2 / 6))
            story.append(Paragraph(f"Lokasjon: {lok_tittel}", styles['h3']))

            # Reelle produkter ved lokasjonen
            if reelle_rader:
                tabell_data = []
                for c in reelle_rader:
                    diff_netto = c.get('diff_netto', 0)
                    org_netto = c.get('org_netto', 0)
                    endring_pct = (diff_netto / org_netto * 100) if org_netto != 0 else 0.0
                    tabell_data.append({
                        "Produkt": c.get('produkt', ''),
                        "Beskrivelse": c.get('beskrivelse', ''),
                        "Org. netto": org_netto,
                        "Sim. netto": c.get('sim_netto', 0),
                        "Diff (kr)": diff_netto,
                        "Endring %": endring_pct,
                    })
                story.append(lag_tabell(tabell_data, styles, kol_bredder))
                story.append(Spacer(1, 2*mm))

            # Simulerte transport-rader til denne lokasjonen
            if transport_rader:
                story.append(Paragraph(f"Transport til {lok}:", styles['body_bold']))
                tabell_data = []
                for c in transport_rader:
                    diff_netto = c.get('diff_netto', 0)
                    org_netto = c.get('org_netto', 0)
                    endring_pct = (diff_netto / org_netto * 100) if org_netto != 0 else 0.0
                    tabell_data.append({
                        "Produkt": f"{c.get('location_code', '')} · {c.get('produkt', '')}",
                        "Beskrivelse": c.get('beskrivelse', ''),
                        "Org. netto": org_netto,
                        "Sim. netto": c.get('sim_netto', 0),
                        "Diff (kr)": diff_netto,
                        "Endring %": endring_pct,
                    })
                story.append(lag_tabell(tabell_data, styles, kol_bredder))
                story.append(Spacer(1, 3*mm))
            elif not reelle_rader:
                story.append(Spacer(1, 3*mm))

        # ── Total effekt ──────────────────────────────────────────
        total_org = sum(c.get('org_netto', 0) for c in sammenligninger)
        total_sim = sum(c.get('sim_netto', 0) for c in sammenligninger)
        total_diff = total_sim - total_org
        total_pct = (total_diff / total_org * 100) if total_org != 0 else 0.0

        story.append(Paragraph("Total effekt", styles['h3']))
        story.append(Paragraph(f"<b>Opprinnelig total netto kostnad:</b> {total_org:,.2f}", styles['body']))
        story.append(Paragraph(f"<b>Simulert total netto kostnad:</b> {total_sim:,.2f}", styles['body']))
        if total_diff >= 0:
            story.append(Paragraph(f"<b>Økning:</b> {total_diff:+,.2f} (kostnadsøkning, {total_pct:+.2f} %)", styles['body']))
        else:
            story.append(Paragraph(f"<b>Reduksjon:</b> {total_diff:+,.2f} (besparelse, {total_pct:+.2f} %)", styles['body']))

        # ── Kapasitetsutnyttelse ──────────────────────────────────
        if kapasitet_data:
            story.append(Spacer(1, 4*mm))
            story.append(Paragraph("Kapasitetsutnyttelse", styles['h2']))
            story.append(Paragraph("Totalt timebehov per arbeidssenter for dette scenariet:", styles['body']))
            cap_rows = []
            for wc, hours in kapasitet_data.items():
                cap_rows.append({
                    "Arbeidssenter": wc,
                    "Simulert timebehov": f"{hours:,.1f} timer"
                })
            story.append(lag_tabell(cap_rows, styles, [60*mm, 60*mm]))

        # ── Detaljer per produkt ──────────────────────────────────
        _first_detail = True
        for c in (sammenligninger if inkluder_detaljer else []):
            if _first_detail:
                story.append(Spacer(1, 4*mm))
                story.append(Paragraph("Detaljer per produkt", styles['h2']))
                _first_detail = False
            prod_navn = f"{c.get('produkt', '')} - {c.get('beskrivelse', '')}"
            story.append(Paragraph(prod_navn, styles['h3']))

            # Materialdetaljer
            mat_det = c.get('materialdetaljer', [])
            if mat_det:
                story.append(Paragraph("<b>Materialdetaljer:</b>", styles['body_bold']))
                mat_data = []
                for md in mat_det:
                    mat_data.append({
                        "Komponent": md.get('komponent', ''),
                        "Qty per": md.get('qty_per', 0),
                        "Svinn %": md.get('scrap_pct', 0),
                        "Enhetskost": md.get('enhetskost', 0),
                        "Total": md.get('total_kost', 0),
                    })
                story.append(lag_tabell(mat_data, styles))
                story.append(Spacer(1, 2*mm))

            # Operasjonsdetaljer
            op_det = c.get('operasjonsdetaljer', [])
            if op_det:
                story.append(Paragraph("<b>Operasjonsdetaljer:</b>", styles['body_bold']))
                op_data = []
                for od in op_det:
                    op_data.append({
                        "Operasjon": od.get('operasjon', ''),
                        "Arb.senter": od.get('arbeidssenter', ''),
                        "Run (min)": od.get('run_time', 0),
                        "Batch": od.get('batch', 0),
                        "Kost/time": od.get('kost_per_time', 0),
                        "Total": od.get('total_kost', 0),
                    })
                story.append(lag_tabell(op_data, styles))
                story.append(Spacer(1, 2*mm))

            # Biprodukter
            bp_det = c.get('biprodukter', [])
            if bp_det:
                story.append(Paragraph("<b>Biprodukter:</b>", styles['body_bold']))
                bp_data = []
                for bd in bp_det:
                    bp_data.append({
                        "Biprodukt": bd.get('biprodukt', ''),
                        "Kvantum": bd.get('kvantum', 0),
                        "Markedsverdi": bd.get('markedsverdi', 0),
                        "Total verdi": bd.get('total_verdi', 0),
                    })
                story.append(lag_tabell(bp_data, styles))
                story.append(Spacer(1, 2*mm))

            # Co-produkter (f.eks. B-vare)
            co_det = c.get('coprodukter', [])
            if co_det:
                story.append(Paragraph("<b>Co-produkter (B-vare):</b>", styles['body_bold']))
                for co in co_det:
                    co_data = []
                    co_data.append({
                        "Post": "Produkt",
                        "Verdi": f"{co.get('produkt', '')} - {co.get('beskrivelse', '')}",
                    })
                    co_data.append({
                        "Post": "Materialkost",
                        "Verdi": f"{co.get('materialkost', 0):,.4f}",
                    })
                    co_data.append({
                        "Post": "Operasjonskost",
                        "Verdi": f"{co.get('operasjonskost', 0):,.4f}",
                    })
                    co_data.append({
                        "Post": "Setupkost",
                        "Verdi": f"{co.get('setupkost', 0):,.4f}",
                    })
                    co_data.append({
                        "Post": "Brutto",
                        "Verdi": f"{co.get('brutto', 0):,.4f}",
                    })
                    if co.get('biproduktverdi', 0) > 0:
                        co_data.append({
                            "Post": "Biproduktverdi",
                            "Verdi": f"-{co.get('biproduktverdi', 0):,.4f}",
                        })
                    co_data.append({
                        "Post": "Netto",
                        "Verdi": f"{co.get('netto', 0):,.4f}",
                    })
                    story.append(lag_tabell(co_data, styles, [60*mm, 60*mm]))
                    story.append(Spacer(1, 2*mm))

            # Scenariototaler
            if c.get('kvantum'):
                story.append(Paragraph(
                    f"<b>Scenariototal:</b> {c['kvantum']:,.0f} enheter → "
                    f"Total netto: {c.get('total_netto', 0):,.2f} | "
                    f"Kost/enhet: {c.get('kost_per_enhet', 0):,.2f} | "
                    f"Timebehov: {c.get('timebehov', 0):,.2f} timer",
                    styles['body']
                ))

            story.append(Spacer(1, 2*mm))

    else:
        story.append(Paragraph("Ingen simuleringsresultater tilgjengelig.", styles['body']))

    # ── Generer PDF ───────────────────────────────────────────────
    doc.build(story,
              onFirstPage=lambda c, d: lag_tittelside(c, d, tittel, "Simuleringsresultater", dato_str),
              onLaterPages=lambda c, d: lag_header_footer(c, d, tittel))

    return output_path


# ═══════════════════════════════════════════════════════════════════════════
#  Markdown → PDF-dokumentasjon
# ═══════════════════════════════════════════════════════════════════════════

def _parse_md_table(linjer, styles):
    """Parse en Markdown-tabell (pipe-syntax) til ReportLab Table."""
    if not linjer:
        return None
    # Fjern tomme linjer før/etter
    while linjer and not linjer[0].strip():
        linjer.pop(0)
    while linjer and not linjer[-1].strip():
        linjer.pop()
    if len(linjer) < 2:
        return None

    # Første linje = header, andre linje = separator (|---|)
    header = [c.strip() for c in linjer[0].split("|")[1:-1]]
    if not header:
        return None

    # Data-rader
    data_rows = []
    for rad in linjer[2:]:
        raden = rad.strip()
        if not raden or raden.startswith("|--") or raden.startswith("|---"):
            continue
        celler = [c.strip() for c in raden.split("|")[1:-1]]
        if celler:
            # Fyll opp til samme antall kolonner
            while len(celler) < len(header):
                celler.append("")
            data_rows.append(celler[:len(header)])

    if not data_rows:
        return None

    # Lag tabell
    kolonner = header
    headers_par = [Paragraph(_md_to_html(k), styles['table_header']) for k in kolonner]
    rows_pdf = [headers_par]
    for rad in data_rows:
        celle_pars = []
        for verdi in rad:
            # Prøv tallformat
            try:
                v = verdi.replace(" ", "").replace(",", ".")
                float(v)
                celle_pars.append(Paragraph(verdi, styles['table_cell']))
            except (ValueError, AttributeError):
                celle_pars.append(Paragraph(_md_to_html(verdi), styles['table_cell_left']))
        while len(celle_pars) < len(kolonner):
            celle_pars.append(Paragraph("", styles['table_cell']))
        rows_pdf.append(celle_pars[:len(kolonner)])

    # Kolonnebredder
    tilgjengelig = A4[0] - 44*mm
    kol_bredder = [tilgjengelig / len(kolonner)] * len(kolonner)

    tbl = Table(rows_pdf, colWidths=kol_bredder, repeatRows=1)
    tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('TOPPADDING', (0, 0), (-1, 0), 5),
        ('BACKGROUND', (0, 1), (-1, -1), LIGHT_BG),
        ('GRID', (0, 0), (-1, -1), 0.5, ACCENT),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 1), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 3),
    ]))

    # Wrap i KeepTogether for å unngå sidebrekk midt i tabell
    return KeepTogether([tbl, Spacer(1, 3*mm)])


def _parse_markdown_to_story(md_content, styles):
    """Parse Markdown-innhold til en liste med ReportLab flowables."""
    story = []

    if not md_content or not md_content.strip():
        return story

    linjer = md_content.split("\n")

    # State for kodeblokk og tabell
    i = 0
    in_code_block = False
    code_buffer = []
    in_table = False
    table_buffer = []
    in_list = False

    while i < len(linjer):
        linje = linjer[i]
        linje_stripped = linje.strip()

        # Kodeblokk (``` eller `````)
        if linje_stripped.startswith("```"):
            if in_code_block:
                # Avslutt kodeblokk
                in_code_block = False
                code_text = "\n".join(code_buffer)
                story.append(Preformatted(code_text, styles['code']))
                story.append(Spacer(1, 2*mm))
                code_buffer = []
            else:
                in_code_block = True
                code_buffer = []
            i += 1
            continue

        if in_code_block:
            code_buffer.append(linje)
            i += 1
            continue

        # Tom linje
        if not linje_stripped:
            # Avslutt tabell hvis vi var i en
            if in_table and len(table_buffer) >= 2:
                result = _parse_md_table(table_buffer, styles)
                if result:
                    story.append(result)
                table_buffer = []
                in_table = False
            in_list = False
            i += 1
            continue

        # Horisontal linje ---
        if re.match(r"^---+$", linje_stripped):
            story.append(Spacer(1, 1*mm))
            story.append(HRFlowable(width="100%", thickness=0.5, color=ACCENT, spaceAfter=3*mm))
            i += 1
            continue

        # Sjekk om dette er en tabell (inneholder |)
        if "|" in linje_stripped and linje_stripped.startswith("|"):
            # Sjekk om neste linje er separator (|---|)
            if (i + 1) < len(linjer) and re.match(r"^[\|\s\-:]+$", linjer[i + 1].strip()):
                in_table = True
                table_buffer.append(linje)
                table_buffer.append(linjer[i + 1])
                i += 2
                continue
            elif in_table:
                table_buffer.append(linje)
                i += 1
                continue
            else:
                # Enkeltstående | er ikke en tabell, behandle som tekst
                pass

        # Avslutt tabell hvis vi var i en
        if in_table and len(table_buffer) >= 2:
            result = _parse_md_table(table_buffer, styles)
            if result:
                story.append(result)
            table_buffer = []
            in_table = False

        # Blockquote (linje som starter med >) – samle påfølgende linjer til én boks
        if linje_stripped.startswith(">"):
            quote_linjer = []
            while i < len(linjer) and linjer[i].strip().startswith(">"):
                q = linjer[i].strip().lstrip("> ").strip()
                if q:
                    quote_linjer.append(_md_to_html(q))
                i += 1
            if quote_linjer:
                quote_html = "<br/>".join(quote_linjer)
                story.append(Paragraph(quote_html, styles['comment_box']))
            continue

        # Overskrifter H1-H3
        if linje_stripped.startswith("### "):
            tekst = _md_to_html(linje_stripped[4:])
            story.append(Paragraph(tekst, styles['h3']))
            i += 1
            continue
        if linje_stripped.startswith("## "):
            # Start kapitlet på ny side hvis vi har passert threshold av siden.
            # CondPageBreak(threshold) bryter kun hvis gjenværende plass < threshold.
            # eksempel 2/6 brytter når det er mindre enn 33% igjen av siden
            story.append(CondPageBreak(A4[1] * 2 / 6))
            tekst = _md_to_html(linje_stripped[3:])
            story.append(Paragraph(tekst, styles['h2']))
            i += 1
            continue
        if linje_stripped.startswith("# "):
            tekst = _md_to_html(linje_stripped[2:])
            story.append(Paragraph(tekst, styles['h1']))
            i += 1
            continue

        # Punktliste
        if linje_stripped.startswith("- ") or linje_stripped.startswith("* "):
            tekst = _md_to_html(linje_stripped[2:])
            story.append(Paragraph(f"• {tekst}", styles['bullet']))
            in_list = True
            i += 1
            continue

        # Nummerert liste (alle tall, ikke bare 1-3)
        match_num = re.match(r"^(\d+)\.\s*(.*)", linje_stripped)
        if match_num:
            nummer = match_num.group(1)
            tekst = _md_to_html(match_num.group(2))
            story.append(Paragraph(f"<b>{nummer}.</b> {tekst}", styles['bullet']))
            in_list = True
            i += 1
            continue

        # Vanlig tekst
        if linje_stripped:
            tekst = _md_to_html(linje_stripped)
            story.append(Paragraph(tekst, styles['body']))
        i += 1

    # Hvis vi slutter midt i en tabell, flush den
    if in_table and len(table_buffer) >= 2:
        result = _parse_md_table(table_buffer, styles)
        if result:
            story.append(result)

    return story


def _md_to_html(tekst):
    """Konverter Markdown inline-formatering til HTML for ReportLab."""
    # Fjern emoji-tegn som PDF-fontene (DejaVu) ikke har glyfer for.
    # Uten dette vises de som små sorte firkanter (tofu) i PDF-en.
    tekst = re.sub(
        r"[\U0001F000-\U0001FFFF\u2600-\u27BF\uFE00-\uFE0F\u200D\uFE0F]",
        "", tekst
    )
    # [tekst](url) → tekst (PDF-parseren lager ikke klikkbare linker)
    tekst = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", tekst)
    # Fjern <img>-taggar (ReportLab Paragraph støtter ikke alt-attributt her)
    tekst = re.sub(r"<img[^>]*?>", "", tekst, flags=re.IGNORECASE)
    # **fet**
    tekst = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", tekst)
    # *kursiv*
    tekst = re.sub(r"\*(.*?)\*", r"<i>\1</i>", tekst)
    # ``kode``
    tekst = re.sub(r"``(.*?)``", r"<font face='Courier'>\1</font>", tekst)
    # `kode`
    tekst = re.sub(r"`(.*?)`", r"<font face='Courier'>\1</font>", tekst)
    # Fjern gjenværende HTML-taggar som ReportLab ikke støtter (alt, class, etc.)
    tekst = re.sub(r"<([a-zA-Z]+)[^>]*?(\/?)>", r"<\1\2>", tekst)
    return tekst


def generer_dokumentasjon_pdf(markdown_content, output_path, tittel="Dokumentasjon", undertittel=None, kommentar=None):
    """
    Generer en stylet PDF fra Markdown-dokumentasjon.

    Gjenbruker all styling fra generer_pdf_rapport.py (farger, fonter, stiler,
    tittelside, header/footer, tabeller).

    Args:
        markdown_content: Rå Markdown-tekst
        output_path: Sti til output PDF
        tittel: Tittel på dokumentet (vises på tittelsiden)
        undertittel: Valgfri undertittel
        kommentar: Valgfri beskrivelse/ingress øverst
    """
    registrer_fonter()
    styles = bygg_stiler()
    dato_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        topMargin=22*mm, bottomMargin=22*mm,
        leftMargin=22*mm, rightMargin=22*mm,
    )

    story = []

    # Tittel
    story.append(Paragraph(tittel, styles['h1']))
    story.append(Paragraph(f"Generert: {dato_str}", styles['small']))
    if undertittel:
        story.append(Paragraph(undertittel, styles['body_bold']))
    story.append(HRFlowable(width="100%", thickness=0.5, color=ACCENT, spaceAfter=4*mm))

    # Kommentar/ingress
    if kommentar:
        story.append(Paragraph(kommentar, styles['comment_box']))
        story.append(Spacer(1, 4*mm))

    # Inline-innholdsfortegnelse — pars H1 og H2 for TOC.
    # Første H1 (dokumenttittel) og første H2 (undertittel) hoppes over,
    # da de allerede vises øverst i dokumentet og på tittelsiden.
    # Et eventuelt "Innhold"-kapittel hoppes også over — det erstattes
    # fullstendig av den automatiske innholdsfortegnelsen.
    _linjer = markdown_content.split("\n")
    toc_items = []
    _innhold_linjer = []
    _sett_tittel = False
    _sett_undertittel = False
    _skip_innhold = False
    for _l in _linjer:
        _s = _l.strip()
        if not _sett_tittel and _s.startswith("# "):
            _sett_tittel = True
            continue
        if not _sett_undertittel and _s.startswith("## "):
            _sett_undertittel = True
            continue
        if _skip_innhold:
            if _s.startswith("## "):
                _skip_innhold = False
            else:
                continue
        if _s.startswith("## ") and _s[3:].strip().lower() in ("innhold", "innholdsfortegnelse"):
            _skip_innhold = True
            continue
        _innhold_linjer.append(_l)
        if _s.startswith("## "):
            toc_items.append((2, _s[3:].replace("**", "").replace("`", "")))
        elif _s.startswith("# "):
            toc_items.append((1, _s[2:].replace("**", "").replace("`", "")))
    _markdown_innhold = "\n".join(_innhold_linjer)

    if toc_items:
        story.append(Paragraph("Innholdsfortegnelse", styles['h2']))
        for _niv, _navn in toc_items:
            if _niv == 1:
                story.append(Paragraph(f"<b>{_navn}</b>", styles['bullet']))
            else:
                story.append(Paragraph(f"  {_navn}", styles['small']))
        story.append(Spacer(1, 3*mm))
        story.append(HRFlowable(width="100%", thickness=0.5, color=ACCENT, spaceAfter=4*mm))

    # Parse Markdown
    story.extend(_parse_markdown_to_story(_markdown_innhold, styles))

    # Footer
    story.append(Spacer(1, 10*mm))
    story.append(HRFlowable(width="100%", thickness=0.3, color=MUTED_TEXT, spaceAfter=2*mm))
    story.append(Paragraph(
        f"{tittel} — Generert {dato_str}",
        styles['small']
    ))

    # Bygg PDF
    _full_tittel = tittel + (" — " + undertittel if undertittel else "")
    doc.build(story,
              onFirstPage=lambda c, d: lag_tittelside(c, d, tittel=tittel,undertittel=undertittel, dato=dato_str),
              onLaterPages=lambda c, d: lag_header_footer(c, d, _full_tittel))

    return output_path


def main():
    parser = argparse.ArgumentParser(description='Generer PDF-rapport fra simuleringsdata')
    parser.add_argument('--data', help='JSON-fil med simuleringsdata')
    parser.add_argument('--output', '-o', default='simuleringsrapport.pdf', help='Output PDF-fil')
    parser.add_argument('--tittel', '-t', default='Simuleringsrapport', help='Tittel på rapporten')
    parser.add_argument('--dokumentasjon', '-d', metavar='FIL', help='Generer dokumentasjon-PDF fra Markdown-fil')
    parser.add_argument('--undertittel', '-u', help='Undertittel for dokumentasjon-PDF')
    args = parser.parse_args()

    if args.dokumentasjon:
        with open(args.dokumentasjon, 'r', encoding='utf-8') as f:
            md_content = f.read()
        generer_dokumentasjon_pdf(md_content, args.output, args.tittel, args.undertittel)
        print(f"Dokumentasjon-PDF generert: {args.output}")
    elif args.data:
        with open(args.data, 'r', encoding='utf-8') as f:
            data = json.load(f)
        generer_rapport(data.get('sammenligninger', []),
                        data.get('overrides', {}),
                        args.output,
                        args.tittel)
        print(f"Rapport generert: {args.output}")
    else:
        print("Bruk: python generer_simuleringsrapport.py --data data.json --output rapport.pdf")
        print("  eller: python generer_simuleringsrapport.py --dokumentasjon fil.md --output dokumentasjon.pdf --tittel 'Min Tittel'")


if __name__ == "__main__":
    registrer_fonter()
    main()
