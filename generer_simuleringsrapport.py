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
import sys
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Preformatted, HRFlowable, Spacer, Table, TableStyle
)
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- Farger ---
PRIMARY = HexColor("#1B3A5C")
SECONDARY = HexColor("#2C5F8A")
ACCENT = HexColor("#4A90D9")
TEXT_COLOR = HexColor("#2D2D2D")
LIGHT_BG = HexColor("#F5F8FC")
DARK_BG = HexColor("#0F2440")
LIGHT_TEXT = HexColor("#C8D7EB")
MUTED_TEXT = HexColor("#96AFC8")
GREEN = HexColor("#1B8A3C")
RED = HexColor("#C0392B")


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
                                      backColor=HexColor("#F5F8FC"), borderPadding=8, borderWidth=0.5, borderColor=HexColor("#1B3A5C")),
    }


def lag_tittelside(canvas, doc, tittel, undertittel, dato):
    """Tegn tittelsiden."""
    canvas.setFillColor(PRIMARY)
    canvas.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)

    canvas.setFillColor(DARK_BG)
    canvas.rect(0, 0, A4[0], A4[1] * 0.4, fill=1, stroke=0)

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
        canvas.setFillColor(MUTED_TEXT)
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
        canvas.setFillColor(HexColor("#888888"))
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

    # ── Ledelsessammendrag / Kommentar ────────────────────────────
    if kommentar:
        story.append(Paragraph("Sammendrag / Ledelseskommentar", styles['h2']))
        story.append(Paragraph(kommentar.replace("\n", "<br/>"), styles['comment_box']))
        story.append(Spacer(1, 4*mm))

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
        tabell_data = []
        for c in sammenligninger:
            diff_netto = c.get('diff_netto', 0)
            diff_str = f"{diff_netto:+,.2f}"
            tabell_data.append({
                "Produkt": c.get('produkt', ''),
                "Beskrivelse": c.get('beskrivelse', ''),
                "Org. netto": c.get('org_netto', 0),
                "Sim. netto": c.get('sim_netto', 0),
                "Diff": diff_str,
            })

        kol_bredder = [30*mm, 50*mm, 35*mm, 35*mm, 30*mm]
        tbl = lag_tabell(tabell_data, styles, kol_bredder)
        story.append(tbl)
        story.append(Spacer(1, 4*mm))

        # ── Total effekt ──────────────────────────────────────────
        total_org = sum(c.get('org_netto', 0) for c in sammenligninger)
        total_sim = sum(c.get('sim_netto', 0) for c in sammenligninger)
        total_diff = total_sim - total_org

        story.append(Paragraph("Total effekt", styles['h3']))
        story.append(Paragraph(f"<b>Opprinnelig total netto kostnad:</b> {total_org:,.2f}", styles['body']))
        story.append(Paragraph(f"<b>Simulert total netto kostnad:</b> {total_sim:,.2f}", styles['body']))
        if total_diff >= 0:
            story.append(Paragraph(f"<b>Økning:</b> {total_diff:+,.2f} (kostnadsøkning)", styles['body']))
        else:
            story.append(Paragraph(f"<b>Reduksjon:</b> {total_diff:+,.2f} (besparelse)", styles['body']))

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


def main():
    parser = argparse.ArgumentParser(description='Generer PDF-rapport fra simuleringsdata')
    parser.add_argument('--data', help='JSON-fil med simuleringsdata')
    parser.add_argument('--output', '-o', default='simuleringsrapport.pdf', help='Output PDF-fil')
    parser.add_argument('--tittel', '-t', default='Simuleringsrapport', help='Tittel på rapporten')
    args = parser.parse_args()

    if args.data:
        with open(args.data, 'r', encoding='utf-8') as f:
            data = json.load(f)
        generer_rapport(data.get('sammenligninger', []),
                        data.get('overrides', {}),
                        args.output,
                        args.tittel)
        print(f"Rapport generert: {args.output}")
    else:
        print("Bruk: python generer_simuleringsrapport.py --data data.json --output rapport.pdf")


if __name__ == "__main__":
    registrer_fonter()
    main()
