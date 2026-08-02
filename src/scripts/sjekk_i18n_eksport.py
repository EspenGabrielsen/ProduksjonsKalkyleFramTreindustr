# -*- coding: utf-8 -*-
"""Sjekk at eksport med norske fane-navn og farger fungerer."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import openpyxl

from excel_bridge import export_sqlite_to_excel, _sheet_nb_navn, _sheet_farge


def main():
    # 1) Test oversettelser og farger direkte
    print("=== Oversettelser ===")
    print("Product Master ->", _sheet_nb_navn("Product Master"))
    print("BOM ->", _sheet_nb_navn("BOM"))
    print("Transport Ruter ->", _sheet_nb_navn("Transport Ruter"))
    print("Endringslogg ->", _sheet_nb_navn("Endringslogg"))
    print("Ukjent ->", _sheet_nb_navn("Ukjent Ark"))

    print("=== Farger ===")
    print("Product Master farge ->", _sheet_farge("Product Master"))
    print("BOM farge ->", _sheet_farge("BOM"))
    print("Transport Ruter farge ->", _sheet_farge("Transport Ruter"))
    print("Capacity farge ->", _sheet_farge("Capacity Calendar"))

    # 2) Test eksport til Excel
    # Skriptet ligger i src/scripts/ → 3 dirname = prosjektroten
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "output", "test_eksport_i18n.xlsx")
    export_sqlite_to_excel(out)

    wb = openpyxl.load_workbook(out)
    print("=== Excel-faner og farger ===")
    for ws in wb.worksheets:
        print(ws.title, "| tabColor=", ws.sheet_properties.tabColor)

    # Sjekk at norske navn er brukt
    titler = [ws.title for ws in wb.worksheets]
    forventet = [
        "Produktregister",
        "Lokasjoner",
        "Arbeidssentre",
        "Operasjonsregister",
        "Varekostnader",
        "Stykkliste",
        "Produksjonsrute",
        "Biproduktregler",
        "Kapasitetskalender",
        "Produksjonsscenario",
        "Transportruter",
        "Endringslogg",
    ]
    print("=== Verifisering ===")
    ok = True
    for navn in forventet:
        if navn not in titler:
            print("MANGLER:", navn)
            ok = False
    def _hex(farge_obj):
        if farge_obj is None:
            return None
        rgb = farge_obj.rgb
        if isinstance(rgb, str) and rgb.startswith("00"):
            return rgb[2:]
        return rgb

    farger = {ws.title: _hex(ws.sheet_properties.tabColor) for ws in wb.worksheets}
    if farger.get("Produktregister") != "C53030":
        print("FEIL farge Produktregister:", farger.get("Produktregister"))
        ok = False
    if farger.get("Transportruter") != "D69E2E":
        print("FEIL farge Transportruter:", farger.get("Transportruter"))
        ok = False
    if farger.get("Stykkliste") != "3182CE":
        print("FEIL farge Stykkliste:", farger.get("Stykkliste"))
        ok = False
    if farger.get("Kapasitetskalender") is not None:
        print("FEIL: Kapasitetskalender skulle hatt ingen farge:", farger.get("Kapasitetskalender"))
        ok = False
    print("ALT_OK" if ok else "HAR_FEIL")


if __name__ == "__main__":
    main()