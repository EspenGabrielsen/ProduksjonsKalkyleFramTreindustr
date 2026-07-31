#!/usr/bin/env python3
"""
generer_dokumentasjon.py - Generer stilede PDF-er fra Markdown-dokumentasjon.

Bruker Fram Treindustri styling-profil (farger, fonter, logo, tittelside)
fra generer_pdf_rapport.py.

Bruk:
    python generer_dokumentasjon.py fil.md                       # Én fil
    python generer_dokumentasjon.py fil1.md fil2.md --output     # Flere filer
    python generer_dokumentasjon.py --all                        # Alle .md bortsett fra CLINE.md
"""

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

from generer_pdf_rapport import generer_dokumentasjon_pdf, registrer_fonter

# Mapping: filnavn → (tittel, undertittel)
DOK_MAPPING = {
    "Produksjonsmodell_Dokumentasjon.md": (
        "ProduksjonsKalkyle — Systemdokumentasjon",
        "Teknisk dokumentasjon av datamodell, beregningslogikk, Marimo web-app og eksport"
    ),
    "Brukermanual_Produksjonsmodell.md": (
        "Brukermanual — Produksjonsmodellen",
        "For produksjonsledere, økonomi og innkjøp"
    ),
}

def generer_for_fil(md_path, output_dir=None):
    """Generer PDF for én Markdown-fil."""
    md_path = Path(md_path)
    if not md_path.exists():
        print(f"❌ Fil ikke funnet: {md_path}")
        return False

    filnavn = md_path.name
    if filnavn in DOK_MAPPING:
        tittel, undertittel = DOK_MAPPING[filnavn]
    else:
        tittel = filnavn.replace(".md", "").replace("_", " ")
        undertittel = None

    # Output-sti (samme mappe som .md-filen, men .pdf)
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / (md_path.stem + ".pdf")
    else:
        output_path = md_path.with_suffix(".pdf")

    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read()

    print(f"Genererer: {output_path}")
    print(f"  Tittel: {tittel}")
    if undertittel:
        print(f"   Undertittel: {undertittel}")

    try:
        generer_dokumentasjon_pdf(
            md_content,
            str(output_path),
            tittel=tittel,
            undertittel=undertittel,
        )
        print(f"  OK - Lagret: {output_path}")
        return True
    except Exception as e:
        print(f"  FEIL: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Generer stilede PDF-er fra Markdown-dokumentasjon"
    )
    parser.add_argument("filer", nargs="*", metavar="FIL", help="Markdown-fil(er) å konvertere")
    parser.add_argument("--all", action="store_true", help="Generer PDF for alle dokumentasjonsfiler")
    parser.add_argument("--output", "-o", help="Output-mappe (standard: samme som .md-filene)")
    args = parser.parse_args()

    registrer_fonter()

    if args.all:
        # Finn alle .md-filer (unntatt CLINE.md og diverse)
        filer = []
        for f in Path(".").glob("*.md"):
            navn = f.name
            if navn in ("CLINE.md", "STYLING.md", "stylingIMarimo.md", "TODO_filtrering_og_datoer.md",
                        "ForslagSimulering.md", "BATCH_SIMULERING_PLAN.md", "VERDIKJEDE_PLAN.md"):
                continue
            filer.append(str(f))
        if not filer:
            print("Ingen dokumentasjonsfiler funnet.")
            return
        print(f"Genererer PDF for {len(filer)} filer...")
        suksess = 0
        for f in filer:
            if generer_for_fil(f, args.output):
                suksess += 1
        print(f"\nOK - {suksess}/{len(filer)} PDF-er generert.")
    elif args.filer:
        suksess = 0
        for f in args.filer:
            if generer_for_fil(f, args.output):
                suksess += 1
        print(f"\nOK - {suksess}/{len(args.filer)} PDF-er generert.")
    else:
        parser.print_help()
        print("\n\nEksempler:")
        print("  python generer_dokumentasjon.py Produksjonsmodell_Dokumentasjon.md")
        print("  python generer_dokumentasjon.py --all")
        print("  python generer_dokumentasjon.py --all --output ./rapporter")


if __name__ == "__main__":
    main()