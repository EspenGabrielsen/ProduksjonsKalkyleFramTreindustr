"""Analyser optimal batch-storrelse vs lagerkostnad for et produkt.

Matematikk:
  - Setup per enhet: setup / batch
  - Lagerkost per enhet: avhenger av hvor mange uker batchen dekker.
    Med ukentlig salg S og batch B dekker batchen T = B/S uker.
    Gjennomsnittlig lager = B/2 over hele perioden.
    Total lagerkost = (B/2) * T * enhetsverdi * (hold_pct/100 per uke)
    Per produsert enhet = total / B = (B * enhetsverdi * hold_pct) / (2 * S * 100)

Bruk (to moduser):
  # Modus A: Fra MILP-resultat (etter a kjore optimization_engine.py)
  python src/scripts/analyser_batch.py --milp resultat.json JD19148

  # Modus B: Direkte fra SQLite-database (uten MILP)
  python src/scripts/analyser_batch.py --db src/produksjonskalkyle.db JD19148
"""
import argparse
import json
import sys
from pathlib import Path

_SRC_DIR = Path(__file__).resolve().parents[1]
if str(_SRC_DIR) not in sys.path:
    sys.path.insert(0, str(_SRC_DIR))


def _hent_historisk_salg(prod: str, db_path: str = "src/produksjonskalkyle.db") -> float:
    """Hent aggregert historisk salg (LM) for et produkt fra SQLite."""
    try:
        from data_repo import DataRepo
        db = DataRepo(db_path)
        db.initialize()
        row = db.conn.execute(
            "SELECT SUM(quantity) as t FROM historical_sales WHERE product_id=?",
            (prod,)
        ).fetchone()
        return row["t"] or 0.0
    except Exception:
        return 0.0


def _print_batch_tabell(setup: float, base: float, S: float, hold_pct: float = 2.0,
                        optimal_batch: float = 0.0):
    """Skriv batch-størrelse vs kostnad-tabell.

    Viser batchkost (én enkelt batch) og årlig kost (52 uker) slik at
    effekten av ulik `hold_pct` blir synlig på årsbasis. Optimal-batchen
    markeres med [OPT] foran raden.
    """
    print(f"\nBATCH-STØRRELSE VS KOSTNAD (setup={setup:.2f} kr, verdi={base:.4f} kr/LM, "
          f"hold={hold_pct}%/uke, ukessalg~{S:.0f} LM):")

    # Legger til 6 tegn på starten av overskriften for å matche prefiks-bredden ("[OPT] ")
    print(f"      {'Batch LM':>12} {'Uker dk.':>9} {'Setup/enh':>11} {'Hold/enh':>10} "
          f"{'Total/enh':>10} {'Batchkost':>11} {'Årlig kost':>12}")
    print(" " * 6 + "-" * 80)

    batcher = [1_000, 2_000, 5_000, 10_000, 20_000, 30_000, 40_000, 50_000, 70_000, 90_000, 100_000, 150_000]
    if optimal_batch and optimal_batch > 0:
        _opt_avrundet = round(optimal_batch, 0)
        if not any(abs(b - _opt_avrundet) < 1 for b in batcher):
            batcher.append(_opt_avrundet)
            
    batcher = sorted(batcher)

    for batch in batcher:
        setup_e = setup / batch
        T = batch / S  # antall uker batchen dekker
        hold_e = (T * base * hold_pct / 100.0) / 2.0  # per produsert enhet
        total_e = setup_e + hold_e
        batch_kr = total_e * batch
        aarlig_kr = total_e * S * 52.0  # årlig kost (52 uker)

        # Prefiks må være nøyaktig like lang (5 tegn + 1 mellomrom = 6 tegn)
        is_opt = optimal_batch and abs(batch - optimal_batch) < 1
        marker = "[OPT] " if is_opt else "      "

        print(f"{marker}{batch:>12,.0f} {T:>9,.1f} {setup_e:>11.4f} {hold_e:>10.4f} "
              f"{total_e:>10.4f} {batch_kr:>11,.0f} {aarlig_kr:>12,.0f}")


def analyse_fra_milp(json_path: str, prod: str, hold_pct: float = 2.0):
    """Modus A: Analyser fra MILP-resultat (resultat.json).

    Bruker 'actual_batch_size' fra optimeringen og 'batch_decisions'
    for setup-kost. Historisk salg er primær etterspørselskilde;
    fallback til produksjonsplanens kvantum per uke.
    """
    with open(json_path, "r", encoding="utf-8") as f:
        d = json.load(f)

    rows = [r for r in d.get("production_plan", []) if r.get("product") == prod]
    if not rows:
        print(f"Produkt {prod} ikke funnet i produksjonsplanen.")
        return

    print(f"PRODUKT: {prod}")
    print(f"Antall produksjonslinjer: {len(rows)}")
    for r in rows:
        print(f"  uke {r['period']}: {r['quantity']:>10,.1f} LM  "
              f"batch={r.get('actual_batch_size')}  setup/enh={r.get('setup_per_unit')}  "
              f"hold/enh/u={r.get('holding_cost_per_unit')}  kost={r['total_cost']:.2f}")

    bc = [b for b in d.get("batch_decisions", []) if b.get("product") == prod]
    if not bc:
        print(f"\nIngen batch-decision for {prod}")
        return
    setup = bc[0]["setup_cost"]
    base = rows[0]["base_unit_cost"]

    # Ukentlig salg: historisk salg hvis mulig, ellers anta batch/52
    hist = _hent_historisk_salg(prod)
    S = hist / 52.0 if hist > 0 else rows[0]["quantity"] / 52.0
    if S <= 0:
        S = 738.0  # fallback

    # Optimal batch per produkt (EOQ) basert på årlig etterspørsel
    _årlig = hist if hist > 0 else S * 52.0
    _optimal = _optimal_batch_eoq(_årlig, setup, base, hold_pct)

    _print_batch_tabell(setup, base, S, hold_pct, _optimal)


def _beregn_setup_per_batch(op_details) -> float:
    """Beregn total setupkost per batch fra operasjonsdetaljer."""
    total = 0.0
    for od in op_details:
        setup_min = getattr(od, "setup_time_min", 0) or 0.0
        cost_per_hour = getattr(od, "cost_per_hour", 0) or 0.0
        total += (setup_min / 60.0) * cost_per_hour
    return total


def _optimal_batch_eoq(annual_demand: float, setup_cost: float,
                       unit_cost: float, hold_pct: float = 2.0) -> float:
    """Beregn optimal batch-størrelse med EOQ-formel (ukentlig lagerkost).

    Matematikk:
      - Ukentlig etterspørsel: W = annual_demand / 52
      - Lagerkost per enhet per uke: h = unit_cost * hold_pct / 100
      - Optimal batch: Q* = sqrt(2 * setup_cost * W / h)
    """
    if annual_demand <= 0 or setup_cost <= 0 or unit_cost <= 0:
        return 0.0
    W = annual_demand / 52.0
    h = unit_cost * hold_pct / 100.0
    if W <= 0 or h <= 0:
        return 0.0
    import math
    return math.sqrt(2.0 * setup_cost * W / h)


def analyse_fra_db(db_path: str, prod: str, hold_pct: float = 2.0):
    """Modus B: Analyser direkte fra SQLite-database (uten MILP).

    Bruker routing-batch fra CostCalculator samt EOQ-formelen.
    Uten historisk salg vises kun nåværende batch.
    """
    from kostberegning import CostCalculator, SqliteData

    data = SqliteData(db_path)
    calculator = CostCalculator(data)
    calculator.calculate_all()
    results = calculator.calculate_product_costs(prod)

    if not results:
        print(f"Produkt {prod} ikke funnet eller har ingen routing.")
        return

    r = results[0]

    print(f"PRODUKT: {prod} - {r.product_desc}")

    # Nåværende batch fra routing
    current_batch = 0.0
    for od in r.operation_details:
        if od.batch_size and od.batch_size > current_batch:
            current_batch = od.batch_size
    if current_batch <= 0:
        current_batch = 10000.0

    setup = _beregn_setup_per_batch(r.operation_details)
    # Samme definisjon som i MILP-motoren:
    # base_unit_cost = operasjonskost - biproduktverdi per enhet
    base = r.operation_cost - r.by_product_value

    hist = _hent_historisk_salg(prod, db_path)
    S = hist / 52.0 if hist > 0 else 0.0

    if S <= 0:
        print(f"  Nåværende batch:       {current_batch:>12,.0f} LM")
        print(f"  Setupkost per batch:   {setup:>12,.2f} kr")
        print(f"  Kost per enhet:        {base:>12.4f} kr/LM")
        print("\n  ⚠️  Ingen historisk salg funnet for produktet —")
        print("      optimal batch kan ikke beregnes uten etterspørselsdata.")
        return

    optimal_batch = _optimal_batch_eoq(hist, setup, base, hold_pct)

    W = hist / 52.0
    h = base * hold_pct / 100.0

    current_setup_per_unit = setup / current_batch if current_batch > 0 else 0.0
    current_weeks = current_batch / W if W > 0 else 0.0
    current_holding_per_unit = (current_batch * h) / (2.0 * W) if W > 0 else 0.0
    current_total_per_unit = current_setup_per_unit + current_holding_per_unit

    if optimal_batch > 0 and W > 0:
        optimal_setup_per_unit = setup / optimal_batch
        optimal_weeks = optimal_batch / W
        optimal_holding_per_unit = (optimal_batch * h) / (2.0 * W)
        optimal_total_per_unit = optimal_setup_per_unit + optimal_holding_per_unit
        savings_per_unit = current_total_per_unit - optimal_total_per_unit
        annual_savings = savings_per_unit * hist
    else:
        optimal_setup_per_unit = 0.0
        optimal_weeks = 0.0
        optimal_holding_per_unit = 0.0
        optimal_total_per_unit = 0.0
        savings_per_unit = 0.0
        annual_savings = 0.0

    print(f"  Nåværende batch:       {current_batch:>12,.0f} LM")
    print(f"  Optimal batch (EOQ):   {optimal_batch:>12,.0f} LM")
    print(f"  Årlig etterspørsel:    {hist:>12,.0f} LM/år")
    print(f"  Ukentlig etterspørsel: {W:>12,.0f} LM/uke")
    print(f"  Setupkost per batch:   {setup:>12,.2f} kr")
    print(f"  Nåværende setup/enh:   {current_setup_per_unit:>12.4f} kr/LM")
    print(f"  Optimal setup/enh:     {optimal_setup_per_unit:>12.4f} kr/LM")
    print(f"  Nåværende lager/enh:   {current_holding_per_unit:>12.4f} kr/LM")
    print(f"  Optimal lager/enh:     {optimal_holding_per_unit:>12.4f} kr/LM")
    print(f"  Nåværende total/enh:   {current_total_per_unit:>12.4f} kr/LM")
    print(f"  Optimal total/enh:     {optimal_total_per_unit:>12.4f} kr/LM")
    print(f"  Nåv. uker dekning:     {current_weeks:>12,.1f} uker")
    print(f"  Opt. uker dekning:     {optimal_weeks:>12,.1f} uker")
    print(f"  Besparelse/enh:        {savings_per_unit:>12.4f} kr/LM")
    print(f"  Årlig besparelse:      {annual_savings:>12,.0f} kr/år")

    _print_batch_tabell(setup, base, S, hold_pct, optimal_batch)


def main():
    parser = argparse.ArgumentParser(
        description="Analyser optimal batch-størrelse vs lagerkostnad for et produkt.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Eksempler:
  # Modus A: Fra MILP-resultat (etter å ha kjørt optimization_engine.py)
  python src/scripts/analyser_batch.py --milp resultat.json JD19148

  # Modus B: Direkte fra SQLite-database (uten MILP)
  python src/scripts/analyser_batch.py --db src/produksjonskalkyle.db JD19148
        """,
    )
    parser.add_argument("produkt", nargs="?", default="JD19098",
                        help="Produktnummer (Item No). Standard: JD19098")
    parser.add_argument("--milp", metavar="JSON", default=None,
                        help="Analyser fra MILP-resultat (resultat.json)")
    parser.add_argument("--db", metavar="DB", default=None,
                        help="Analyser direkte fra SQLite-database")
    parser.add_argument("--hold-pct", type=float, default=2.0,
                        help="Lagerholdskost i prosent av enhetskost per uke (standard: 2.0)")
    args = parser.parse_args()

    if args.milp:
        analyse_fra_milp(args.milp, args.produkt, args.hold_pct)
    elif args.db:
        analyse_fra_db(args.db, args.produkt, args.hold_pct)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()