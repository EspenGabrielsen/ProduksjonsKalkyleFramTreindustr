"""Analyser optimal batch-storrelse vs lagerkostnad for et produkt.

Matematikk:
  - Setup per enhet: setup / batch
  - Lagerkost per enhet: avhenger av hvor mange uker batchen dekker.
    Med ukentlig salg S og batch B dekker batchen T = B/S uker.
    Gjennomsnittlig lager = B/2 over hele perioden.
    Total lagerkost = (B/2) * T * enhetsverdi * (hold_pct/100 per uke)
    Per produsert enhet = total / B = (B * enhetsverdi * hold_pct) / (2 * S * 100)
"""
import json
import sys

def main():
    d = json.load(open("resultat.json", encoding="utf-8"))
    prod = sys.argv[1] if len(sys.argv) > 1 else "JD19098"

    rows = [r for r in d["production_plan"] if r["product"] == prod]
    if not rows:
        print(f"Produkt {prod} ikke funnet i produksjonsplanen.")
        return

    print(f"PRODUKT: {prod}")
    print(f"Antall produksjonslinjer: {len(rows)}")
    for r in rows:
        print(f"  uke {r['period']}: {r['quantity']:>10,.1f} LM  "
              f"batch={r.get('actual_batch_size')}  setup/enh={r.get('setup_per_unit')}  "
              f"hold/enh/u={r.get('holding_cost_per_unit')}  kost={r['total_cost']:.2f}")

    bc = [b for b in d["batch_decisions"] if b["product"] == prod]
    if not bc:
        print(f"\nIngen batch-decision for {prod}")
        return
    setup = bc[0]["setup_cost"]
    base = rows[0]["base_unit_cost"]
    hold_pct = 2.0
    hold_unit = base * hold_pct / 100.0  # kr per enhet per uke

    # Ukentlig salg: bruk aggregert historisk salg hvis mulig,
    # ellers anta batch/52
    hist = 0.0
    try:
        import sys as _sys
        _sys.path.insert(0, "src")
        from data_repo import DataRepo
        db = DataRepo("src/produksjonskalkyle.db")
        db.initialize()
        row = db.conn.execute(
            "SELECT SUM(quantity) as t FROM historical_sales WHERE product_id=?",
            (prod,)
        ).fetchone()
        hist = row["t"] or 0.0
    except Exception:
        hist = 0.0
    S = hist / 52.0 if hist > 0 else rows[0]["quantity"] / 52.0
    if S <= 0:
        S = 738.0  # fallback

    print(f"\nBATCH-STØRRELSE VS KOSTNAD (setup={setup:.2f} kr, verdi={base:.4f} kr/LM, "
          f"hold={hold_pct}%/uke, ukessalg~{S:.0f} LM):")
    print(f"  {'Batch LM':>12} {'Uker dk.':>8} {'Setup/enh':>10} {'Hold/enh':>10} "
          f"{'Total/enh':>10} {'Total kr':>12}")
    for batch in [1000, 2000, 5000, 10000, 20000, 25000, 38411, 50000, 75000]:
        setup_e = setup / batch
        T = batch / S  # antall uker batchen dekker
        hold_e = (T * base * hold_pct / 100.0) / 2.0  # per produsert enhet
        total_e = setup_e + hold_e
        total_kr = total_e * batch
        print(f"  {batch:>12,.0f} {T:>8,.1f} {setup_e:>10.4f} {hold_e:>10.4f} "
              f"{total_e:>10.4f} {total_kr:>12,.0f}")

if __name__ == "__main__":
    main()