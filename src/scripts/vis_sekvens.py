import json
import sys
from pathlib import Path

_root = Path(__file__).resolve().parent.parent.parent
d = json.load(open(_root / "resultat.json", encoding="utf-8"))
seq = d.get("sequence", [])
print(f"SEKVENS ANTALL: {len(seq)}")
for s in seq[:10]:
    print(f"  pos {s['pos']}: uke {s['uke']} {s['arbeidssenter']:<15} "
          f"{s['produkt']:<20} {s['familie']:<12} {s['kvantum']:>8,.0f} LM "
          f"changeover={s['changeover_minutter']}min")