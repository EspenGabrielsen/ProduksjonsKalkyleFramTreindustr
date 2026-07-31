import json, sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from kostberegning import ExcelData, CostCalculator, _serialize

_SRC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(_SRC_DIR, "Produksjonsmodell_Testdata_v3.xlsx")
BASELINE_KALKYLE = os.path.join(os.path.dirname(_SRC_DIR), "output", "baseline_kalkyle.json")

print("Laster data...")
data = ExcelData(EXCEL_PATH)

print("Kjører kalkyle...")
calculator = CostCalculator(data)
results = calculator.calculate_all()
ny = _serialize(results)

print(f"Generert {len(ny)} resultater")

with open(BASELINE_KALKYLE, 'r', encoding='utf-8') as f:
    b_raw = json.load(f)
# baseline har formatet {"results": [...]} eller rett liste
if isinstance(b_raw, dict) and 'results' in b_raw:
    b = b_raw['results']
else:
    b = b_raw
print(f"Baseline {len(b)} resultater")

# Finn forskjeller
forskjeller = 0
for i in range(min(len(b), len(ny))):
    b_str = json.dumps(b[i], sort_keys=True, ensure_ascii=False, default=str)
    ny_str = json.dumps(ny[i], sort_keys=True, ensure_ascii=False, default=str)
    if b_str != ny_str:
        forskjeller += 1
        if forskjeller <= 5:
            prod = b[i].get('product_no', '?')
            loc = b[i].get('location_code', '?')
            print(f'\nRow {i} ({prod}/{loc}) er forskjellig:')
            all_keys = sorted(set(list(b[i].keys()) + list(ny[i].keys())))
            for key in all_keys:
                if key == 'product_no' or key == 'location_code':
                    continue
                b_val = b[i].get(key)
                ny_val = ny[i].get(key)
                if str(b_val) != str(ny_val):
                    print(f'  {key}:')
                    print(f'    base: {str(b_val)[:300]}')
                    print(f'    new:  {str(ny_val)[:300]}')

print(f'\nTotalt: {forskjeller} av {len(b)} resultater er forskjellige')