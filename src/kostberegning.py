"""
kostberegning.py - Beregning av produktkost fra Excel-modell

Beregner:
  - Materialkost (fra BOM + Item Costs, inkl. svinn)
  - Operasjonskost (fra Routing + Work Centers)
  - Setupkost (fordelt på batch size)
  - Biproduktverdi (fra By Product Rules)
  - Netto produksjonskost

Bruk:
  python kostberegning.py --excel Produksjonsmodell_Testdata.xlsx
  python kostberegning.py --test
  python kostberegning.py --help
"""

import argparse
import json
import sys
import io


# Force UTF-8 for console output (Windows cp1252 workaround)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
elif hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd

# ──────────────────────────────────────────────────────────────────────
#  1. DATASTRUKTURER
# ──────────────────────────────────────────────────────────────────────

@dataclass
class Product:
    item_no: str
    description: str
    item_type: str
    product_group: str
    base_uom: str
    active: bool = True


@dataclass
class Location:
    code: str
    name: str
    location_type: str
    active: bool = True


@dataclass
class WorkCenter:
    code: str
    description: str
    location_code: str
    labor_cost_hour: float = 0.0
    machine_cost_hour: float = 0.0
    overhead_cost_hour: float = 0.0
    capacity_hours_day: float = 0.0
    effective_capacity_pct: float = 100.0
    active: bool = True

    @property
    def total_cost_hour(self) -> float:
        nominal_cost = self.labor_cost_hour + self.machine_cost_hour + self.overhead_cost_hour
        # Juster kostnad per time for å inkludere ståtid basert på effektiv kapasitet %
        if self.effective_capacity_pct > 0:
            return nominal_cost / (self.effective_capacity_pct / 100.0)
        return nominal_cost

    @property
    def effective_hours_day(self) -> float:
        return self.capacity_hours_day * (self.effective_capacity_pct / 100.0)


@dataclass
class Operation:
    code: str
    description: str
    default_work_center: str = ""
    standard_unit: str = "Minutes"
    active: bool = True


@dataclass
class ItemCost:
    item_no: str
    cost_type: str
    unit_cost: float = 0.0
    currency: str = "NOK"
    effective_date: Optional[date] = None


@dataclass
class BOMLine:
    parent_item_no: str
    component_item_no: str
    quantity_per: float = 1.0
    uom: str = ""
    scrap_pct: float = 0.0
    co_product_pct: float = 0.0
    co_product_item_no: str = ""
    valid_from: Optional[date] = None
    valid_to: Optional[date] = None


@dataclass
class RoutingLine:
    item_no: str
    operation_no: int
    operation_code: str
    work_center_code: str
    setup_time_minutes: float = 0.0
    run_time_minutes: float = 0.0
    batch_size: float = 1.0
    valid_from: Optional[date] = None
    valid_to: Optional[date] = None



@dataclass
class ByProductRule:
    parent_item_no: str
    by_product_item_no: str
    expected_quantity: float = 0.0
    uom: str = ""
    market_value: float = 0.0
    allocation_method: str = "Reduce Main Product Cost"


@dataclass
class CapacityDay:
    work_center: str
    date: date
    available_hours: float = 0.0
    planned_downtime: float = 0.0

    @property
    def available_production_hours(self) -> float:
        return self.available_hours - self.planned_downtime


@dataclass
class ProductionScenario:
    scenario_name: str
    product: str
    planned_quantity: float = 0.0
    start_date: Optional[date] = None
    end_date: Optional[date] = None


# ──────────────────────────────────────────────────────────────────────
#  2. RESULTATSTRUKTURER
# ──────────────────────────────────────────────────────────────────────

@dataclass
class MaterialCostDetail:
    component: str
    component_desc: str
    quantity_per: float
    scrap_pct: float
    unit_cost: float
    total_cost: float
    uom: str
    pct_of_gross: float = 0.0
    pct_of_material: float = 0.0


@dataclass
class OperationCostDetail:
    operation_no: int
    operation_desc: str
    work_center: str
    run_time_min: float
    setup_time_min: float
    batch_size: float
    cost_per_hour: float
    run_cost: float
    setup_cost_per_unit: float
    total_cost: float
    pct_of_gross: float = 0.0
    pct_of_operation: float = 0.0


@dataclass
class ByProductDetail:
    item_no: str
    description: str
    quantity: float
    uom: str
    market_value: float
    total_value: float
    allocation_method: str
    pct_of_gross: float = 0.0
    pct_of_byproduct: float = 0.0


@dataclass
class ProductCostResult:
    product_no: str
    product_desc: str
    product_group: str
    base_uom: str
    location_code: str = ""
    location_name: str = ""
    material_cost: float = 0.0
    operation_cost: float = 0.0
    setup_cost: float = 0.0
    gross_production_cost: float = 0.0
    by_product_value: float = 0.0
    net_production_cost: float = 0.0
    material_details: list[MaterialCostDetail] = field(default_factory=list)
    operation_details: list[OperationCostDetail] = field(default_factory=list)
    byproduct_details: list[ByProductDetail] = field(default_factory=list)
    cost_breakdown: list[dict] = field(default_factory=list)


@dataclass
class ScenarioResult:
    scenario_name: str
    product_no: str
    product_desc: str
    planned_quantity: float
    uom: str
    cost_per_unit: float
    total_material_cost: float = 0.0
    total_operation_cost: float = 0.0
    total_setup_cost: float = 0.0
    total_byproduct_value: float = 0.0
    total_net_cost: float = 0.0
    total_hours_needed: float = 0.0
    work_center_hours: dict = field(default_factory=dict)
    total_changeover_cost: float = 0.0
    total_changeover_hours: float = 0.0
    changeover_details: dict = field(default_factory=dict)



# ──────────────────────────────────────────────────────────────────────
#  3. EXCEL-LASTER
# ──────────────────────────────────────────────────────────────────────

class ExcelData:
    """Laster alle ark fra Excel-filen og gjør dem tilgjengelig som dict of DataFrames."""

    def __init__(self, filepath: str):
        self.filepath = Path(filepath)
        if not self.filepath.exists():
            raise FileNotFoundError(f"Finner ikke filen: {filepath}")

        self.raw: dict[str, pd.DataFrame] = {}
        self._load_all()

        # Strukturerte lister
        self.products: list[Product] = []
        self.locations: list[Location] = []
        self.work_centers: list[WorkCenter] = []
        self.operations: list[Operation] = []
        self.item_costs: list[ItemCost] = []
        self.bom_lines: list[BOMLine] = []
        self.routing_lines: list[RoutingLine] = []
        self.byproduct_rules: list[ByProductRule] = []
        self.capacity_days: list[CapacityDay] = []
        self.scenarios: list[ProductionScenario] = []

        # Indekser (dict) for raskt oppslag - bygges i _build_indexes()
        self._product_index: dict[str, Product] = {}
        self._location_index: dict[str, Location] = {}
        self._wc_index: dict[str, WorkCenter] = {}
        self._op_index: dict[str, Operation] = {}
        self._bom_index: dict[str, list[BOMLine]] = {}
        self._routing_index: dict[str, list[RoutingLine]] = {}
        self._byproduct_index: dict[str, list[ByProductRule]] = {}
        self._scenario_index: dict[str, ProductionScenario] = {}
        self._routing_locations_index: dict[str, list[str]] = {}
        self._item_cost_index: dict[str, list[ItemCost]] = {}

        self._parse_all()

    def _load_all(self):
        """Les alle ark fra Excel-filen."""
        xls = pd.ExcelFile(self.filepath)
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            # Fjern helt tomme rader
            df = df.dropna(how="all").reset_index(drop=True)
            self.raw[sheet] = df

    # ── Hjelpefunksjoner ──────────────────────────────────────────

    @staticmethod
    def _s(value) -> str:
        """Trygg streng-konvertering."""
        if pd.isna(value):
            return ""
        return str(value).strip()

    @staticmethod
    def _f(value) -> float:
        """Trygg float-konvertering."""
        if pd.isna(value):
            return 0.0
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    @staticmethod
    def _d(value) -> Optional[date]:
        """Trygg dato-konvertering."""
        if pd.isna(value):
            return None
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        return None

    @staticmethod
    def _b(value) -> bool:
        """Trygg bool-konvertering (Ja/Nei → True/False)."""
        s = ExcelData._s(value).lower()
        return s in ("ja", "yes", "true", "1", "y")

    # ── Parser for hvert ark ──────────────────────────────────────

    def _parse_products(self):
        df = self.raw.get("Product Master")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.products.append(Product(
                item_no=self._s(row.get("Item No", "")),
                description=self._s(row.get("Description", "")),
                item_type=self._s(row.get("Item Type", "")),
                product_group=self._s(row.get("Product Group", "")),
                base_uom=self._s(row.get("Base Unit of Measure", "")),
                active=self._b(row.get("Active", "Ja")),
            ))

    def _parse_locations(self):
        df = self.raw.get("Locations")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.locations.append(Location(
                code=self._s(row.get("Location Code", "")),
                name=self._s(row.get("Location Name", "")),
                location_type=self._s(row.get("Location Type", "")),
                active=self._b(row.get("Active", "Ja")),
            ))

    def _parse_work_centers(self):
        df = self.raw.get("Work Centers")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.work_centers.append(WorkCenter(
                code=self._s(row.get("Work Center Code", "")),
                description=self._s(row.get("Description", "")),
                location_code=self._s(row.get("Location Code", "")),
                labor_cost_hour=self._f(row.get("Labor Cost per Hour", 0)),
                machine_cost_hour=self._f(row.get("Machine Cost per Hour", 0)),
                overhead_cost_hour=self._f(row.get("Overhead Cost per Hour", 0)),
                capacity_hours_day=self._f(row.get("Capacity Hours per Day", 0)),
                effective_capacity_pct=self._f(row.get("Effective Capacity %", 100)),
                active=self._b(row.get("Active", "Ja")),
            ))

    def _parse_operations(self):
        df = self.raw.get("Operation Master")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.operations.append(Operation(
                code=self._s(row.get("Operation Code", "")),
                description=self._s(row.get("Description", "")),
                default_work_center=self._s(row.get("Default Work Center", "")),
                standard_unit=self._s(row.get("Standard Unit", "Minutes")),
                active=self._b(row.get("Active", "Ja")),
            ))

    def _parse_item_costs(self):
        df = self.raw.get("Item Costs")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.item_costs.append(ItemCost(
                item_no=self._s(row.get("Item No", "")),
                cost_type=self._s(row.get("Cost Type", "Standard Cost")),
                unit_cost=self._f(row.get("Unit Cost", 0)),
                currency=self._s(row.get("Currency", "NOK")),
                effective_date=self._d(row.get("Effective Date")),
            ))

    def _parse_bom(self):
        df = self.raw.get("BOM")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.bom_lines.append(BOMLine(
                parent_item_no=self._s(row.get("Parent Item No", "")),
                component_item_no=self._s(row.get("Component Item No", "")),
                quantity_per=self._f(row.get("Quantity Per", 1)),
                uom=self._s(row.get("Unit of Measure", "")),
                scrap_pct=self._f(row.get("Scrap %", 0)),
                co_product_pct=self._f(row.get("Co-Prod %", 0)),
                co_product_item_no=self._s(row.get("Co-Prod Item No", "")),
                valid_from=self._d(row.get("Valid From")),
                valid_to=self._d(row.get("Valid To")),
            ))

    def _parse_routing(self):
        df = self.raw.get("Routing")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.routing_lines.append(RoutingLine(
                item_no=self._s(row.get("Item No", "")),
                operation_no=int(self._f(row.get("Operation No", 0))),
                operation_code=self._s(row.get("Operation Code", "")),
                work_center_code=self._s(row.get("Work Center Code", "")),
                setup_time_minutes=self._f(row.get("Setup Time Minutes", 0)),
                run_time_minutes=self._f(row.get("Run Time Minutes", 0)),
                batch_size=self._f(row.get("Batch Size", 1)),
                valid_from=self._d(row.get("Valid From")),
                valid_to=self._d(row.get("Valid To")),
            ))

    def _parse_byproduct_rules(self):
        df = self.raw.get("By Product Rules")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.byproduct_rules.append(ByProductRule(
                parent_item_no=self._s(row.get("Parent Item No", "")),
                by_product_item_no=self._s(row.get("By Product Item No", "")),
                expected_quantity=self._f(row.get("Expected Quantity", 0)),
                uom=self._s(row.get("Unit of Measure", "")),
                market_value=self._f(row.get("Market Value", 0)),
                allocation_method=self._s(row.get("Allocation Method", "Reduce Main Product Cost")),
            ))

    def _parse_capacity(self):
        df = self.raw.get("Capacity Calendar")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.capacity_days.append(CapacityDay(
                work_center=self._s(row.get("Work Center", "")),
                date=self._d(row.get("Date")),
                available_hours=self._f(row.get("Available Hours", 0)),
                planned_downtime=self._f(row.get("Planned Downtime", 0)),
            ))

    def _parse_scenarios(self):
        df = self.raw.get("Production Scenario")
        if df is None or df.empty:
            return
        for row in df.to_dict('records'):
            self.scenarios.append(ProductionScenario(
                scenario_name=self._s(row.get("Scenario Name", "")),
                product=self._s(row.get("Product", "")),
                planned_quantity=self._f(row.get("Planned Quantity", 0)),
                start_date=self._d(row.get("Start Date")),
                end_date=self._d(row.get("End Date")),
            ))

    def _build_indexes(self):
        """Bygg dict-indekser for raskt oppslag."""
        self._product_index = {p.item_no: p for p in self.products}
        self._location_index = {l.code: l for l in self.locations}
        self._wc_index = {w.code: w for w in self.work_centers}
        self._op_index = {o.code: o for o in self.operations}
        
        # BOM: parent_item_no -> liste med BOMLine
        self._bom_index = {}
        for bl in self.bom_lines:
            self._bom_index.setdefault(bl.parent_item_no, []).append(bl)
        
        # Routing: item_no -> liste med RoutingLine (sortert på operation_no)
        self._routing_index = {}
        for rl in self.routing_lines:
            self._routing_index.setdefault(rl.item_no, []).append(rl)
        for item_no in self._routing_index:
            self._routing_index[item_no].sort(key=lambda r: r.operation_no)
        
        # Byproduct rules: parent_item_no -> liste med ByProductRule
        self._byproduct_index = {}
        for br in self.byproduct_rules:
            self._byproduct_index.setdefault(br.parent_item_no, []).append(br)
        
        # Scenarios: scenario_name -> ProductionScenario
        self._scenario_index = {s.scenario_name: s for s in self.scenarios}
        
        # Item costs: item_no -> liste med ItemCost (pre-sortert på effective_date, nyeste først)
        self._item_cost_index = {}
        for ic in self.item_costs:
            self._item_cost_index.setdefault(ic.item_no, []).append(ic)
        for _item_no in self._item_cost_index:
            self._item_cost_index[_item_no].sort(
                key=lambda c: c.effective_date if c.effective_date is not None else date.min,
                reverse=True
            )
        
        # Routing locations: item_no -> sortert liste med unike location codes
        _temp_locs: dict[str, set[str]] = {}
        for rl in self.routing_lines:
            wc = self._wc_index.get(rl.work_center_code)
            if wc and wc.location_code:
                _temp_locs.setdefault(rl.item_no, set()).add(wc.location_code)
        self._routing_locations_index = {
            item_no: sorted(locs) for item_no, locs in _temp_locs.items()
        }

    def _parse_all(self):
        self._parse_products()
        self._parse_locations()
        self._parse_work_centers()
        self._parse_operations()
        self._parse_item_costs()
        self._parse_bom()
        self._parse_routing()
        self._parse_byproduct_rules()
        self._parse_capacity()
        self._parse_scenarios()
        # Valider alle kryss-referanser etter at alt er lastet
        self._validate()
        # Bygg indekser for raskt oppslag
        self._build_indexes()

    def _validate(self):
        """Valider alle kryss-referanser i dataene.
        
        Kaster ValueError med detaljert feilmelding hvis noe ikke stemmer.
        """
        feil = []

        # Hjelpesett for raskt oppslag
        product_set = {p.item_no for p in self.products}
        location_set = {l.code for l in self.locations}
        wc_set = {w.code for w in self.work_centers}
        op_set = {o.code for o in self.operations}

        # 1. Work Centers -> Locations
        for i, wc in enumerate(self.work_centers, 1):
            if wc.location_code and wc.location_code not in location_set:
                feil.append(f"Work Centers rad {i}: Location Code '{wc.location_code}' finnes ikke i Locations-arket")

        # 2. Operation Master -> Work Centers
        for i, op in enumerate(self.operations, 1):
            if op.default_work_center and op.default_work_center not in wc_set:
                feil.append(f"Operation Master rad {i}: Default Work Center '{op.default_work_center}' finnes ikke i Work Centers-arket")

        # 3. Item Costs -> Product Master
        for i, ic in enumerate(self.item_costs, 1):
            if ic.item_no not in product_set:
                feil.append(f"Item Costs rad {i}: Item No '{ic.item_no}' finnes ikke i Product Master-arket")

        # 4. BOM -> Product Master (Parent og Component)
        for i, bl in enumerate(self.bom_lines, 1):
            if bl.parent_item_no not in product_set:
                feil.append(f"BOM rad {i}: Parent Item No '{bl.parent_item_no}' finnes ikke i Product Master-arket")
            if bl.component_item_no not in product_set:
                feil.append(f"BOM rad {i}: Component Item No '{bl.component_item_no}' finnes ikke i Product Master-arket")
            if bl.co_product_item_no and bl.co_product_item_no not in product_set:
                feil.append(f"BOM rad {i}: Co-Prod Item No '{bl.co_product_item_no}' finnes ikke i Product Master-arket")

        # 5. Routing -> Product Master, Operation Master, Work Centers
        for i, rl in enumerate(self.routing_lines, 1):
            if rl.item_no not in product_set:
                feil.append(f"Routing rad {i}: Item No '{rl.item_no}' finnes ikke i Product Master-arket")
            if rl.operation_code not in op_set:
                feil.append(f"Routing rad {i}: Operation Code '{rl.operation_code}' finnes ikke i Operation Master-arket")
            if rl.work_center_code not in wc_set:
                feil.append(f"Routing rad {i}: Work Center Code '{rl.work_center_code}' finnes ikke i Work Centers-arket")

        # 6. By Product Rules -> Product Master
        for i, br in enumerate(self.byproduct_rules, 1):
            if br.parent_item_no not in product_set:
                feil.append(f"By Product Rules rad {i}: Parent Item No '{br.parent_item_no}' finnes ikke i Product Master-arket")
            if br.by_product_item_no not in product_set:
                feil.append(f"By Product Rules rad {i}: By Product Item No '{br.by_product_item_no}' finnes ikke i Product Master-arket")

        # 7. Capacity Calendar -> Work Centers
        for i, cd in enumerate(self.capacity_days, 1):
            if cd.work_center not in wc_set:
                feil.append(f"Capacity Calendar rad {i}: Work Center '{cd.work_center}' finnes ikke i Work Centers-arket")

        # 8. Production Scenario -> Product Master
        for i, sc in enumerate(self.scenarios, 1):
            if sc.product not in product_set:
                feil.append(f"Production Scenario rad {i}: Product '{sc.product}' finnes ikke i Product Master-arket")

        if feil:
            raise ValueError("Valideringsfeil i Excel-data:\n" + "\n".join(f"  - {f}" for f in feil))

    # ── Oppslagsverk (bruker dict-indekser) ──────────────────────

    def product(self, item_no: str) -> Optional[Product]:
        return self._product_index.get(item_no)

    def work_center(self, code: str) -> Optional[WorkCenter]:
        return self._wc_index.get(code)

    def operation(self, code: str) -> Optional[Operation]:
        return self._op_index.get(code)

    def item_cost(self, item_no: str, cost_type: str = "Standard Cost") -> Optional[ItemCost]:
        """Finn nyeste kost av angitt type for en vare."""
        matches = self._item_cost_index.get(item_no, [])
        if not matches:
            return None
        # Prøv først spesifikk cost_type
        exact = [c for c in matches if c.cost_type == cost_type]
        if exact:
            matches = exact
        matches.sort(key=lambda c: c.effective_date if c.effective_date is not None else date.min, reverse=True)
        return matches[0]

    def bom_for(self, item_no: str) -> list[BOMLine]:
        return self._bom_index.get(item_no, [])

    def routing_for(self, item_no: str) -> list[RoutingLine]:
        return self._routing_index.get(item_no, [])

    def byproduct_rules_for(self, parent_item_no: str) -> list[ByProductRule]:
        return self._byproduct_index.get(parent_item_no, [])

    def location(self, code: str) -> Optional[Location]:
        return self._location_index.get(code)

    def routing_locations_for(self, item_no: str) -> list[str]:
        """Hent unike location codes for et produkt basert pa routing."""
        return self._routing_locations_index.get(item_no, [])

    def scenario(self, name: str) -> Optional[ProductionScenario]:
        return self._scenario_index.get(name)


# ──────────────────────────────────────────────────────────────────────
#  3B. SQLite-LASTER (parallell med ExcelData)
# ──────────────────────────────────────────────────────────────────────

class SqliteData:
    """Laster data fra SQLite i stedet for Excel.
    
    Har samme interface som ExcelData, slik at CostCalculator og
    SimulationEngine kan bruke begge kilder uten endringer.
    """

    def __init__(self, db_path: Optional[str] = None):
        # Importer DataRepo her for å unngå sirkulære imports
        from data_repo import DataRepo
        
        self.db = DataRepo(db_path)
        self.db.initialize()
        self.filepath = Path(self.db.db_path)
        self.raw: dict[str, pd.DataFrame] = {}

        # Strukturerte lister (samme som ExcelData)
        self.products: list[Product] = []
        self.locations: list[Location] = []
        self.work_centers: list[WorkCenter] = []
        self.operations: list[Operation] = []
        self.item_costs: list[ItemCost] = []
        self.bom_lines: list[BOMLine] = []
        self.routing_lines: list[RoutingLine] = []
        self.byproduct_rules: list[ByProductRule] = []
        self.capacity_days: list[CapacityDay] = []
        self.scenarios: list[ProductionScenario] = []

        # Indekser (samme som ExcelData)
        self._product_index: dict[str, Product] = {}
        self._location_index: dict[str, Location] = {}
        self._wc_index: dict[str, WorkCenter] = {}
        self._op_index: dict[str, Operation] = {}
        self._bom_index: dict[str, list[BOMLine]] = {}
        self._routing_index: dict[str, list[RoutingLine]] = {}
        self._byproduct_index: dict[str, list[ByProductRule]] = {}
        self._scenario_index: dict[str, ProductionScenario] = {}
        self._routing_locations_index: dict[str, list[str]] = {}
        self._item_cost_index: dict[str, list[ItemCost]] = {}

        self._load_all()

    def _load_all(self):
        """Les alle data fra SQLite."""
        self._load_products()
        self._load_locations()
        self._load_work_centers()
        self._load_operations()
        self._load_item_costs()
        self._load_bom()
        self._load_routing()
        self._load_byproduct_rules()
        self._load_capacity()
        self._load_scenarios()
        self._build_indexes()

    def _load_products(self):
        rows = self.db.conn.execute(
            "SELECT item_no, description, item_type, product_group, base_uom FROM products ORDER BY item_no"
        ).fetchall()
        for r in rows:
            self.products.append(Product(
                item_no=r["item_no"],
                description=r["description"],
                item_type=r["item_type"],
                product_group=r["product_group"],
                base_uom=r["base_uom"],
            ))

    def _load_locations(self):
        rows = self.db.conn.execute(
            "SELECT code, name, location_type FROM locations ORDER BY code"
        ).fetchall()
        for r in rows:
            self.locations.append(Location(
                code=r["code"],
                name=r["name"],
                location_type=r["location_type"],
            ))

    def _load_work_centers(self):
        rows = self.db.conn.execute(
            """SELECT code, description, location_code, labor_cost_hour,
                      machine_cost_hour, overhead_cost_hour, capacity_hours_day,
                      effective_capacity_pct
               FROM work_centers ORDER BY code"""
        ).fetchall()
        for r in rows:
            self.work_centers.append(WorkCenter(
                code=r["code"],
                description=r["description"],
                location_code=r["location_code"],
                labor_cost_hour=r["labor_cost_hour"],
                machine_cost_hour=r["machine_cost_hour"],
                overhead_cost_hour=r["overhead_cost_hour"],
                capacity_hours_day=r["capacity_hours_day"],
                effective_capacity_pct=r["effective_capacity_pct"],
            ))

    def _load_operations(self):
        rows = self.db.conn.execute(
            "SELECT code, description, default_work_center, standard_unit FROM operations ORDER BY code"
        ).fetchall()
        for r in rows:
            self.operations.append(Operation(
                code=r["code"],
                description=r["description"],
                default_work_center=r["default_work_center"],
                standard_unit=r["standard_unit"],
            ))

    def _load_item_costs(self):
        rows = self.db.conn.execute(
            "SELECT item_no, cost_type, unit_cost, currency, effective_date FROM item_costs ORDER BY item_no"
        ).fetchall()
        for r in rows:
            ed = None
            if r["effective_date"]:
                try:
                    ed = date.fromisoformat(r["effective_date"])
                except (ValueError, TypeError):
                    ed = None
            self.item_costs.append(ItemCost(
                item_no=r["item_no"],
                cost_type=r["cost_type"],
                unit_cost=r["unit_cost"],
                currency=r["currency"],
                effective_date=ed,
            ))

    def _load_bom(self):
        rows = self.db.conn.execute(
            """SELECT parent_item_no, component_item_no, quantity_per, uom,
                      scrap_pct, co_product_pct, co_product_item_no
               FROM bom_lines ORDER BY parent_item_no, component_item_no"""
        ).fetchall()
        for r in rows:
            self.bom_lines.append(BOMLine(
                parent_item_no=r["parent_item_no"],
                component_item_no=r["component_item_no"],
                quantity_per=r["quantity_per"],
                uom=r["uom"],
                scrap_pct=r["scrap_pct"],
                co_product_pct=r["co_product_pct"],
                co_product_item_no=r["co_product_item_no"],
            ))

    def _load_routing(self):
        rows = self.db.conn.execute(
            """SELECT item_no, operation_no, operation_code, work_center_code,
                      setup_time_minutes, run_time_minutes, batch_size
               FROM routing_lines ORDER BY item_no, operation_no"""
        ).fetchall()
        for r in rows:
            self.routing_lines.append(RoutingLine(
                item_no=r["item_no"],
                operation_no=r["operation_no"],
                operation_code=r["operation_code"],
                work_center_code=r["work_center_code"],
                setup_time_minutes=r["setup_time_minutes"],
                run_time_minutes=r["run_time_minutes"],
                batch_size=r["batch_size"],
            ))

    def _load_byproduct_rules(self):
        rows = self.db.conn.execute(
            """SELECT parent_item_no, by_product_item_no, expected_quantity, uom,
                      market_value, allocation_method
               FROM byproduct_rules ORDER BY parent_item_no, by_product_item_no"""
        ).fetchall()
        for r in rows:
            self.byproduct_rules.append(ByProductRule(
                parent_item_no=r["parent_item_no"],
                by_product_item_no=r["by_product_item_no"],
                expected_quantity=r["expected_quantity"],
                uom=r["uom"],
                market_value=r["market_value"],
                allocation_method=r["allocation_method"],
            ))

    def _load_capacity(self):
        rows = self.db.conn.execute(
            """SELECT work_center, date, available_hours, planned_downtime
               FROM capacity_days ORDER BY work_center, date"""
        ).fetchall()
        for r in rows:
            d = None
            if r["date"]:
                try:
                    d = date.fromisoformat(r["date"])
                except (ValueError, TypeError):
                    d = None
            self.capacity_days.append(CapacityDay(
                work_center=r["work_center"],
                date=d,
                available_hours=r["available_hours"],
                planned_downtime=r["planned_downtime"],
            ))

    def _load_scenarios(self):
        rows = self.db.conn.execute(
            "SELECT scenario_name, product, planned_quantity FROM production_scenarios ORDER BY scenario_name"
        ).fetchall()
        for r in rows:
            self.scenarios.append(ProductionScenario(
                scenario_name=r["scenario_name"],
                product=r["product"],
                planned_quantity=r["planned_quantity"],
            ))

    def _build_indexes(self):
        """Bygg dict-indekser for raskt oppslag (samme som ExcelData)."""
        self._product_index = {p.item_no: p for p in self.products}
        self._location_index = {l.code: l for l in self.locations}
        self._wc_index = {w.code: w for w in self.work_centers}
        self._op_index = {o.code: o for o in self.operations}

        # BOM: parent_item_no -> liste med BOMLine
        self._bom_index = {}
        for bl in self.bom_lines:
            self._bom_index.setdefault(bl.parent_item_no, []).append(bl)

        # Routing: item_no -> liste med RoutingLine (sortert på operation_no)
        self._routing_index = {}
        for rl in self.routing_lines:
            self._routing_index.setdefault(rl.item_no, []).append(rl)
        for item_no in self._routing_index:
            self._routing_index[item_no].sort(key=lambda r: r.operation_no)

        # Byproduct rules: parent_item_no -> liste med ByProductRule
        self._byproduct_index = {}
        for br in self.byproduct_rules:
            self._byproduct_index.setdefault(br.parent_item_no, []).append(br)

        # Scenarios: scenario_name -> ProductionScenario
        self._scenario_index = {s.scenario_name: s for s in self.scenarios}

        # Item costs: item_no -> liste med ItemCost
        self._item_cost_index = {}
        for ic in self.item_costs:
            self._item_cost_index.setdefault(ic.item_no, []).append(ic)

        # Routing locations: item_no -> sortert liste med unike location codes
        _temp_locs: dict[str, set[str]] = {}
        for rl in self.routing_lines:
            wc = self._wc_index.get(rl.work_center_code)
            if wc and wc.location_code:
                _temp_locs.setdefault(rl.item_no, set()).add(wc.location_code)
        self._routing_locations_index = {
            item_no: sorted(locs) for item_no, locs in _temp_locs.items()
        }

    # ── Oppslagsverk (samme interface som ExcelData) ─────────────

    def product(self, item_no: str) -> Optional[Product]:
        return self._product_index.get(item_no)

    def work_center(self, code: str) -> Optional[WorkCenter]:
        return self._wc_index.get(code)

    def operation(self, code: str) -> Optional[Operation]:
        return self._op_index.get(code)

    def item_cost(self, item_no: str, cost_type: str = "Standard Cost") -> Optional[ItemCost]:
        """Finn nyeste kost av angitt type for en vare."""
        matches = self._item_cost_index.get(item_no, [])
        if not matches:
            return None
        exact = [c for c in matches if c.cost_type == cost_type]
        if exact:
            matches = exact
        matches.sort(key=lambda c: c.effective_date if c.effective_date is not None else date.min, reverse=True)
        return matches[0]

    def bom_for(self, item_no: str) -> list[BOMLine]:
        return self._bom_index.get(item_no, [])

    def routing_for(self, item_no: str) -> list[RoutingLine]:
        return self._routing_index.get(item_no, [])

    def byproduct_rules_for(self, parent_item_no: str) -> list[ByProductRule]:
        return self._byproduct_index.get(parent_item_no, [])

    def location(self, code: str) -> Optional[Location]:
        return self._location_index.get(code)

    def routing_locations_for(self, item_no: str) -> list[str]:
        return self._routing_locations_index.get(item_no, [])

    def scenario(self, name: str) -> Optional[ProductionScenario]:
        return self._scenario_index.get(name)


# ──────────────────────────────────────────────────────────────────────
#  4. KOSTNADSBEREGNER
# ──────────────────────────────────────────────────────────────────────

class CostCalculator:
    """Beregner kost per produkt basert på data fra ExcelData.
    
    Stotter dynamisk cost roll-up: hvis en komponent i BOM er et
    ferdigvare/halvfabrikat som allerede er beregnet, brukes den
    dynamisk beregnede brutto produksjonskosten i stedet for den
    statiske verdien fra Item Costs-arket.
    """

    def __init__(self, data: ExcelData):
        self.data = data
        # Cache for dynamisk beregnede kostnader: {item_no: {location_code: gross_production_cost}}
        self.calculated_costs: dict[str, dict[str, float]] = {}

    def _calc_percentages(self, result: ProductCostResult):
        """Beregn prosentandeler for alle detaljer og lag cost_breakdown."""
        gross = result.gross_production_cost
        if gross == 0:
            return

        # ── Toppnivå cost_breakdown ──────────────────────────────
        result.cost_breakdown = [
            {"category": "Materialkost",     "amount": round(result.material_cost, 4),
             "pct_of_gross": round(result.material_cost / gross * 100, 1)},
            {"category": "Operasjonskost",   "amount": round(result.operation_cost, 4),
             "pct_of_gross": round(result.operation_cost / gross * 100, 1)},
            {"category": "Setupkost",        "amount": round(result.setup_cost, 4),
             "pct_of_gross": round(result.setup_cost / gross * 100, 1)},
        ]
        if result.by_product_value > 0:
            result.cost_breakdown.append(
                {"category": "Biproduktverdi", "amount": -round(result.by_product_value, 4),
                 "pct_of_gross": -round(result.by_product_value / gross * 100, 1)}
            )
        result.cost_breakdown.append(
            {"category": "Netto kost",       "amount": round(result.net_production_cost, 4),
             "pct_of_gross": round(result.net_production_cost / gross * 100, 1)}
        )

        # ── Materialdetaljer ─────────────────────────────────────
        mat_total = result.material_cost
        for md in result.material_details:
            md.pct_of_gross = round(md.total_cost / gross * 100, 1)
            md.pct_of_material = round(md.total_cost / mat_total * 100, 1) if mat_total else 0.0

        # ── Operasjonsdetaljer ───────────────────────────────────
        op_total = result.operation_cost + result.setup_cost
        for od in result.operation_details:
            od.pct_of_gross = round(od.total_cost / gross * 100, 1)
            od.pct_of_operation = round(od.total_cost / op_total * 100, 1) if op_total else 0.0

        # ── Biproduktdetaljer ────────────────────────────────────
        bp_total = result.by_product_value
        for bd in result.byproduct_details:
            bd.pct_of_gross = round(bd.total_value / gross * 100, 1)
            bd.pct_of_byproduct = round(bd.total_value / bp_total * 100, 1) if bp_total else 0.0

    def _calc_product_cost_for_location(self, item_no: str, location_code: str) -> Optional[ProductCostResult]:
        """Beregn full kostnad for ett produkt pa en spesifikk fabrikk."""
        prod = self.data.product(item_no)
        if prod is None:
            return None

        loc = self.data.location(location_code)
        result = ProductCostResult(
            product_no=prod.item_no,
            product_desc=prod.description,
            product_group=prod.product_group,
            base_uom=prod.base_uom,
            location_code=location_code,
            location_name=loc.name if loc else "",
        )

        # 1. Materialkost (fra BOM) - felles uavhengig av fabrikk
        self._calc_material_cost(result)

        # 2. Operasjonskost (fra Routing) - filtrert pa fabrikk
        self._calc_operation_cost(result, location_code)

        # 3. Brutto produksjonskost
        result.gross_production_cost = (
            result.material_cost + result.operation_cost + result.setup_cost
        )

        # 4. Biproduktverdi (fra By Product Rules)
        self._calc_byproduct_value(result)

        # 5. Netto produksjonskost
        result.net_production_cost = result.gross_production_cost - result.by_product_value

        # 6. Beregn prosentandeler og cost_breakdown
        self._calc_percentages(result)

        return result

    def calculate_product_costs(self, item_no: str) -> list[ProductCostResult]:
        """Beregn kost for ett produkt pa alle fabrikker som har routing.
        
        Returnerer en liste med ett resultat per fabrikk.
        """
        prod = self.data.product(item_no)
        if prod is None:
            return []

        # Finn alle fabrikker som har routing for dette produktet
        location_codes = self.data.routing_locations_for(item_no)
        if not location_codes:
            return []

        results = []
        for loc_code in location_codes:
            r = self._calc_product_cost_for_location(item_no, loc_code)
            if r is not None:
                results.append(r)
        return results

    def calculate_product_cost(self, item_no: str) -> Optional[ProductCostResult]:
        """Beregn full kostnad for ett produkt (forste fabrikk).
        
        Beholdes for bakoverkompatibilitet. Returnerer resultatet for
        forste fabrikk som har routing for produktet.
        """
        results = self.calculate_product_costs(item_no)
        return results[0] if results else None

    def _calc_material_cost(self, result: ProductCostResult):
        """
        Beregn materialkost fra BOM med dynamisk cost roll-up.

        Quantity Per = antall output-enheter per input-enhet.
        Materialkost = Unit Cost / Quantity Per * (1 + Scrap%)

        Hvis komponenten er et ferdigvare/halvfabrikat som allerede er
        beregnet dynamisk, brukes den dynamisk beregnede brutto produksjonskosten
        i stedet for den statiske verdien fra Item Costs-arket.

        Co-Prod % pavirker IKKE materialkosten per enhet for hovedproduktet.
        Materialkost per LM er den samme uansett om co-produkt produseres eller ikke.
        Co-produktet far sin egen materialkost i _calc_co_product_results.
        """
        bom_lines = self.data.bom_for(result.product_no)
        if not bom_lines:
            return

        total = 0.0
        for bl in bom_lines:
            comp = self.data.product(bl.component_item_no)
            
            # Dynamisk cost roll-up: sjekk om komponenten er beregnet
            if bl.component_item_no in self.calculated_costs:
                loc_costs = self.calculated_costs[bl.component_item_no]
                if isinstance(loc_costs, dict):
                    # Use cost for the same location if available, otherwise first available
                    unit_cost = loc_costs.get(result.location_code, next(iter(loc_costs.values()), 0.0))
                else:
                    unit_cost = loc_costs
            else:
                # Fallback til statisk verdi fra Excel (for ravarer)
                cost = self.data.item_cost(bl.component_item_no)
                unit_cost = cost.unit_cost if cost else 0.0

            # Quantity Per = output-enheter per input-enhet
            # Forbruk per output = 1 / Quantity Per
            qty = bl.quantity_per if bl.quantity_per > 0 else 1.0
            
            # Materialkost per output-enhet beregnes UTEN co-prod-justering.
            # Co-Prod % pavirker KUN operasjonstiden (mer tid per LM),
            # IKKE materialforbruket per LM.
            consumption_per_unit = 1.0 / qty
                
            consumption_with_scrap = consumption_per_unit * (1 + bl.scrap_pct / 100.0)
            line_cost = consumption_with_scrap * unit_cost
            total += line_cost

            result.material_details.append(MaterialCostDetail(
                component=bl.component_item_no,
                component_desc=comp.description if comp else "",
                quantity_per=bl.quantity_per,
                scrap_pct=bl.scrap_pct,
                unit_cost=unit_cost,
                total_cost=round(line_cost, 4),
                uom=bl.uom,
            ))

        result.material_cost = round(total, 4)


    def _calc_operation_cost(self, result: ProductCostResult, location_code: str = ""):
        """Beregn operasjonskost og setupkost fra Routing.
        
        Hvis location_code er angitt, filtreres routing-linjer til kun
        de arbeidssentrene som tilhorer den angitte fabrikken.
        
        Co-Prod %: hvis produktet har BOM-linjer med co_product_pct > 0,
        oker run time marginalt fordi maskinen ma handtere ogsa
        co-produktet. Run time justeres med (1 + co_pct).
        Setup time justeres IKKE - omstilling av maskin er den samme
        uansett om det blir A- eller B-vare. Kvalitetsfordeling (A/B)
        avhenger av ravarekvalitet, ikke maskininnstillinger.
        """
        routing = self.data.routing_for(result.product_no)
        if not routing:
            return

        # Filtrer pa fabrikk hvis location_code er angitt
        if location_code:
            routing = [
                r for r in routing
                if self.data.work_center(r.work_center_code)
                and self.data.work_center(r.work_center_code).location_code == location_code
            ]
            if not routing:
                return

        # Sjekk om produktet har co-produkt (Co-Prod % > 0)
        bom_lines = self.data.bom_for(result.product_no)
        co_pct = 0.0
        for bl in bom_lines:
            if bl.co_product_pct > 0:
                co_pct = bl.co_product_pct / 100.0
                break
        
        # Co-produkt justering: operasjonstiden oker marginalt
        # fordi maskinen ma handtere ogsa co-produktet
        co_factor = 1.0 + co_pct  # F.eks. 6% co-prod => 1.06

        total_run = 0.0
        total_setup = 0.0

        for rl in routing:
            wc = self.data.work_center(rl.work_center_code)
            op = self.data.operation(rl.operation_code)
            cost_per_hour = wc.total_cost_hour if wc else 0.0

            # Kjøretid per enhet (i timer) - justert for co-produkt
            run_hours = (rl.run_time_minutes / 60.0) * co_factor
            run_cost = run_hours * cost_per_hour

            # Setupkost per enhet - IKKE justert for co-produkt.
            # Omstilling av maskin er den samme uansett om det blir A- eller B-vare.
            # Kvalitetsfordeling (A/B) avhenger av råvarekvalitet, ikke maskininnstillinger.
            setup_hours = rl.setup_time_minutes / 60.0
            batch = rl.batch_size if rl.batch_size > 0 else 1
            setup_cost_per_unit = (setup_hours * cost_per_hour) / batch

            total_run += run_cost
            total_setup += setup_cost_per_unit

            result.operation_details.append(OperationCostDetail(
                operation_no=rl.operation_no,
                operation_desc=op.description if op else rl.operation_code,
                work_center=rl.work_center_code,
                run_time_min=rl.run_time_minutes * co_factor,
                setup_time_min=rl.setup_time_minutes,  # IKKE justert for co-prod
                batch_size=rl.batch_size,
                cost_per_hour=cost_per_hour,
                run_cost=round(run_cost, 4),
                setup_cost_per_unit=round(setup_cost_per_unit, 4),
                total_cost=round(run_cost + setup_cost_per_unit, 4),
            ))

        result.operation_cost = round(total_run, 4)
        result.setup_cost = round(total_setup, 4)

    def _calc_byproduct_value(self, result: ProductCostResult):
        """Beregn verdi av biprodukter knyttet til dette produktet."""
        rules = self.data.byproduct_rules_for(result.product_no)
        total_byproduct_value = 0.0
        for rule in rules:
            bp = self.data.product(rule.by_product_item_no)
            value = rule.expected_quantity * rule.market_value
            total_byproduct_value += value

            result.byproduct_details.append(ByProductDetail(
                item_no=rule.by_product_item_no,
                description=bp.description if bp else "",
                quantity=rule.expected_quantity,
                uom=rule.uom,
                market_value=rule.market_value,
                total_value=round(value, 4),
                allocation_method=rule.allocation_method,
            ))

        result.by_product_value = round(total_byproduct_value, 4)

    def _calc_co_product_results(self, parent_result: ProductCostResult) -> list[ProductCostResult]:
        """Generer resultater for co-produkter (f.eks. B-vare) basert pa BOM-linjer med co_product_pct.
        
        Nar en BOM-linje har co_product_pct > 0, betyr det at en andel av
        produksjonen gar til et co-produkt (f.eks. B-vare). Co-produktet
        skal ha sin egen kalkyle med allokert materialkost og operasjonskost.
        
        Materialkost: fordeles basert pa fysisk andel (Physical Allocation).
        Operasjonskost: fordeles proporsjonalt med co_product_pct.
        """
        co_results = []
        
        # Sjekk om dette produktet har BOM-linjer med co_product_pct
        bom_lines = self.data.bom_for(parent_result.product_no)
        has_co_product = any(bl.co_product_pct > 0 for bl in bom_lines)
        if not has_co_product:
            return co_results
        
        for bl in bom_lines:
            if bl.co_product_pct <= 0:
                continue
            
            co_item_no = bl.co_product_item_no
            if not co_item_no:
                continue
            
            co_prod = self.data.product(co_item_no)
            if co_prod is None:
                continue
            
            co_pct = bl.co_product_pct / 100.0  # 6% -> 0.06
            
            # Totalt utbytte = quantity_per / (1 - co_pct)
            total_qty = bl.quantity_per / (1.0 - co_pct)
            co_qty = total_qty - bl.quantity_per  # B-vare kvantum per input-enhet
            
            # Opprett co-produkt resultat
            co_result = ProductCostResult(
                product_no=co_prod.item_no,
                product_desc=co_prod.description,
                product_group=co_prod.product_group,
                base_uom=co_prod.base_uom,
                location_code=parent_result.location_code,
                location_name=parent_result.location_name,
            )
            
            # 1. Materialkost for co-produkt
            # Samme materialkost per enhet som hovedproduktet (felles ravare).
            # Materialkost per output-enhet = unit_cost / quantity_per (UTEN co-prod-justering)
            # Co-produktet far samme materialkost per LM som hovedproduktet.
            comp = self.data.product(bl.component_item_no)
            if bl.component_item_no in self.calculated_costs:
                loc_costs = self.calculated_costs[bl.component_item_no]
                if isinstance(loc_costs, dict):
                    unit_cost = loc_costs.get(parent_result.location_code, next(iter(loc_costs.values()), 0.0))
                else:
                    unit_cost = loc_costs
            else:
                cost = self.data.item_cost(bl.component_item_no)
                unit_cost = cost.unit_cost if cost else 0.0
            
            # Materialkost per enhet beregnes UTEN co-prod-justering
            # (samme som hovedproduktet)
            consumption_per_unit = 1.0 / bl.quantity_per
            consumption_with_scrap = consumption_per_unit * (1 + bl.scrap_pct / 100.0)
            material_cost_per_unit = consumption_with_scrap * unit_cost
            
            co_result.material_cost = round(material_cost_per_unit, 4)
            co_result.material_details.append(MaterialCostDetail(
                component=bl.component_item_no,
                component_desc=comp.description if comp else "",
                quantity_per=bl.quantity_per,
                scrap_pct=bl.scrap_pct,
                unit_cost=unit_cost,
                total_cost=round(material_cost_per_unit, 4),
                uom=bl.uom,
            ))
            
            # 2. Operasjonskost for co-produkt
            # Samme operasjoner som hovedproduktet, fordelt proporsjonalt med co_pct
            # Co-produktet har ikke egne operasjoner - det oppstar som en andel av
            # hovedproduktets produksjon. Operasjonskostnaden allokeres basert pa
            # co-produktets andel av total produksjon.
            routing = self.data.routing_for(parent_result.product_no)
            if parent_result.location_code:
                routing = [
                    r for r in routing
                    if self.data.work_center(r.work_center_code)
                    and self.data.work_center(r.work_center_code).location_code == parent_result.location_code
                ]
            
            total_run = 0.0
            total_setup = 0.0
            for rl in routing:
                wc = self.data.work_center(rl.work_center_code)
                op = self.data.operation(rl.operation_code)
                cost_per_hour = wc.total_cost_hour if wc else 0.0
                
                run_hours = rl.run_time_minutes / 60.0
                run_cost = run_hours * cost_per_hour * co_pct  # Allokert andel
                
                setup_hours = rl.setup_time_minutes / 60.0
                batch = rl.batch_size if rl.batch_size > 0 else 1
                setup_cost_per_unit = (setup_hours * cost_per_hour) / batch * co_pct  # Allokert andel
                
                total_run += run_cost
                total_setup += setup_cost_per_unit
                
                co_result.operation_details.append(OperationCostDetail(
                    operation_no=rl.operation_no,
                    operation_desc=op.description if op else rl.operation_code,
                    work_center=rl.work_center_code,
                    run_time_min=rl.run_time_minutes,
                    setup_time_min=rl.setup_time_minutes,
                    batch_size=rl.batch_size,
                    cost_per_hour=cost_per_hour,
                    run_cost=round(run_cost, 4),
                    setup_cost_per_unit=round(setup_cost_per_unit, 4),
                    total_cost=round(run_cost + setup_cost_per_unit, 4),
                ))
            
            co_result.operation_cost = round(total_run, 4)
            co_result.setup_cost = round(total_setup, 4)
            
            # 3. Brutto produksjonskost
            co_result.gross_production_cost = (
                co_result.material_cost + co_result.operation_cost + co_result.setup_cost
            )
            
            # 4. Biproduktverdi for co-produkt (hvis det har egne byproduct rules)
            self._calc_byproduct_value(co_result)
            
            # 5. Netto produksjonskost
            co_result.net_production_cost = co_result.gross_production_cost - co_result.by_product_value
            
            # 6. Prosentandeler
            self._calc_percentages(co_result)
            
            co_results.append(co_result)
        
        return co_results

    def calculate_all(self) -> list[ProductCostResult]:
        """Beregn kost for alle ferdigvarer og halvfabrikata.
        
        Beregner per fabrikk: for hvert produkt som har routing pa flere
        fabrikker, returneres ett resultat per fabrikk.
        
        Sorterer produkter i avhengighetsrekkefolge (barn for foreldre)
        slik at dynamisk cost roll-up fungerer korrekt.
        
        Inkluderer ogsa co-produkter (f.eks. B-vare) som egne resultater.
        """
        # Finn alle ferdigvarer/halvfabrikata
        target_items = [
            prod for prod in self.data.products
            if prod.item_type in ("Finished Good", "Semi Finished")
        ]
        
        # Bygg avhengighetsgraf: hvilke FG/HF-produkter er komponenter i BOM?
        # Et produkt ma beregnes for det brukes som komponent i et annet
        def get_dependency_order(items: list[Product]) -> list[Product]:
            """Topologisk sortering: barn for foreldre."""
            item_set = {p.item_no for p in items}
            # Finn BOM-relasjoner mellom FG/HF-produkter
            edges: dict[str, list[str]] = {}  # parent -> [children]
            for bl in self.data.bom_lines:
                if bl.parent_item_no in item_set and bl.component_item_no in item_set:
                    if bl.parent_item_no not in edges:
                        edges[bl.parent_item_no] = []
                    edges[bl.parent_item_no].append(bl.component_item_no)
            
            # Topologisk sortering (DFS)
            visited: set[str] = set()
            order: list[str] = []
            
            def dfs(item_no: str):
                if item_no in visited:
                    return
                visited.add(item_no)
                # Besok barn forst
                for child in edges.get(item_no, []):
                    dfs(child)
                order.append(item_no)
            
            for p in items:
                dfs(p.item_no)
            
            # Map tilbake til Product-objekter
            item_map = {p.item_no: p for p in items}
            return [item_map[no] for no in order if no in item_map]
        
        sorted_items = get_dependency_order(target_items)
        
        results = []
        for prod in sorted_items:
            # Beregn for alle fabrikker som har routing for dette produktet
            product_results = self.calculate_product_costs(prod.item_no)
            results.extend(product_results)
            
            # Generer co-produkt resultater (f.eks. B-vare)
            for pr in product_results:
                co_results = self._calc_co_product_results(pr)
                results.extend(co_results)
            
            # Store calculated gross costs for dynamic cost roll-up
            if product_results:
                self.calculated_costs[prod.item_no] = {
                    r.location_code: r.gross_production_cost for r in product_results
                }
        return results



# ──────────────────────────────────────────────────────────────────────
#  5. SIMULERINGSVERKTØY (for Marimo-appen)
# ──────────────────────────────────────────────────────────────────────

@dataclass
class SimulationOverride:
    """Definerer hvilke parametere som skal overstyres i en simulering.
    
    Alle felter er optional - kun de som settes vil bli overstyrt.
    """
    # Overstyring av Item Costs (råvarepriser)
    # key: item_no, value: ny unit_cost
    item_costs: dict[str, float] = field(default_factory=dict)
    
    # Overstyring av Work Center rates
    # key: work_center_code, value: dict med felter som skal overstyres
    # Støttede felter: labor_cost_hour, machine_cost_hour, overhead_cost_hour
    work_centers: dict[str, dict[str, float]] = field(default_factory=dict)
    
    # Overstyring av BOM (scrap %)
    # key: (parent_item_no, component_item_no), value: ny scrap_pct
    bom_scrap: dict[tuple[str, str], float] = field(default_factory=dict)
    
    # Overstyring av BOM (co-prod %)
    # key: (parent_item_no, component_item_no), value: ny co_product_pct
    bom_co_product: dict[tuple[str, str], float] = field(default_factory=dict)
    
    # Overstyring av Routing (setup/run times, batch size)
    # key: (item_no, operation_no, work_center_code), value: dict med felter som skal overstyres
    # Støttede felter: setup_time_minutes, run_time_minutes, batch_size
    routing: dict[tuple[str, int, str], dict[str, float]] = field(default_factory=dict)
    
    # Overstyring av By Product Rules (market value)
    # key: (parent_item_no, by_product_item_no), value: ny market_value
    byproduct_values: dict[tuple[str, str], float] = field(default_factory=dict)
    
    # Overstyring av Transport Ruter (cost_per_m3)
    # key: (from_loc, to_loc), value: dict med felter som skal overstyres
    # Støttede felter: cost_per_m3
    transport_ruter: dict[tuple[str, str], dict[str, float]] = field(default_factory=dict)
    
    # Scenario-parametere
    planned_quantity: Optional[float] = None


@dataclass
class SimulationComparison:
    """Sammenligning mellom original og simulert kostnad for ett produkt."""
    product_no: str
    product_desc: str
    product_group: str
    base_uom: str
    location_code: str = ""
    location_name: str = ""
    
    # Originale verdier
    original_material_cost: float = 0.0
    original_operation_cost: float = 0.0
    original_setup_cost: float = 0.0
    original_gross_cost: float = 0.0
    original_byproduct_value: float = 0.0
    original_net_cost: float = 0.0
    
    # Simulerte verdier
    simulated_material_cost: float = 0.0
    simulated_operation_cost: float = 0.0
    simulated_setup_cost: float = 0.0
    simulated_gross_cost: float = 0.0
    simulated_byproduct_value: float = 0.0
    simulated_net_cost: float = 0.0
    
    # Detaljer
    original_material_details: list[MaterialCostDetail] = field(default_factory=list)
    simulated_material_details: list[MaterialCostDetail] = field(default_factory=list)
    original_operation_details: list[OperationCostDetail] = field(default_factory=list)
    simulated_operation_details: list[OperationCostDetail] = field(default_factory=list)
    original_byproduct_details: list[ByProductDetail] = field(default_factory=list)
    simulated_byproduct_details: list[ByProductDetail] = field(default_factory=list)
    
    # Co-produkt resultater (f.eks. B-vare)
    original_co_product_results: list[ProductCostResult] = field(default_factory=list)
    simulated_co_product_results: list[ProductCostResult] = field(default_factory=list)
    
    # Co-produkt info
    co_product_pct: float = 0.0
    co_product_item_no: str = ""
    
    # Scenario (hvis relevant)
    planned_quantity: Optional[float] = None
    simulated_total_net_cost: float = 0.0
    simulated_cost_per_unit: float = 0.0
    simulated_total_hours: float = 0.0
    simulated_work_center_hours: dict = field(default_factory=dict)
    
    @property
    def material_diff(self) -> float:
        return self.simulated_material_cost - self.original_material_cost
    
    @property
    def operation_diff(self) -> float:
        return self.simulated_operation_cost - self.original_operation_cost
    
    @property
    def setup_diff(self) -> float:
        return self.simulated_setup_cost - self.original_setup_cost
    
    @property
    def gross_diff(self) -> float:
        return self.simulated_gross_cost - self.original_gross_cost
    
    @property
    def byproduct_diff(self) -> float:
        return self.simulated_byproduct_value - self.original_byproduct_value
    
    @property
    def net_diff(self) -> float:
        return self.simulated_net_cost - self.original_net_cost


class SimulationEngine:
    """Motor for a kjore "what-if" simuleringer med overstyrte parametere.
    
    Tar et ExcelData-objekt og en SimulationOverride, og beregner
    baade original (baseline) og simulert kostnad for sammenligning.
    """
    
    def __init__(self, data: ExcelData):
        self.data = data
    
    def _apply_overrides(self, overrides: SimulationOverride) -> ExcelData:
        """Lag en kopi av ExcelData med overstyrte verdier.
        
        Dette er en "deep copy" strategi: vi oppretter et nytt ExcelData-objekt
        og kopierer alle lister, men med overstyrte verdier.
        """
        # Opprett et tomt ExcelData-objekt
        sim_data = ExcelData.__new__(ExcelData)
        sim_data.filepath = self.data.filepath
        sim_data.raw = self.data.raw
        
        # Kopier produkter (uendret)
        sim_data.products = list(self.data.products)
        sim_data.locations = list(self.data.locations)
        sim_data.operations = list(self.data.operations)
        sim_data.capacity_days = list(self.data.capacity_days)
        sim_data.scenarios = list(self.data.scenarios)
        
        # Kopier og overstyr Item Costs
        sim_data.item_costs = []
        for ic in self.data.item_costs:
            if ic.item_no in overrides.item_costs:
                sim_data.item_costs.append(ItemCost(
                    item_no=ic.item_no,
                    cost_type=ic.cost_type,
                    unit_cost=overrides.item_costs[ic.item_no],
                    currency=ic.currency,
                    effective_date=ic.effective_date,
                ))
            else:
                sim_data.item_costs.append(ic)
        
        # Kopier og overstyr Work Centers
        sim_data.work_centers = []
        for wc in self.data.work_centers:
            if wc.code in overrides.work_centers:
                wc_overrides = overrides.work_centers[wc.code]
                sim_data.work_centers.append(WorkCenter(
                    code=wc.code,
                    description=wc.description,
                    location_code=wc.location_code,
                    labor_cost_hour=wc_overrides.get("labor_cost_hour", wc.labor_cost_hour),
                    machine_cost_hour=wc_overrides.get("machine_cost_hour", wc.machine_cost_hour),
                    overhead_cost_hour=wc_overrides.get("overhead_cost_hour", wc.overhead_cost_hour),
                    capacity_hours_day=wc_overrides.get("capacity_hours_day", wc.capacity_hours_day),
                    effective_capacity_pct=wc_overrides.get("effective_capacity_pct", wc.effective_capacity_pct),
                    active=wc.active,
                ))
            else:
                sim_data.work_centers.append(wc)
        
        # Kopier og overstyr BOM (scrap % og co-prod %)
        sim_data.bom_lines = []
        for bl in self.data.bom_lines:
            key = (bl.parent_item_no, bl.component_item_no)
            if key in overrides.bom_scrap or key in overrides.bom_co_product:
                sim_data.bom_lines.append(BOMLine(
                    parent_item_no=bl.parent_item_no,
                    component_item_no=bl.component_item_no,
                    quantity_per=bl.quantity_per,
                    uom=bl.uom,
                    scrap_pct=overrides.bom_scrap.get(key, bl.scrap_pct),
                    co_product_pct=overrides.bom_co_product.get(key, bl.co_product_pct),
                    co_product_item_no=bl.co_product_item_no,
                    valid_from=bl.valid_from,
                    valid_to=bl.valid_to,
                ))
            else:
                sim_data.bom_lines.append(bl)
        
        # Kopier og overstyr Routing
        # Nøkkel: (item_no, operation_no, work_center_code)
        sim_data.routing_lines = []
        for rl in self.data.routing_lines:
            key = (rl.item_no, rl.operation_no, rl.work_center_code)
            if key in overrides.routing:
                rl_overrides = overrides.routing[key]
                sim_data.routing_lines.append(RoutingLine(
                    item_no=rl.item_no,
                    operation_no=rl.operation_no,
                    operation_code=rl.operation_code,
                    work_center_code=rl.work_center_code,
                    setup_time_minutes=rl_overrides.get("setup_time_minutes", rl.setup_time_minutes),
                    run_time_minutes=rl_overrides.get("run_time_minutes", rl.run_time_minutes),
                    batch_size=rl_overrides.get("batch_size", rl.batch_size),
                    valid_from=rl.valid_from,
                    valid_to=rl.valid_to,
                ))
            else:
                sim_data.routing_lines.append(rl)
        
        # Kopier og overstyr By Product Rules
        sim_data.byproduct_rules = []
        for bpr in self.data.byproduct_rules:
            key = (bpr.parent_item_no, bpr.by_product_item_no)
            if key in overrides.byproduct_values:
                sim_data.byproduct_rules.append(ByProductRule(
                    parent_item_no=bpr.parent_item_no,
                    by_product_item_no=bpr.by_product_item_no,
                    expected_quantity=bpr.expected_quantity,
                    uom=bpr.uom,
                    market_value=overrides.byproduct_values[key],
                    allocation_method=bpr.allocation_method,
                ))
            else:
                sim_data.byproduct_rules.append(bpr)
        
        # Bygg indekser for det simulerte data-objektet
        sim_data._product_index = {p.item_no: p for p in sim_data.products}
        sim_data._location_index = {l.code: l for l in sim_data.locations}
        sim_data._wc_index = {w.code: w for w in sim_data.work_centers}
        sim_data._op_index = {o.code: o for o in sim_data.operations}
        
        sim_data._bom_index = {}
        for bl in sim_data.bom_lines:
            sim_data._bom_index.setdefault(bl.parent_item_no, []).append(bl)
        
        sim_data._routing_index = {}
        for rl in sim_data.routing_lines:
            sim_data._routing_index.setdefault(rl.item_no, []).append(rl)
        for item_no in sim_data._routing_index:
            sim_data._routing_index[item_no].sort(key=lambda r: r.operation_no)
        
        sim_data._byproduct_index = {}
        for br in sim_data.byproduct_rules:
            sim_data._byproduct_index.setdefault(br.parent_item_no, []).append(br)
        
        sim_data._scenario_index = {s.scenario_name: s for s in sim_data.scenarios}
        
        sim_data._item_cost_index = {}
        for ic in sim_data.item_costs:
            sim_data._item_cost_index.setdefault(ic.item_no, []).append(ic)
        
        _temp_locs: dict[str, set[str]] = {}
        for rl in sim_data.routing_lines:
            wc = sim_data._wc_index.get(rl.work_center_code)
            if wc and wc.location_code:
                _temp_locs.setdefault(rl.item_no, set()).add(wc.location_code)
        sim_data._routing_locations_index = {
            item_no: sorted(locs) for item_no, locs in _temp_locs.items()
        }
        
        return sim_data
    
    def _calc_scenario_totals(self, result: ProductCostResult, quantity: float) -> dict:
        """Beregn scenario-totaler (total kost, timer, etc.) for et produkt."""
        totals = {
            "total_material_cost": result.material_cost * quantity,
            "total_operation_cost": result.operation_cost * quantity,
            "total_setup_cost": result.setup_cost * quantity,
            "total_byproduct_value": result.by_product_value * quantity,
            "total_net_cost": result.net_production_cost * quantity,
            "cost_per_unit": result.net_production_cost,
            "total_hours": 0.0,
            "work_center_hours": {},
        }
        
        # Beregn timebehov per arbeidssenter
        for od in result.operation_details:
            wc = od.work_center
            # Run-tid per enhet i timer
            run_hours = od.run_time_min / 60.0
            # Setup-tid per enhet i timer
            setup_hours = od.setup_time_min / 60.0 / od.batch_size if od.batch_size > 0 else 0
            
            total_hours_wc = (run_hours + setup_hours) * quantity
            if wc not in totals["work_center_hours"]:
                totals["work_center_hours"][wc] = 0.0
            totals["work_center_hours"][wc] += total_hours_wc
            totals["total_hours"] += total_hours_wc
        
        return totals
    
    def compare(self, product_no: str, overrides: SimulationOverride, 
                location_code: str = "") -> Optional[SimulationComparison]:
        """Sammenlign original vs simulert kostnad for ett produkt.
        
        Args:
            product_no: Produktnummer
            overrides: Overstyringsparametere
            location_code: Spesifikk fabrikk/lokasjon. Hvis tom, brukes forste.
        """
        prod = self.data.product(product_no)
        if prod is None:
            return None
        
        # Finn lokasjon
        if not location_code:
            location_codes = self.data.routing_locations_for(product_no)
            if not location_codes:
                return None
            location_code = location_codes[0]
        loc = self.data.location(location_code)
        
        # 1. Beregn original (baseline)
        calculator_orig = CostCalculator(self.data)
        # Kjor calculate_all forst for dynamisk cost roll-up
        calculator_orig.calculate_all()
        orig_results = calculator_orig.calculate_product_costs(product_no)
        if not orig_results:
            return None
        # Finn resultat for spesifikk lokasjon
        orig = next((r for r in orig_results if r.location_code == location_code), orig_results[0])
        
        # 2. Beregn simulert (med overrides)
        sim_data = self._apply_overrides(overrides)
        calculator_sim = CostCalculator(sim_data)
        calculator_sim.calculate_all()
        sim_results = calculator_sim.calculate_product_costs(product_no)
        if not sim_results:
            return None
        # Finn resultat for spesifikk lokasjon
        sim = next((r for r in sim_results if r.location_code == location_code), sim_results[0])
        
        # 3. Hent co-produkt resultater (f.eks. B-vare)
        orig_co_results = calculator_orig._calc_co_product_results(orig)
        sim_co_results = calculator_sim._calc_co_product_results(sim)
        
        # Finn co-prod info fra BOM
        co_pct = 0.0
        co_item_no = ""
        for bl in self.data.bom_lines:
            if bl.parent_item_no == product_no and bl.co_product_pct > 0:
                co_pct = bl.co_product_pct
                co_item_no = bl.co_product_item_no
                break
        
        # 4. Bygg sammenligning
        comparison = SimulationComparison(
            product_no=prod.item_no,
            product_desc=prod.description,
            product_group=prod.product_group,
            base_uom=prod.base_uom,
            location_code=location_code,
            location_name=loc.name if loc else "",
            
            original_material_cost=orig.material_cost,
            original_operation_cost=orig.operation_cost,
            original_setup_cost=orig.setup_cost,
            original_gross_cost=orig.gross_production_cost,
            original_byproduct_value=orig.by_product_value,
            original_net_cost=orig.net_production_cost,
            
            simulated_material_cost=sim.material_cost,
            simulated_operation_cost=sim.operation_cost,
            simulated_setup_cost=sim.setup_cost,
            simulated_gross_cost=sim.gross_production_cost,
            simulated_byproduct_value=sim.by_product_value,
            simulated_net_cost=sim.net_production_cost,
            
            original_material_details=list(orig.material_details),
            simulated_material_details=list(sim.material_details),
            original_operation_details=list(orig.operation_details),
            simulated_operation_details=list(sim.operation_details),
            original_byproduct_details=list(orig.byproduct_details),
            simulated_byproduct_details=list(sim.byproduct_details),
            
            original_co_product_results=orig_co_results,
            simulated_co_product_results=sim_co_results,
            co_product_pct=co_pct,
            co_product_item_no=co_item_no,
        )
        
        # 5. Scenario-beregning hvis kvantum er satt
        if overrides.planned_quantity is not None and overrides.planned_quantity > 0:
            qty = overrides.planned_quantity
            comparison.planned_quantity = qty
            sim_totals = self._calc_scenario_totals(sim, qty)
            comparison.simulated_total_net_cost = sim_totals["total_net_cost"]
            comparison.simulated_cost_per_unit = sim_totals["cost_per_unit"]
            comparison.simulated_total_hours = sim_totals["total_hours"]
            comparison.simulated_work_center_hours = sim_totals["work_center_hours"]
        
        return comparison
    
    @staticmethod
    def _index_results(results: list[ProductCostResult]) -> dict[str, dict[str, ProductCostResult]]:
        """Indekser en liste med ProductCostResults for raskt oppslag.
        
        Returnerer: {item_no: {location_code: ProductCostResult}}
        """
        idx: dict[str, dict[str, ProductCostResult]] = {}
        for r in results:
            if r.product_no not in idx:
                idx[r.product_no] = {}
            idx[r.product_no][r.location_code] = r
        return idx

    @staticmethod
    def _get_co_product_pct(bom_lines: list[BOMLine], product_no: str) -> tuple[float, str]:
        """Hent co-prod prosent og item_no for et produkt fra BOM."""
        for bl in bom_lines:
            if bl.parent_item_no == product_no and bl.co_product_pct > 0:
                return bl.co_product_pct, bl.co_product_item_no
        return 0.0, ""

    def _build_comparison(self, orig: ProductCostResult, sim: ProductCostResult,
                          overrides: SimulationOverride,
                          orig_calculator: CostCalculator, sim_calculator: CostCalculator,
                          co_pct: float, co_item_no: str) -> SimulationComparison:
        """Bygg et SimulationComparison-objekt fra original og simulert resultat."""
        # Co-produkt resultater
        orig_co_results = orig_calculator._calc_co_product_results(orig)
        sim_co_results = sim_calculator._calc_co_product_results(sim)
        
        comparison = SimulationComparison(
            product_no=orig.product_no,
            product_desc=orig.product_desc,
            product_group=orig.product_group,
            base_uom=orig.base_uom,
            location_code=orig.location_code,
            location_name=orig.location_name,
            
            original_material_cost=orig.material_cost,
            original_operation_cost=orig.operation_cost,
            original_setup_cost=orig.setup_cost,
            original_gross_cost=orig.gross_production_cost,
            original_byproduct_value=orig.by_product_value,
            original_net_cost=orig.net_production_cost,
            
            simulated_material_cost=sim.material_cost,
            simulated_operation_cost=sim.operation_cost,
            simulated_setup_cost=sim.setup_cost,
            simulated_gross_cost=sim.gross_production_cost,
            simulated_byproduct_value=sim.by_product_value,
            simulated_net_cost=sim.net_production_cost,
            
            original_material_details=list(orig.material_details),
            simulated_material_details=list(sim.material_details),
            original_operation_details=list(orig.operation_details),
            simulated_operation_details=list(sim.operation_details),
            original_byproduct_details=list(orig.byproduct_details),
            simulated_byproduct_details=list(sim.byproduct_details),
            
            original_co_product_results=orig_co_results,
            simulated_co_product_results=sim_co_results,
            co_product_pct=co_pct,
            co_product_item_no=co_item_no,
        )
        
        # Scenario-beregning hvis kvantum er satt
        if overrides.planned_quantity is not None and overrides.planned_quantity > 0:
            qty = overrides.planned_quantity
            comparison.planned_quantity = qty
            sim_totals = self._calc_scenario_totals(sim, qty)
            comparison.simulated_total_net_cost = sim_totals["total_net_cost"]
            comparison.simulated_cost_per_unit = sim_totals["cost_per_unit"]
            comparison.simulated_total_hours = sim_totals["total_hours"]
            comparison.simulated_work_center_hours = sim_totals["work_center_hours"]
        
        return comparison

    def compare_all(self, overrides: SimulationOverride, 
                    product_filter: Optional[list[str]] = None) -> list[SimulationComparison]:
        """Sammenlign original vs simulert for alle (eller filtrerte) produkter.
        
        *** OPTIMERT VERSJON ***
        Kjor calculate_all() kun 2 ganger totalt (istedenfor 2x per produkt).
        Bruker indeksert oppslag for a bygge SimulationComparison-objekter.
        
        Returnerer resultater per fabrikk/lokasjon: hvis et produkt har routing
        pa flere fabrikker, returneres ett resultat per fabrikk.
        """
        # Finn alle ferdigvarer/halvfabrikata
        target_items = [
            prod for prod in self.data.products
            if prod.item_type in ("Finished Good", "Semi Finished")
        ]
        
        if product_filter:
            target_items = [p for p in target_items if p.item_no in product_filter]
        
        # --- Trinn 1: Beregn original (baseline) ÉN gang ---
        calculator_orig = CostCalculator(self.data)
        orig_all_results = calculator_orig.calculate_all()
        orig_index = self._index_results(orig_all_results)
        
        # --- Trinn 2: Beregn simulert (med overrides) ÉN gang ---
        sim_data = self._apply_overrides(overrides)
        calculator_sim = CostCalculator(sim_data)
        sim_all_results = calculator_sim.calculate_all()
        sim_index = self._index_results(sim_all_results)
        
        # --- Trinn 3: Bygg comparisons ved a sla opp i indeksene ---
        results = []
        for prod in target_items:
            # Finn alle lokasjoner for dette produktet
            location_codes = self.data.routing_locations_for(prod.item_no)
            if not location_codes:
                continue
            
            # Hent co-prod info (samme for alle lokasjoner)
            co_pct, co_item_no = self._get_co_product_pct(
                self.data.bom_lines, prod.item_no
            )
            
            for loc_code in location_codes:
                orig_result = orig_index.get(prod.item_no, {}).get(loc_code)
                sim_result = sim_index.get(prod.item_no, {}).get(loc_code)
                
                if orig_result is None or sim_result is None:
                    continue
                
                comparison = self._build_comparison(
                    orig_result, sim_result, overrides,
                    calculator_orig, calculator_sim,
                    co_pct, co_item_no
                )
                results.append(comparison)
        
        return results





# ──────────────────────────────────────────────────────────────────────
#  5B. TRANSPORTVARE-UTVIDELSE (visuelt simuleringslag)
#
#  Transport beregnes som et RENT VISUELT LAG oppå den eksisterende
#  datamodellen. Datamodellen (products, bom_lines, routing_lines) muteres
#  ALDRI. For flaggede varer (transport_flagg.is_transport=1) lages det
#  fiktive rader med location_code f.eks. "KOD→KV" som inkluderer frakt.
#
#  Den gamle muterende logikken (sync_transport_varer i excel_bridge.py)
#  er beholdt urørt som LEGACY for fremtidig Business Central-integrasjon.
# ──────────────────────────────────────────────────────────────────────

def _prod_locations_from_data(data, item_no: str) -> set[str]:
    """Finn produksjonslokasjoner for et produkt via routing → work_centers."""
    locs: set[str] = set()
    for rl in data.routing_for(item_no):
        wc = data.work_center(rl.work_center_code)
        if wc and wc.location_code:
            locs.add(wc.location_code)
    return locs


def _finn_lm_per_m3(data, item_no: str):
    """Finn antall LM per M3 for en vare ved å følge BOM-kjeden til råvare.

    Rekursiv traversering: hvis en BOM-komponent er Raw Material, returneres
    dens quantity_per (LM/M3). Ellers går vi rekursivt ned i komponenten
    (f.eks. Semi Finished → Raw Material).

    Returns:
        float (LM/M3) eller None hvis ingen råvarekomponent finnes.
    """
    for bl in data.bom_for(item_no):
        comp = data.product(bl.component_item_no)
        if comp and comp.item_type == "Raw Material":
            return bl.quantity_per if bl.quantity_per and bl.quantity_per > 0 else None
        # Rekursivt nedover i kjeden (følg semi-finished/ferdigvare-komponent)
        result = _finn_lm_per_m3(data, bl.component_item_no)
        if result:
            return result
    return None


def _load_transport_konfig(data, db=None) -> tuple[set[str], dict[tuple[str, str], tuple[float, float, float]]]:
    """Les transport_flagg + transport_ruter fra databasen."""
    if db is None:
        db = getattr(data, "db", None)
    if db is None:
        return set(), {}
    try:
        flagged = {r["item_no"] for r in db.conn.execute(
            "SELECT item_no FROM transport_flagg WHERE is_transport=1"
        ).fetchall()}
        ruter: dict[tuple[str, str], tuple[float, float, float]] = {}
        for r in db.conn.execute(
            "SELECT from_loc, to_loc, cost_per_m3, distance_km, hours FROM transport_ruter"
        ).fetchall():
            ruter[(r["from_loc"], r["to_loc"])] = (
                r["cost_per_m3"], r["distance_km"], r["hours"],
            )
        return flagged, ruter
    except Exception:
        return set(), {}


def _beregn_transportkost_per_lm(rute: tuple[float, float, float],
                                 lm_per_m3: float) -> float:
    """Beregn transportkost per LM fra en rute.

    Args:
        rute: (cost_per_m3, distance_km, hours)
        lm_per_m3: antall LM per M3 for varen

    Returns:
        transportkost per LM. Returnerer 0.0 hvis lm_per_m3 mangler/<= 0.
    """
    cost_per_m3, _, _ = rute
    if lm_per_m3 and lm_per_m3 > 0:
        return cost_per_m3 / lm_per_m3
    return 0.0


def _transport_operation_detail(from_loc: str, to_loc: str, rute,
                                transport_kost: float, lm_per_m3: float) -> OperationCostDetail:
    """Bygg en OPERASJONSDETALJ som representerer frakt mellom to lokasjoner."""
    cost_per_m3, distance_km, hours = rute
    opp_desc = "Frakt {f} -> {t}".format(f=from_loc, t=to_loc)
    if lm_per_m3 and lm_per_m3 > 0:
        opp_desc = "{d} ({c:.2f} kr/M3 / {l} LM/M3)".format(
            d=opp_desc, c=cost_per_m3, l=round(lm_per_m3, 1))
    return OperationCostDetail(
        operation_no=999,
        operation_desc=opp_desc,
        work_center="TRANSPORT",
        run_time_min=float(hours),
        setup_time_min=0.0,
        batch_size=1.0,
        cost_per_hour=round(cost_per_m3, 2),
        run_cost=round(transport_kost, 4),
        setup_cost_per_unit=0.0,
        total_cost=round(transport_kost, 4),
    )


def expand_product_costs_with_transport(results: list[ProductCostResult], data, db=None
                                        ) -> list[ProductCostResult]:
    """Utvid en baseline-liste med fiktive transportrader for flaggede varer."""
    flagged, ruter = _load_transport_konfig(data, db)
    if not flagged or not ruter:
        return list(results)

    expanded = list(results)
    factory_locs = {l.code for l in data.locations if l.location_type == "Factory"}

    by_product: dict[str, list[ProductCostResult]] = {}
    for r in results:
        by_product.setdefault(r.product_no, []).append(r)

    for prod_no, prod_results in by_product.items():
        if prod_no not in flagged:
            continue
        prod_locs = _prod_locations_from_data(data, prod_no)
        if not prod_locs:
            continue
        from_loc = sorted(prod_locs)[0]
        base = next((r for r in prod_results if r.location_code == from_loc), None)
        if base is None:
            continue
        lm_per_m3 = _finn_lm_per_m3(data, prod_no)
        if not lm_per_m3 or lm_per_m3 <= 0:
            continue  # Uten LM/M3-konvertering kan vi ikke beregne frakt per LM

        for to_loc in sorted(factory_locs - prod_locs):
            rute = ruter.get((from_loc, to_loc))
            if rute is None:
                continue
            transport_kost = _beregn_transportkost_per_lm(rute, lm_per_m3)
            if transport_kost <= 0:
                continue

            ny = ProductCostResult(
                product_no=base.product_no,
                product_desc=base.product_desc,
                product_group=base.product_group,
                base_uom=base.base_uom,
                location_code="{f}->{t}".format(f=from_loc, t=to_loc),
                location_name="{f} -> {t}".format(f=from_loc, t=to_loc),
                material_cost=base.material_cost,
                operation_cost=round(base.operation_cost + transport_kost, 4),
                setup_cost=base.setup_cost,
                gross_production_cost=round(base.gross_production_cost + transport_kost, 4),
                by_product_value=base.by_product_value,
                net_production_cost=round(base.net_production_cost + transport_kost, 4),
                material_details=list(base.material_details),
                operation_details=list(base.operation_details) + [
                    _transport_operation_detail(from_loc, to_loc, rute, transport_kost, lm_per_m3)
                ],
                byproduct_details=list(base.byproduct_details),
            )

            gross = ny.gross_production_cost
            if gross:
                ny.cost_breakdown = [
                    {"category": "Materialkost", "amount": round(ny.material_cost, 4),
                     "pct_of_gross": round(ny.material_cost / gross * 100, 1)},
                    {"category": "Operasjonskost", "amount": round(ny.operation_cost, 4),
                     "pct_of_gross": round(ny.operation_cost / gross * 100, 1)},
                    {"category": "Setupkost", "amount": round(ny.setup_cost, 4),
                     "pct_of_gross": round(ny.setup_cost / gross * 100, 1)},
                ]
                if ny.by_product_value > 0:
                    ny.cost_breakdown.append(
                        {"category": "Biproduktverdi", "amount": -round(ny.by_product_value, 4),
                         "pct_of_gross": -round(ny.by_product_value / gross * 100, 1)}
                    )
                ny.cost_breakdown.append(
                    {"category": "Netto kost", "amount": round(ny.net_production_cost, 4),
                     "pct_of_gross": round(ny.net_production_cost / gross * 100, 1)}
                )

            expanded.append(ny)

    return expanded


def expand_simulations_with_transport(comparisons: list[SimulationComparison], data, db=None,
                                      transport_ruter_overrides: Optional[dict] = None
                                      ) -> list[SimulationComparison]:
    """Utvid en simuleringssammenligningsliste med fiktive transportrader.

    Args:
        comparisons: liste med SimulationComparison fra SimulationEngine
        data: datakilde (SqliteData/ExcelData)
        db: valgfri database (DataRepo). Standard: data.db
        transport_ruter_overrides: dict {(from_loc, to_loc): dict[str, float]}
            Overstyrer cost_per_m3 for transportrutene i simuleringen.
    """
    flagged, ruter = _load_transport_konfig(data, db)
    if not flagged or not ruter:
        return list(comparisons)

    expanded = list(comparisons)
    factory_locs = {l.code for l in data.locations if l.location_type == "Factory"}

    by_product: dict[str, list[SimulationComparison]] = {}
    for c in comparisons:
        by_product.setdefault(c.product_no, []).append(c)

    for prod_no, comps in by_product.items():
        if prod_no not in flagged:
            continue
        prod_locs = _prod_locations_from_data(data, prod_no)
        if not prod_locs:
            continue
        from_loc = sorted(prod_locs)[0]
        base = next((c for c in comps if c.location_code == from_loc), None)
        if base is None:
            continue
        lm_per_m3 = _finn_lm_per_m3(data, prod_no)
        if not lm_per_m3 or lm_per_m3 <= 0:
            continue

        for to_loc in sorted(factory_locs - prod_locs):
            rute_orig = ruter.get((from_loc, to_loc))
            if rute_orig is None:
                continue
            # Bruk evt. overstyrt cost_per_m3 fra simuleringen
            cost_per_m3 = rute_orig[0]
            if transport_ruter_overrides:
                ovr = transport_ruter_overrides.get((from_loc, to_loc))
                if ovr and "cost_per_m3" in ovr:
                    cost_per_m3 = ovr["cost_per_m3"]
            rute = (cost_per_m3, rute_orig[1], rute_orig[2])
            transport_kost = _beregn_transportkost_per_lm(rute, lm_per_m3)
            if transport_kost <= 0:
                continue
            transport_detail = _transport_operation_detail(from_loc, to_loc, rute, transport_kost, lm_per_m3)

            ny = SimulationComparison(
                product_no=base.product_no,
                product_desc=base.product_desc,
                product_group=base.product_group,
                base_uom=base.base_uom,
                location_code="{f}->{t}".format(f=from_loc, t=to_loc),
                location_name="{f} -> {t}".format(f=from_loc, t=to_loc),
                original_material_cost=base.original_material_cost,
                original_operation_cost=round(base.original_operation_cost + transport_kost, 4),
                original_setup_cost=base.original_setup_cost,
                original_gross_cost=round(base.original_gross_cost + transport_kost, 4),
                original_byproduct_value=base.original_byproduct_value,
                original_net_cost=round(base.original_net_cost + transport_kost, 4),
                simulated_material_cost=base.simulated_material_cost,
                simulated_operation_cost=round(base.simulated_operation_cost + transport_kost, 4),
                simulated_setup_cost=base.simulated_setup_cost,
                simulated_gross_cost=round(base.simulated_gross_cost + transport_kost, 4),
                simulated_byproduct_value=base.simulated_byproduct_value,
                simulated_net_cost=round(base.simulated_net_cost + transport_kost, 4),
                original_material_details=list(base.original_material_details),
                simulated_material_details=list(base.simulated_material_details),
                original_operation_details=list(base.original_operation_details) + [transport_detail],
                simulated_operation_details=list(base.simulated_operation_details) + [transport_detail],
                original_byproduct_details=list(base.original_byproduct_details),
                simulated_byproduct_details=list(base.simulated_byproduct_details),
                original_co_product_results=list(base.original_co_product_results),
                simulated_co_product_results=list(base.simulated_co_product_results),
                co_product_pct=base.co_product_pct,
                co_product_item_no=base.co_product_item_no,
                planned_quantity=base.planned_quantity,
                simulated_total_net_cost=base.simulated_total_net_cost,
                simulated_cost_per_unit=base.simulated_cost_per_unit,
                simulated_total_hours=base.simulated_total_hours,
                simulated_work_center_hours=base.simulated_work_center_hours,
            )
            expanded.append(ny)

    return expanded


# ──────────────────────────────────────────────────────────────────────
#  6. JSON-EKSPORT
# ──────────────────────────────────────────────────────────────────────

def _serialize(obj):
    """Hjelpefunksjon for a serialisere dataclass-objekter til dict."""
    if hasattr(obj, '__dataclass_fields__'):
        return {k: _serialize(v) for k, v in obj.__dict__.items()}
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    if isinstance(obj, dict):
        return {k: _serialize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_serialize(v) for v in obj]
    return obj


def export_product_costs_to_json(results: list[ProductCostResult], filepath: str):
    """Eksporter produktkalkyle til JSON."""
    data = {
        "type": "product_cost_calculation",
        "exported_at": datetime.now().isoformat(),
        "results": [_serialize(r) for r in results],
    }
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"  >>> Produktkalkyle eksportert til: {filepath}")


def export_scenario_results_to_json(results: list[ScenarioResult], filepath: str):
    """Eksporter simuleringsresultater til JSON."""
    data = {
        "type": "scenario_simulation",
        "exported_at": datetime.now().isoformat(),
        "results": [_serialize(r) for r in results],
    }
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"  >>> Simuleringsresultater eksportert til: {filepath}")



# ──────────────────────────────────────────────────────────────────────
#  7. TESTDATA (Kodal Hovleri-eksemplet)
# ──────────────────────────────────────────────────────────────────────

def create_test_data() -> ExcelData:
    """Opprett et ExcelData-objekt med testdata for Kodal Hovleri."""
    data = ExcelData.__new__(ExcelData)
    data.filepath = Path("Produksjonsmodell_Mal.xlsx")
    data.raw = {}

    data.products = [
        Product("RM001", "Skrulast 48x198", "Raw Material", "Skrulast", "M3", True),
        Product("FG001", "Utvendig Panel 21x95", "Finished Good", "Panel", "LM", True),
        Product("BP001", "Hovelspon", "By Product", "Spon", "KG", True),
    ]

    data.locations = [
        Location("KOD", "Kodal Fabrikk", "Factory", True),
        Location("KV", "Kvås", "Factory", True),
    ]

    data.work_centers = [
        WorkCenter("HOVEDHOVEL", "Hovedhovel", "KOD",
                   labor_cost_hour=550, machine_cost_hour=900, overhead_cost_hour=150,
                   capacity_hours_day=16, effective_capacity_pct=85, active=True),
        WorkCenter("SPESIALHOVEL", "Spesialhovel", "KOD",
                   labor_cost_hour=550, machine_cost_hour=950, overhead_cost_hour=150,
                   capacity_hours_day=16, effective_capacity_pct=85, active=True),
        WorkCenter("KVHOVEL", "Kvås høvel", "KV",
                   labor_cost_hour=500, machine_cost_hour=200, overhead_cost_hour=100,
                   capacity_hours_day=16, effective_capacity_pct=85, active=True),
    ]

    data.operations = [
        Operation("RIP", "Oppdeling", "HOVEDHOVEL", "Minutes", True),
        Operation("PLANING", "Hovling", "HOVEDHOVEL", "Minutes", True),
        Operation("PROFILE", "Profilering", "SPESIALHOVEL", "Minutes", True),
        Operation("PACKING", "Pakking", "", "Minutes", True),
    ]

    data.item_costs = [
        ItemCost("RM001", "Standard Cost", 3000.00, "NOK", date(2026, 1, 1)),
        ItemCost("FG001", "Standard Cost", 0.0, "NOK", date(2026, 1, 1)),
        ItemCost("BP001", "Standard Cost", 1.50, "NOK", date(2026, 1, 1)),
    ]

    # Quantity Per = output-enheter per input-enhet
    # 1 M3 RM001 = 400 LM FG001
    data.bom_lines = [
        BOMLine("FG001", "RM001", 400, "LM", 5.0, 0.0, "", date(2026, 1, 1), None),
    ]

    data.routing_lines = [
        # KOD (Kodal) - full produksjon
        RoutingLine("FG001", 10, "RIP", "HOVEDHOVEL", 15.0, 0.15, 500, date(2026, 1, 1), None),
        RoutingLine("FG001", 20, "PLANING", "HOVEDHOVEL", 10.0, 0.10, 500, date(2026, 1, 1), None),
        RoutingLine("FG001", 30, "PROFILE", "SPESIALHOVEL", 20.0, 0.12, 500, date(2026, 1, 1), None),
        # KV (Kvås) - kun hovling, annen batch
        RoutingLine("FG001", 10, "RIP", "KVHOVEL", 15.0, 0.15, 1000, date(2026, 1, 1), None),
    ]


    # By Product Rules knyttet til produktet (Parent Item No)
    data.byproduct_rules = [
        ByProductRule("FG001", "BP001", 0.5, "KG", 1.50, "Reduce Main Product Cost"),
    ]

    data.capacity_days = []

    data.scenarios = [
        ProductionScenario("Normal Produksjon", "FG001", 100_000, date(2026, 1, 1), date(2026, 12, 31)),
        ProductionScenario("Full Kapasitet", "FG001", 200_000, date(2026, 1, 1), date(2026, 12, 31)),
    ]

    # Bygg indekser (gjort manuelt siden vi bruker __new__ i stedet for __init__)
    data._product_index = {p.item_no: p for p in data.products}
    data._location_index = {l.code: l for l in data.locations}
    data._wc_index = {w.code: w for w in data.work_centers}
    data._op_index = {o.code: o for o in data.operations}
    
    data._bom_index = {}
    for bl in data.bom_lines:
        data._bom_index.setdefault(bl.parent_item_no, []).append(bl)
    
    data._routing_index = {}
    for rl in data.routing_lines:
        data._routing_index.setdefault(rl.item_no, []).append(rl)
    for item_no in data._routing_index:
        data._routing_index[item_no].sort(key=lambda r: r.operation_no)
    
    data._byproduct_index = {}
    for br in data.byproduct_rules:
        data._byproduct_index.setdefault(br.parent_item_no, []).append(br)
    
    data._scenario_index = {s.scenario_name: s for s in data.scenarios}
    
    data._item_cost_index = {}
    for ic in data.item_costs:
        data._item_cost_index.setdefault(ic.item_no, []).append(ic)
    
    _temp_locs: dict[str, set[str]] = {}
    for rl in data.routing_lines:
        wc = data._wc_index.get(rl.work_center_code)
        if wc and wc.location_code:
            _temp_locs.setdefault(rl.item_no, set()).add(wc.location_code)
    data._routing_locations_index = {
        item_no: sorted(locs) for item_no, locs in _temp_locs.items()
    }
    
    return data


# ──────────────────────────────────────────────────────────────────────
#  8. RAPPORTERING (konsoll)
# ──────────────────────────────────────────────────────────────────────

def print_product_cost(result: ProductCostResult):
    """Skriv en detaljert kostkalkyle for ett produkt til konsoll."""
    sep = "-" * 55
    print(f"\n{sep}")
    if result.location_code:
        print(f"  {result.product_no}: {result.product_desc}")
        print(f"  Fabrikk: {result.location_code} - {result.location_name}")
    else:
        print(f"  {result.product_no}: {result.product_desc}")
    print(f"  Produktgruppe: {result.product_group}  |  Enhet: {result.base_uom}")
    print(sep)

    if result.cost_breakdown:
        print(f"\n  Kostnadsfordeling:")
        for cb in result.cost_breakdown:
            print(f"     {cb['category']:20s}  {cb['amount']:>10.4f} kr  ({cb['pct_of_gross']:>6.1f} %)")

    if result.material_details:
        print(f"\n  Materialkost:          {result.material_cost:>10.4f} kr/{result.base_uom}")
        for md in result.material_details:
            print(f"     {md.component}: {md.component_desc}")
            print(f"       Pris: {md.unit_cost:.2f} kr/{md.uom}")
            print(f"       Output per input: {md.quantity_per} {md.uom}")
            print(f"       Svinn: {md.scrap_pct}%  |  Kost: {md.total_cost:.4f} kr"
                  f"  ({md.pct_of_gross:.1f}% av brutto, {md.pct_of_material:.1f}% av material)")

    if result.operation_details:
        print(f"\n  Operasjonskost:        {result.operation_cost:>10.4f} kr/{result.base_uom}")
        print(f"  Setupkost:             {result.setup_cost:>10.4f} kr/{result.base_uom}")
        for od in result.operation_details:
            print(f"     Op {od.operation_no}: {od.operation_desc} @ {od.work_center}")
            print(f"       Kjoretid: {od.run_time_min} min  |  Setup: {od.setup_time_min} min / batch({od.batch_size})")
            print(f"       Timekost: {od.cost_per_hour:.2f} kr/t  |  = {od.total_cost:.4f} kr"
                  f"  ({od.pct_of_gross:.1f}% av brutto, {od.pct_of_operation:.1f}% av operasjon)")

    print(f"\n  --------------------------------")
    print(f"  Brutto produksjonskost:  {result.gross_production_cost:>10.4f} kr/{result.base_uom}")

    if result.byproduct_details:
        print(f"\n  Biproduktverdi:        {result.by_product_value:>10.4f} kr/{result.base_uom}")
        for bd in result.byproduct_details:
            print(f"     {bd.item_no}: {bd.description}")
            print(f"       {bd.quantity} {bd.uom} x {bd.market_value:.2f} kr/{bd.uom} = {bd.total_value:.4f} kr"
                  f"  ({bd.allocation_method})  ({bd.pct_of_gross:.1f}% av brutto)")

    print(f"\n  =================================")
    print(f"  NETTO PRODUKSJONSKOST:  {result.net_production_cost:>10.4f} kr/{result.base_uom}")
    print(f"  =================================\n")


def print_scenario_result(result: ScenarioResult):
    """Skriv simuleringsresultat til konsoll."""
    sep = "=" * 55
    print(f"\n{sep}")
    print(f"  Scenario: {result.scenario_name}")
    print(f"  Produkt:    {result.product_no} - {result.product_desc}")
    print(f"  Kvantum:    {result.planned_quantity:,.0f} {result.uom}")
    print(sep)
    print(f"  Kost per enhet:          {result.cost_per_unit:>10.4f} kr/{result.uom}")
    print(f"  Total materialkost:      {result.total_material_cost:>12.2f} kr")
    print(f"  Total operasjonskost:    {result.total_operation_cost:>12.2f} kr")
    print(f"  Total setupkost:         {result.total_setup_cost:>12.2f} kr")
    if result.total_changeover_cost > 0:
        print(f"  Total omstillingskost:   {result.total_changeover_cost:>12.2f} kr")
    print(f"  Total biproduktverdi:    {result.total_byproduct_value:>12.2f} kr")
    print(f"  --------------------------------")
    print(f"  TOTAL NETTO KOST:        {result.total_net_cost:>12.2f} kr")
    print(f"  Totalt timebehov:        {result.total_hours_needed:>10.2f} timer")
    if result.work_center_hours:
        print(f"  Per arbeidssenter:")
        for wc, hrs in result.work_center_hours.items():
            print(f"     {wc}: {hrs:>8.2f} timer")
    if result.changeover_details:
        print(f"  Omstillinger:")
        for wc, det in result.changeover_details.items():
            print(f"     {wc}: {det['num_changeovers']} omstillinger, "
                  f"{det['total_minutes']:.0f} min, "
                  f"{det['total_cost']:.2f} kr")
    print(f"{sep}\n")



# ---------------------------------------------------------------------
#  9. HOVEDFUNKSJON
# ---------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Beregn produktkost fra Excel-modell",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Eksempler:
  python kostberegning.py --excel Produksjonsmodell_Testdata.xlsx
  python kostberegning.py --test
  python kostberegning.py --scenario "Normal"
  python kostberegning.py --product FG001
  python kostberegning.py --no-json

        """,
    )
    parser.add_argument("--test", action="store_true",
                        help="Bruk innebygget testdata (Kodal Hovleri)")
    parser.add_argument("--scenario", type=str, default=None,
                        help="Simuler et spesifikt scenario")
    parser.add_argument("--product", type=str, default=None,
                        help="Beregn kun for ett produkt (Item No)")
    parser.add_argument("--no-json", action="store_true",
                        help="Ikke eksporter til JSON")
    parser.add_argument("--json-dir", type=str, default=".",
                        help="Mappe for JSON-eksport (standard: .)")
    parser.add_argument("--excel", type=str, default=None,
                        help="Excel-fil med data (standard: Produksjonsmodell_Mal.xlsx)")

    args = parser.parse_args()


    # -- Last data ------------------------------------------------
    if args.test:
        print("  >>> Bruker testdata (Kodal Hovleri)")
        data = create_test_data()
    else:
        excel_file = args.excel if args.excel else "Produksjonsmodell_Mal.xlsx"
        if not Path(excel_file).exists():
            print(f"  >>> Finner ikke {excel_file}")
            print(f"     Bruk --test for a kjore med testdata")
            print(f"     Eller --excel Produksjonsmodell_Testdata.xlsx")
            sys.exit(1)
        print(f"  >>> Laster data fra {excel_file}")
        data = ExcelData(excel_file)

    print(f"     {len(data.products)} produkter, {len(data.work_centers)} arbeidssentre, "
          f"{len(data.bom_lines)} BOM-linjer, {len(data.routing_lines)} routing-linjer")

    calculator = CostCalculator(data)

    # -- Scenario-simulering (UTKOMMENTERT - flyttet til Marimo) --
    if args.scenario:
        print(f"\n  >>> Scenario-simulering er flyttet til Marimo.")
        print(f"     Kjor: marimo edit scenario_analyse.py")
        print(f"     for a apne analyseverktoyet.")
        return

    # -- Produktkalkyle -------------------------------------------
    if args.product:
        print(f"\n  >>> Beregner kost for: {args.product}")
        # Kjor calculate_all() forst for a fa dynamisk cost roll-up
        # (beregner alle FG/HF i avhengighetsrekkefolge)
        calculator.calculate_all()
        # Hent resultater for det spesifikke produktet (alle fabrikker)
        results = calculator.calculate_product_costs(args.product)
        if results:
            for r in results:
                print_product_cost(r)
            if not args.no_json:
                json_path = Path(args.json_dir) / "produktkalkyle.json"
                export_product_costs_to_json(results, str(json_path))
        else:
            print(f"  >>> Produkt '{args.product}' ikke funnet eller har ingen routing.")
        return

    # -- Full kalkyle ---------------------------------------------
    print(f"\n  >>> Beregner kost for alle ferdigvarer/halvfabrikata...")
    results = calculator.calculate_all()

    if not results:
        print("  >>> Ingen produkter a beregne.")
        return

    for r in results:
        print_product_cost(r)

    # -- JSON-eksport (kun produktkalkyle, scenario er flyttet til Marimo) --
    if not args.no_json:
        json_prod = Path(args.json_dir) / "produktkalkyle.json"
        export_product_costs_to_json(results, str(json_prod))

    print("  >>> Ferdig!\n")




if __name__ == "__main__":
    main()


