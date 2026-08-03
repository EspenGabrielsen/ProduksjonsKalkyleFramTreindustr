#!/usr/bin/env python3
"""
optimization_engine.py — MILP-produksjonsoptimering med flernivå-BOM
====================================================================

Frittstående optimeringsmotor som finner den MEST LØNNSOMME
produksjonsplanen for en gitt etterspørsel, på tvers av:

  - Flernivå-BOM (Skurlast → Ubehandlet → Grunnet → Malt)
  - Flere lokasjoner (hvor skal hvert steg produseres?)
  - Flere perioder (når skal det produseres, og hvor mye lager?)
  - Transport mellom lokasjoner
  - Kapasitetsbegrensninger per arbeidssenter

Viktige prinsipper:
  - Eksisterende kode (kostberegning.py, varekost_app.py, excel_bridge.py)
    endres ALDRI. Motoren importerer og gjenbruker SqliteData + CostCalculator.
  - Etterspørsel trigges fra toppen — kun sluttprodukter har demand.
  - Materialbalanse tvinger frem produksjon av alle nødvendige delprodukter.

Bruk:
  python src/optimization_engine.py --test                  # Syntetisk test-case
  python src/optimization_engine.py --db path/to.db         # Kjør mot ekte DB
  python src/optimization_engine.py --db path/to.db --output resultat.json
"""

import argparse
import json
import os
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import sys, io

# Sørg for at src/ er på sys.path (gjør at modulen kan kjøres fra hvor som helst)
_src = Path(__file__).resolve().parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

import pulp

from data_repo import DataRepo
from kostberegning import SqliteData, CostCalculator
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ──────────────────────────────────────────────────────────────────────
#  Datastrukturer
# ──────────────────────────────────────────────────────────────────────

@dataclass
class DemandRecord:
    """Én etterspørselslinje fra demand-tabellen."""
    product_id: str
    period: int
    quantity: float
    location_code: str
    customer_region: str = ""


@dataclass
class BOMStructureRecord:
    """Én flernivå-BOM-avhengighet fra bom_structure-tabellen.

    Betyr: For å lage 1 enhet av parent_product_id, kreves det
    yield_factor enheter av child_product_id.
    """
    parent_product_id: str
    child_product_id: str
    yield_factor: float = 1.0


@dataclass
class OptimizationResult:
    """Resultat fra optimeringskjøringen."""
    status: str
    total_cost: float
    production_plan: list[dict] = field(default_factory=list)
    transport_plan: list[dict] = field(default_factory=list)
    inventory_levels: list[dict] = field(default_factory=list)
    batch_decisions: list[dict] = field(default_factory=list)
    cost_breakdown: dict = field(default_factory=dict)
    solve_time_seconds: float = 0.0

    def to_dict(self) -> dict:
        return {
            "status": self.status,
            "total_cost": round(self.total_cost, 4),
            "solve_time_seconds": round(self.solve_time_seconds, 3),
            "cost_breakdown": {
                k: round(v, 4) for k, v in self.cost_breakdown.items()
            },
            "production_plan": self.production_plan,
            "transport_plan": self.transport_plan,
            "batch_decisions": self.batch_decisions,
            "inventory_levels": self.inventory_levels,
        }


# ──────────────────────────────────────────────────────────────────────
#  Optimeringsmotor
# ──────────────────────────────────────────────────────────────────────

class OptimizationEngine:
    """MILP-produksjonsoptimering bygget på PuLP.

    Leser data fra SQLite via SqliteData (stamdata) + DataRepo
    (demand, bom_structure, transport_ruter), beregner enhetskostnader
    via CostCalculator, og setter opp en heltallsmodell som finner den
    optimale produksjons-/transport-/lagerplanen.
    """

    def __init__(
        self,
        db_path: Optional[str] = None,
        days_per_period: float = 5.0,
        holding_cost_pct: float = 0.0,
        include_transport: bool = True,
        forecast_weeks: int = 12,
        forecast_max_setup_pct: float = 30.0,
    ):
        self.db_path = db_path
        self.days_per_period = days_per_period
        self.holding_cost_pct = holding_cost_pct
        self.include_transport = include_transport
        self.forecast_weeks = forecast_weeks
        self.forecast_max_setup_pct = forecast_max_setup_pct

        # Database + datasource (gjenbruk av eksisterende kode)
        self.db = DataRepo(db_path)
        self.db.initialize()
        self.data = SqliteData(db_path)

        # Lastede data
        self.demand: list[DemandRecord] = []
        self.bom_structure: list[BOMStructureRecord] = []
        self.transport_ruter: dict[tuple[str, str], dict] = {}

        # Beregnede oppslag
        self.unit_costs: dict[tuple[str, str], float] = {}   # (p,l) → operasjonskost - byproduct per enhet
        self.full_costs: dict[tuple[str, str], dict] = {}    # (p,l) → {operations, material, setup, byproduct, gross, net}
        self.fixed_setup: dict[tuple[str, str], float] = {}  # (p,l) → fast setup-kost per batch
        self.purchase_costs: dict[str, float] = {}           # p → statisk innkjøpskost (råvarer)
        self.lm_per_m3: dict[str, float] = {}                # p → LM/M3-konvertering
        self.route_costs: dict[tuple[str, str, str], float] = {}  # (p,from,to) → transportkost per enhet

        # Changeover-data (familie-basert, fra changeover_matrix-tabellen + FTI-fallback)
        self.changeover_minutes: dict[tuple[str, str, str], float] = {}  # (wc, from_fam, to_fam) → min
        self.changeover_cost: dict[tuple[str, str, str], float] = {}     # (wc, from_fam, to_fam) → kr
        self.family_of: dict[str, str] = {}                              # produkt → FTI-familie
        self.family_dims: dict[str, tuple] = {}                          # familie → (tykkelse, bredde)
        self.family_penalty: dict[tuple[str, str], float] = {}           # (wc, familie) → snitt changeover-min per familie

    # ── Datalasting ───────────────────────────────────────────────

    def load_all(self):
        """Last alle nødvendige data fra SQLite."""
        self._load_transport_ruter()  # må lastes FØR demand (validator sjekker ruter)
        self._load_demand()
        self._load_bom_structure()
        self._load_changeover()
        self._calculate_costs()
        self._calculate_lm_per_m3()
        self._calculate_route_costs()

    def validate_demand(self) -> tuple[list[dict], list[DemandRecord]]:
        """Valider alle demand-linjer i databasen FØR MILP-en bygges.

        Returnerer:
            (feil_liste, gyldig_demand)

        feil_liste: liste med dicts:
            {"product_id", "period", "location", "quantity", "reason"}

        gyldig_demand: liste med DemandRecord som kan modelleres.
        """
        rows = self.db.conn.execute(
            "SELECT product_id, period, quantity, location_code, customer_region "
            "FROM demand ORDER BY period, product_id"
        ).fetchall()

        prod_locs = {
            wc.location_code for wc in self.data.work_centers if wc.location_code
        }

        co_products = {
            bl.co_product_item_no for bl in self.data.bom_lines if bl.co_product_item_no
        }

        prod_locs_by_item: dict[str, set[str]] = {}
        for rl in self.data.routing_lines:
            wc = self.data.work_center(rl.work_center_code)
            if wc and wc.location_code:
                prod_locs_by_item.setdefault(rl.item_no, set()).add(wc.location_code)

        routing_items = {rl.item_no for rl in self.data.routing_lines}
        purchasable = {ic.item_no for ic in self.data.item_costs}

        lm_items = {
            p.item_no for p in self.data.products
            if _finn_lm_per_m3_robust(self.data, p.item_no)
        }

        feil = []
        gyldig = []

        for r in rows:
            pid = r["product_id"]
            loc = r["location_code"]
            qty = float(r["quantity"] or 0)
            period = int(r["period"])

            if qty <= 0:
                continue  # ignorere tomme linjer uten feilmelding

            reason = None

            # 1. Lokasjon må ha produksjonsmulighet
            if loc not in prod_locs:
                reason = "UKJENT_LOKASJON"
            # 2. Co-produkter håndteres ikke som selvstendige
            elif pid in co_products:
                reason = "CO_PRODUKT_HÅNDTERES_IKKE"
            # 3. Produkt må ha routing (produseres) ELLER kjøpspris (kjøpes)
            elif pid not in routing_items and pid not in purchasable:
                reason = "VERKEN_PRODUKSJON_ELLER_KJØP"
            # 4. Hvis produktet har routing men produksjonslokasjonen er 
            #    forskjellig fra demand-lokasjonen, kreves transport:
            #    - Produktet må ha LM/M3-konvertering
            #    - Det må finnes en transportrute fra produksjonslokasjon til demand-lokasjon
            elif pid in prod_locs_by_item and loc not in prod_locs_by_item[pid]:
                if pid not in lm_items:
                    reason = "MANGLER_LM_M3_KONVERTERING"
                else:
                    can_reach = any(
                        (pl, loc) in self.transport_ruter for pl in prod_locs_by_item[pid]
                    )
                    if not can_reach:
                        reason = "MANGLER_TRANSPORTRUTE"
            # 5. Kjøpsvarer (ingen routing) kan leveres overalt → OK

            if reason:
                feil.append({
                    "product_id": pid,
                    "period": period,
                    "location": loc,
                    "quantity": qty,
                    "reason": reason,
                })
                continue

            gyldig.append(DemandRecord(
                product_id=pid,
                period=period,
                quantity=qty,
                location_code=loc,
                customer_region=r["customer_region"] or "",
            ))

        return feil, gyldig

    def _load_demand(self):
        """Last demand, valider først, og behold kun gyldige linjer."""
        feil, gyldig = self.validate_demand()

        if feil:
            print(f"[VALIDERING] {len(feil)} demand-linjer filtrert bort:")
            by_reason: dict[str, int] = {}
            for f in feil:
                by_reason[f["reason"]] = by_reason.get(f["reason"], 0) + 1
            for reason, cnt in sorted(by_reason.items()):
                print(f"   ❌ {reason}: {cnt}")

        self.demand = gyldig
        self.validation_errors = feil

    # ── Prognose fra historisk salg ──────────────────────────────

    def _load_historical_sales(self) -> dict[tuple[str, str], float]:
        """Last historisk salg fra historical_sales-tabellen.

        Returnerer:
            dict {(product_id, location_code): samlet_volum_siste_52_uker}
        """
        rows = self.db.conn.execute(
            "SELECT product_id, period, quantity, location_code "
            "FROM historical_sales ORDER BY period"
        ).fetchall()
        totals: dict[tuple[str, str], float] = {}
        for r in rows:
            qty = float(r["quantity"] or 0)
            if qty <= 0:
                continue
            key = (r["product_id"], r["location_code"])
            totals[key] = totals.get(key, 0.0) + qty
        return totals

    def _forecast_records(self) -> list[DemandRecord]:
        """Bygg prognose-records fra historisk salg.

        For hvert (produkt, lokasjon) med historisk salg beregnes
        gjennomsnittlig ukentlig salg. Dette projiseres fremover i de neste
        `forecast_weeks` periodene, men KUN i perioder som ikke allerede
        har faktisk demand (åpne ordre). Dette gir modellen et bilde av
        fremtidig etterspørsel — slik at den kan velge riktig batch-størrelse:
        en stor batch nå sparer setup-kost, men påløper lagerholdskost.

        Setup-kost-terskel:
          Hvis setup-kostnaden er uforholdsmessig høy i forhold til
          operasjonskosten for én ukes salg, skal prognosen IKKE genereres
          hver uke. I stedet slås N uker sammen til én prognoseperiode,
          der N = ceil(min_mengde / snitt_uke_salg). Dette unngår at
          modellen foreslår meningsløse mikrobatcher (f.eks. 0.1 LM).
          Merknaden "x uker forsinket" kan leses som at produktet får
          prognose hver N-te periode istedenfor hver uke.

        Returns:
            Liste med DemandRecord for prognose-periodene.
        """
        totals = self._load_historical_sales()
        if not totals:
            return []

        # Finn siste faktiske demand-periode — prognoser starter etter dette.
        max_demand_period = max(
            (d.period for d in self.demand), default=0
        )

        # Gjennomsnittlig ukentlig salg per (produkt, lokasjon)
        HIST_UKER = 52
        avg = {key: v / HIST_UKER for key, v in totals.items()}

        # Setup-kost-terskel (Alternativ A: kun operasjonskost/routing, eksklusiv råvare)
        # 0% = ubegrenset (alle prognoselinjer genereres)
        setup_pct = self.forecast_max_setup_pct / 100.0

        # Prognose-perioder: fortsettelse etter siste faktiske demand.
        records: list[DemandRecord] = []
        for (pid, loc), uke_qty in avg.items():
            if uke_qty <= 0:
                continue

            # Beregn minimum økonomisk mengde per batch:
            #   min_mengde = setup_kost / (enhetskost × max_setup_pct)
            # Hvis snitt ukentlig salg < min_mengde, slås N uker sammen.
            unit_cost = self.unit_costs.get((pid, loc), 0.0)
            setup_cost = self.fixed_setup.get((pid, loc), 0.0)

            if setup_pct > 0 and unit_cost > 0 and setup_cost > 0:
                min_mengde = setup_cost / (unit_cost * setup_pct)
                if uke_qty < min_mengde:
                    # Slå sammen N = ceil(min_mengde / uke_qty) uker.
                    # Produktet får da prognose kun hver N-te periode.
                    import math
                    n = max(1, math.ceil(min_mengde / uke_qty))
                    for t in range(max_demand_period + 1,
                                   max_demand_period + 1 + self.forecast_weeks):
                        # Kun hver n-te periode fra start
                        if (t - (max_demand_period + 1)) % n != 0:
                            continue
                        records.append(DemandRecord(
                            product_id=pid,
                            period=t,
                            quantity=uke_qty * n,  # N ukers behov samlet
                            location_code=loc,
                            customer_region="prognose",
                        ))
                    continue

            # Normaltilfellet: ukentlig prognose
            for t in range(max_demand_period + 1, max_demand_period + 1 + self.forecast_weeks):
                records.append(DemandRecord(
                    product_id=pid,
                    period=t,
                    quantity=uke_qty,
                    location_code=loc,
                    customer_region="prognose",
                ))
        return records

    def _expand_records(
        self, records: list[DemandRecord]
    ) -> dict[tuple[str, str, int], float]:
        """Pre-prosesser demand via BOM-traversering.

        Går gjennom records og følger BOM-kjeden rekursivt for å finne
        totalbehov per produkt, periode og lokasjon.

        Eksempel:
          Demand: JJ16123EF 620 LM + JJ16123EH 620 LM + JJ16123 1 LM
          → JJ16123GH trengs: 620 LM (fra EH)
          → JJ16123GF trengs: 620 LM (fra EF)
          → JJ16123 trengs: 620 + 620 + 1 = 1241 LM
          → RM_38x125 trengs: 1241 × (1/420.8754) = 2.95 M3

        Args:
            records: Liste med DemandRecord (faktisk demand + prognose)

        Returns:
            dict {(product_id, location_code, period): total_behov}
        """
        total: dict[tuple[str, str, int], float] = {}

        # Hvilke produkter kan produseres hvor, og hvilke kan kjøpes?
        routing_items = {rl.item_no for rl in self.data.routing_lines}
        purchasable = {ic.item_no for ic in self.data.item_costs}
        prod_locs_by_item: dict[str, set[str]] = {}
        for rl in self.data.routing_lines:
            wc = self.data.work_center(rl.work_center_code)
            if wc and wc.location_code:
                prod_locs_by_item.setdefault(rl.item_no, set()).add(wc.location_code)

        def can_cover(item_no: str, loc: str) -> bool:
            """Kan dette produktet dekkes på demand-lokasjonen?"""
            if item_no not in routing_items:
                # Kjøpsvare kan kjøpes overalt (hvis den har pris)
                return item_no in purchasable
            prod_locs = prod_locs_by_item.get(item_no, set())
            if loc in prod_locs:
                return True
            # Må transporteres inn → krever LM/M3 + transportrute
            if not _finn_lm_per_m3_robust(self.data, item_no):
                return False
            return any((pl, loc) in self.transport_ruter for pl in prod_locs)

        def traverse(item_no: str, loc: str, period: int, qty: float,
                     visited: Optional[set[str]] = None):
            """Rekursiv traversering med sirkelbeskyttelse.

            Hopp over produkter som verken kan produseres eller kjøpes på
            demand-lokasjonen (ignorer det vi ikke kan gjøre noe med).
            """
            if visited is None:
                visited = set()
            # Sirkelbeskyttelse: unngå uendelig løkke
            if item_no in visited:
                return
            visited.add(item_no)

            # Hopp over produkter som ikke kan dekkes på denne lokasjonen
            if not can_cover(item_no, loc):
                return

            key = (item_no, loc, period)
            total[key] = total.get(key, 0.0) + qty

            # Følg BOM-kjeden
            for bl in self.data.bom_for(item_no):
                # Hopp over selve co-produktet (B-varen produseres som biprodukt).
                # Hovedkomponenten (bl.component_item_no) er fortsatt nødvendig!
                if bl.co_product_item_no and bl.co_product_item_no == bl.component_item_no:
                    continue
                qty_per = bl.quantity_per if bl.quantity_per and bl.quantity_per > 0 else 1.0
                scrap = bl.scrap_pct if bl.scrap_pct else 0.0
                child_qty = qty * (1.0 / qty_per) * (1.0 + scrap / 100.0)
                traverse(bl.component_item_no, loc, period, child_qty,
                         visited.copy())

        for d in records:
            traverse(d.product_id, d.location_code, d.period, d.quantity)

        # Filtrer bort mikro-behov (< 1 enhet) etter BOM-ekspansjon.
        # Disse kommer fra åpne ordrer med ubetydelige volumer (f.eks. 0.1 LM)
        # og fra avledet behov i BOM-kjeden. Mikro-behov gir meningsløse
        # produksjonsforslag som aldri kan forsvares økonomisk.
        MIN_BEHOV = 1.0
        total = {k: v for k, v in total.items() if v >= MIN_BEHOV}

        return total

    def _expand_demand(self) -> dict[tuple[str, str, int], float]:
        """Pre-prosesser demand via BOM-traversering.

        Slår sammen faktisk demand (åpne ordre) med prognose fra historisk
        salg, og følger BOM-kjeden rekursivt for å finne totalbehov per
        produkt, periode og lokasjon.

        Returns:
            dict {(product_id, location_code, period): total_behov}
        """
        forecast = self._forecast_records()
        records = list(self.demand) + forecast
        return self._expand_records(records)

    def _load_bom_structure(self):
        """Ingen egen BOM-tabell — demand pre-prosesseres via _expand_demand()."""
        self.bom_structure = []

    def _load_transport_ruter(self):
        rows = self.db.conn.execute(
            "SELECT from_loc, to_loc, cost_per_m3, distance_km, hours "
            "FROM transport_ruter ORDER BY from_loc, to_loc"
        ).fetchall()
        self.transport_ruter = {
            (r["from_loc"], r["to_loc"]): {
                "cost_per_m3": float(r["cost_per_m3"] or 0),
                "distance_km": float(r["distance_km"] or 0),
                "hours": float(r["hours"] or 0),
            }
            for r in rows
        }

    def _load_changeover(self):
        """Last changeover_matrix fra SQLite + bygg FTI-familie-oppslag.

        Bruker changeover_matrix-tabellen hvis den er fylt. Hvis ikke,
        genereres estimater fra FTI-nummerstrukturen (_changeover_minutes).
        """
        # Familie per produkt (kun produserbare produkter)
        self.family_of = {}
        for rl in self.data.routing_lines:
            if rl.item_no not in self.family_of:
                fam = _fti_family(rl.item_no)
                self.family_of[rl.item_no] = fam
                # Dimensjoner for familien
                digits = "".join(ch for ch in fam if ch.isdigit())
                if len(digits) >= 4:
                    self.family_dims[fam] = (int(digits[:2]), int(digits[-3:]))

        # Oppslag: work_center_code → total_cost_hour (for å beregne changeover-kost ALLTID)
        wc_hour_cost = {
            wc.code: wc.total_cost_hour for wc in self.data.work_centers
        }

        # Les changeover_matrix fra DB — OPTIONALT overstyringslag for spesifikke par.
        # Kun TID lagres i tabellen. Kostnad beregnes ALLTID i koden:
        #   changeover_cost = (changeover_minutes / 60) × work_centers.total_cost_hour
        # Dette unngår dobbelt vedlikehold av kostnadsdata.
        rows = self.db.conn.execute(
            "SELECT work_center_code, from_family, to_family, changeover_minutes "
            "FROM changeover_matrix ORDER BY work_center_code, from_family, to_family"
        ).fetchall()

        # Hvilke familier kjøres på hvilket arbeidssenter (fra routing)
        wc_families: dict[str, set[str]] = {}
        for rl in self.data.routing_lines:
            wc_families.setdefault(rl.work_center_code, set()).add(
                self.family_of.get(rl.item_no, _fti_family(rl.item_no))
            )

        if rows:
            # Bruk eksplisitt matrise (manuelt fylt i changeover_matrix) — kun tid
            for r in rows:
                key = (r["work_center_code"], r["from_family"], r["to_family"])
                minutes = float(r["changeover_minutes"] or 0)
                self.changeover_minutes[key] = minutes
                # Kost beregnes ALLTID fra arbeidssenterets timepris
                cost_per_hour = wc_hour_cost.get(r["work_center_code"], 0.0)
                self.changeover_cost[key] = (minutes / 60.0) * cost_per_hour
        else:
            # Fallback: generer estimater fra FTI-strukturen per arbeidssenter,
            # kun for familier som faktisk kjøres på dette arbeidssenteret
            for wc in self.data.work_centers:
                families = sorted(wc_families.get(wc.code, set()))
                if len(families) < 2:
                    continue
                cost_per_hour = wc.total_cost_hour
                for fa in families:
                    for fb in families:
                        if fa == fb:
                            continue
                        minutes = _changeover_minutes(fa, fb, self.family_dims)
                        key = (wc.code, fa, fb)
                        self.changeover_minutes[key] = minutes
                        self.changeover_cost[key] = (minutes / 60.0) * cost_per_hour

        # Snitt-straff per (wc, familie): hvor mange minutter koster det i snitt
        # å starte produksjon av denne familien når den ikke allerede kjører?
        for wc in self.data.work_centers:
            for fam in set(self.family_of.values()):
                times = [
                    self.changeover_minutes.get((wc.code, other, fam), 0.0)
                    for other in set(self.family_of.values())
                    if other != fam
                ]
                times = [t for t in times if t > 0]
                if times:
                    self.family_penalty[(wc.code, fam)] = sum(times) / len(times)
                else:
                    self.family_penalty[(wc.code, fam)] = 0.0

    def _calculate_costs(self):
        """Beregn enhetskostnader via eksisterende CostCalculator.

        For produserbare produkter (har routing):
            unit_cost = operasjonskost - biproduktverdi  (materialkost ligger i
            modellen via materialbalansen, slik at råvarer ikke dobbelttelles)
            fixed_setup = fast setup-kost per batch-start

        For øvrige produkter (råvarer/innkjøp):
            purchase_cost = statisk item_cost
        """
        calculator = CostCalculator(self.data)
        results = calculator.calculate_all()

        for r in results:
            key = (r.product_no, r.location_code)
            if r.location_code:
                # Operasjonskost minus biproduktverdi = "verdiskaping" per enhet i dette steget
                self.unit_costs[key] = r.operation_cost - r.by_product_value
                # Fast setup-kost per batch (setup per enhet × batch-størrelse)
                setup_per_batch = 0.0
                for od in r.operation_details:
                    setup_per_batch += od.setup_cost_per_unit * od.batch_size
                self.fixed_setup[key] = setup_per_batch

                # Lagre fullkost-oppslag for referanse i output
                self.full_costs[key] = {
                    "operations": r.operation_cost,
                    "material": r.material_cost,
                    "setup": r.setup_cost,
                    "byproduct": r.by_product_value,
                    "gross": r.gross_production_cost,
                    "net": r.net_production_cost,
                }

        # Innkjøpskost for produkter som ikke produseres (råvarer)
        for p in self.data.products:
            cost = self.data.item_cost(p.item_no)
            if cost is not None:
                self.purchase_costs[p.item_no] = cost.unit_cost

    def _calculate_lm_per_m3(self):
        """Finn LM/M3-konvertering for alle produkter i modellen.

        Bruker lokal robust versjon som følger kun LM/M3-kjeden
        (hopper over maling med LTR-enhet).
        """
        for p in self.data.products:
            lm = _finn_lm_per_m3_robust(self.data, p.item_no)
            if lm and lm > 0:
                self.lm_per_m3[p.item_no] = lm

    def _calculate_route_costs(self):
        """Beregn transportkost per enhet for alle produkter/ruter.

        cost_per_m3 hentes fra transport_ruter, og konverteres til
        kost per LM ved å dele på produktets LM/M3.
        """
        if not self.include_transport:
            return
        for p, lm in self.lm_per_m3.items():
            for (frm, to), rute in self.transport_ruter.items():
                if lm > 0:
                    self.route_costs[(p, frm, to)] = rute["cost_per_m3"] / lm

    # ── Modellbygging ─────────────────────────────────────────────

    def build_model(self) -> pulp.LpProblem:
        """Bygg PuLP-modellen (MILP)."""

        # ── Sett ──────────────────────────────────────────────────
        # Pre-prosesser demand: følg BOM-kjeden for å finne totalbehov
        # per produkt, periode og lokasjon. Ingen materialbalanse i MILP!
        expanded_demand = self._expand_demand()

        # Produkter som inngår i modellen (fra ekspandert demand)
        model_products: set[str] = set()
        for (p, _l, _t) in expanded_demand:
            model_products.add(p)

        # Hvis ingen produkter finnes, avslutt med tom modell
        if not model_products:
            raise ValueError("Ingen demand-data funnet. Fyll demand-tabellen først.")

        # Perioder (sortert)
        periods = sorted({t for (_p, _l, t) in expanded_demand})
        if not periods:
            raise ValueError("Ingen perioder i demand-data.")

        # Lokasjoner (fabrikker + demand-lokasjoner)
        locations: set[str] = {
            l.code for l in self.data.locations if l.location_type == "Factory"
        }
        locations |= {l for (_p, l, _t) in expanded_demand}

        # Hvilke produkter kan produseres (har routing)?
        routing_items = {rl.item_no for rl in self.data.routing_lines}

        # Produksjonslokasjoner per produkt
        prod_locations: dict[str, set[str]] = {}
        for p in model_products:
            if p in routing_items:
                prod_locations[p] = set(self.data.routing_locations_for(p)) & locations
            else:
                prod_locations[p] = set(locations)  # kan "kjøpes" overalt

        # ── Variabel-nøkler ───────────────────────────────────────
        X_keys = [
            (p, l, t)
            for p in sorted(model_products)
            for l in sorted(prod_locations[p])
            for t in periods
        ]
        Y_keys = [
            (p, l, t) for (p, l, t) in X_keys if p in routing_items
        ]
        # Lager-variabler: I[p,l,t] = utgående lager etter periode t
        # Bruker ALLE lokasjoner (ikke kun prod_locations) fordi demand
        # kan ligge på lokasjoner der produktet ikke produseres.
        I_keys = [
            (p, l, t)
            for p in sorted(model_products)
            for l in sorted(locations)
            for t in periods
        ]
        T_keys = []
        if self.include_transport and len(locations) > 1:
            for (p, frm, to) in sorted(self.route_costs.keys()):
                if p in model_products and frm in locations and to in locations and frm != to:
                    for t in periods:
                        T_keys.append((p, frm, to, t))

        # ── Variabler ─────────────────────────────────────────────
        X = {}
        for (p, l, t) in X_keys:
            X[(p, l, t)] = pulp.LpVariable(
                f"Prod_{_safe(p)}_{l}_{t}", lowBound=0.0, cat="Continuous"
            )

        Y = {}
        for (p, l, t) in Y_keys:
            Y[(p, l, t)] = pulp.LpVariable(
                f"Batch_{_safe(p)}_{l}_{t}", cat="Binary"
            )

        I = {}
        for (p, l, t) in I_keys:
            I[(p, l, t)] = pulp.LpVariable(
                f"Lager_{_safe(p)}_{l}_{t}", lowBound=0.0, cat="Continuous"
            )

        T = {}
        for (p, frm, to, t) in T_keys:
            T[(p, frm, to, t)] = pulp.LpVariable(
                f"Trans_{_safe(p)}_{frm}_{to}_{t}", lowBound=0.0, cat="Continuous"
            )

        model = pulp.LpProblem("Hovleri_MultiStage_Optimering", pulp.LpMinimize)

        # ── Familie-variabler (changeover) ────────────────────────
        # YF[f, wc, t] = 1 hvis minst ett produkt i familie f produseres på
        # arbeidssenter wc (lokasjonen wc tilhører) i periode t.
        # Dette brukes til å trekke changeover-straff fra kapasiteten når
        # flere ulike familier kjøres på samme WC i samme periode.
        wc_loc = {
            wc.code: wc.location_code
            for wc in self.data.work_centers
            if wc.location_code
        }
        routing_wc = {
            rl.item_no: rl.work_center_code for rl in self.data.routing_lines
        }

        # Familier som faktisk produseres i modellen (har routing + inngår i model_products)
        model_families = {
            self.family_of.get(p, _fti_family(p))
            for p in model_products
            if p in routing_wc
        }

        YF_keys = [
            (f, wc_code, t)
            for f in sorted(model_families)
            for wc_code in sorted(wc_loc)
            for t in periods
        ]
        YF = {}
        for (f, wc_code, t) in YF_keys:
            YF[(f, wc_code, t)] = pulp.LpVariable(
                f"Fam_{_safe(f)}_{wc_code}_{t}", cat="Binary"
            )

        # Kobling: Y[p,l,t] → YF[f, wc, t] (hvis p produseres, er familien aktiv)
        for (p, l, t) in Y_keys:
            fam = self.family_of.get(p, _fti_family(p))
            wc_code = routing_wc.get(p)
            if wc_code is None or (fam, wc_code, t) not in YF:
                continue
            if wc_loc.get(wc_code) == l:
                model += Y[(p, l, t)] <= YF[(fam, wc_code, t)], (
                    f"FamKoble_{_safe(p)}_{l}_{t}"
                )
            else:
                # Produkt produseres på en annen lokasjon enn wc → bruk alternativt
                # (dette skjer kun hvis routing_workcenter er på feil lokasjon; hopp over)
                pass

        # Changeover-straff i kapasiteten: for hver aktiv familie utover den første
        # trekkes en estimert omstillingstid fra kapasiteten på arbeidssenteret.
        # Dette straffer blanding av mange familier i samme periode (Gemini-Tilnærming 3).
        # Snittstraffen per (wc, familie) kommer fra self.family_penalty.
        # Vi knytter en "startkost" til hver familie via YF — første familie koster
        # ingenting ekstra (den settes opp fra null), men hver påfølgende familie
        # koster FAMILY_PENALTY i tapt kapasitet.
        # Forenkling: straff per familie som IKKE er første, approx = 0.5 × snittstraff.
        # Dette er en konservativ tilnærming uten å innføre Z[i,j]-sekvensering.

        # ── Objektfunksjon ────────────────────────────────────────
        def _obj_cost(p: str, l: str) -> float:
            """Kost per enhet i objektfunksjonen for produksjon av p på l."""
            if (p, l) in self.unit_costs:
                return self.unit_costs[(p, l)]
            return self.purchase_costs.get(p, 0.0)

        objective_terms = []

        # 1. Produksjons-/innkjøpskost
        prod_cost_expr = pulp.lpSum(
            X[(p, l, t)] * _obj_cost(p, l) for (p, l, t) in X_keys
        )
        objective_terms.append(prod_cost_expr)

        # 2. Fast setup-kost per batch-start
        setup_expr = pulp.lpSum(
            Y[(p, l, t)] * self.fixed_setup.get((p, l), 0.0) for (p, l, t) in Y_keys
        )
        objective_terms.append(setup_expr)

        # 3. Transportkost
        if T_keys:
            trans_expr = pulp.lpSum(
                T[(p, frm, to, t)] * self.route_costs[(p, frm, to)]
                for (p, frm, to, t) in T_keys
            )
            objective_terms.append(trans_expr)

        # 4. Lagerholdskost (valgfritt, % av enhetskost per periode)
        #    Dette er nøkkelen til å balansere batch-størrelse: stor batch →
        #    lager → kostnad. Modellen velger da kun å produsere stort når
        #    setup-besparelsen overstiger lagerholdskostnaden.
        if self.holding_cost_pct > 0:
            hold_expr = pulp.lpSum(
                I[(p, l, t)] * _obj_cost(p, l) * (self.holding_cost_pct / 100.0)
                for (p, l, t) in I_keys
            )
            objective_terms.append(hold_expr)

        model += pulp.lpSum(objective_terms), "Total_Kost"

        # ── Big-M batch-kobling ───────────────────────────────────
        total_demand = sum(expanded_demand.values())
        M = max(total_demand * 10, 100000.0)  # stor nok verdi

        for (p, l, t) in Y_keys:
            model += X[(p, l, t)] <= M * Y[(p, l, t)], (
                f"Rigg_{_safe(p)}_{l}_{t}"
            )

        # ── Kapasitet per arbeidssenter ───────────────────────────
        # Beregn run- og setup-tid per (produkt, work_center) fra CostCalculator-detaljer
        run_hours: dict[tuple[str, str], float] = {}
        setup_hours: dict[tuple[str, str], float] = {}
        for (p, l), setup in self.fixed_setup.items():
            pass  # setup per batch håndteres i objektfunksjonen

        # Samle belastning per (p, wc) fra operation_details
        for (p, l) in self.unit_costs:
            # Vi bygger nye detaljer fra routing direkte for kapasitet
            pass

        # Beregn belastning fra routing-linjer (konsistent med CostCalculator)
        # run_hours per (item, wc) og setup_hours per (item, wc)
        for rl in self.data.routing_lines:
            wc = self.data.work_center(rl.work_center_code)
            if wc is None:
                continue
            key = (rl.item_no, rl.work_center_code)
            run_hours[key] = run_hours.get(key, 0.0) + rl.run_time_minutes / 60.0
            setup_hours[key] = setup_hours.get(key, 0.0) + rl.setup_time_minutes / 60.0

        # wc → lokasjon
        wc_loc = {
            wc.code: wc.location_code
            for wc in self.data.work_centers
            if wc.location_code
        }
        # wc → kapasitet per periode (timer × effektivitet × dager)
        wc_capacity: dict[str, float] = {}
        for wc in self.data.work_centers:
            eff = (wc.effective_capacity_pct or 100.0) / 100.0
            wc_capacity[wc.code] = wc.capacity_hours_day * eff * self.days_per_period

        # Bygg kapasitetsconstraints: for hvert (wc, t) summeres belastning
        # over alle produkter som produseres på wc sin lokasjon, PLUS
        # changeover-straff for hver ekstra produktfamilie som kjøres.
        # Prinsipp (Gemini-tilnærming 3): Blir det kjørt K familier på samme
        # WC i samme periode, kreves minimum (K-1) omstillinger. Vi straffer
        # hver aktiv familie med 0.5 × snitt-changeover (ca. (K-1) × snitt
        # når K >= 2) — trekkes fra kapasiteten som tapt tid.
        for wc_code, loc in wc_loc.items():
            if loc not in locations:
                continue
            cap = wc_capacity.get(wc_code, 0.0)
            if cap <= 0:
                continue
            for t in periods:
                load_expr = []
                for rl in self.data.routing_lines:
                    if rl.work_center_code != wc_code:
                        continue
                    p = rl.item_no
                    if (p, loc, t) not in X:
                        continue
                    run_h = run_hours.get((p, wc_code), 0.0)
                    setup_h = setup_hours.get((p, wc_code), 0.0)
                    if run_h > 0:
                        load_expr.append(run_h * X[(p, loc, t)])
                    if setup_h > 0 and (p, loc, t) in Y:
                        load_expr.append(setup_h * Y[(p, loc, t)])

                # Changeover-straff: hver aktiv familie trekker 0.5 × snittstraff
                # (i timer) fra kapasiteten. YF[familie, wc, t] er binær.
                for f in model_families:
                    key = (f, wc_code, t)
                    if key in YF:
                        penalty_min = self.family_penalty.get((wc_code, f), 0.0)
                        if penalty_min > 0:
                            load_expr.append(
                                YF[key] * (0.5 * penalty_min / 60.0)
                            )

                if load_expr:
                    model += pulp.lpSum(load_expr) <= cap, f"Kapasitet_{wc_code}_{t}"

        # ── Materialbalanse med lager (BOM er pre-prosessert!) ──
        # For hvert produkt, lokasjon og periode:
        #   Inngående lager + Produksjon/kjøp + Transport inn
        #     = Etterspørsel + Transport ut + Utgående lager
        # Dette gjør at modellen kan produsere i én periode og dekke
        # etterspørsel i senere perioder via lager — mot en lagerholdskost.
        for (p, dloc, t), qty in expanded_demand.items():
            prev_inv = I[(p, dloc, t - 1)] if (p, dloc, t - 1) in I else 0.0
            prod = X[(p, dloc, t)] if (p, dloc, t) in X else 0.0
            trans_in = []
            for pl in prod_locations[p]:
                if pl != dloc and (p, pl, dloc, t) in T:
                    trans_in.append(T[(p, pl, dloc, t)])
            model += (
                prev_inv + prod + (pulp.lpSum(trans_in) if trans_in else 0.0)
                == qty + I[(p, dloc, t)],
                f"MatBalanse_{_safe(p)}_{dloc}_{t}",
            )

        # ── Transport/produksjon-kobling ─────────────────────────
        # Det som transporteres ut fra en lokasjon MÅ være produsert/kjøpt
        # på den lokasjonen. Uten denne koblingen kan modellen transportere
        # gratis mengde fra ingensteds (fysisk umulig).
        for (p, frm, to, t) in T_keys:
            if (p, frm, t) in X:
                model += T[(p, frm, to, t)] <= X[(p, frm, t)], (
                    f"TransKoble_{_safe(p)}_{frm}_{to}_{t}"
                )

        # Lagre intern referanse for resultatekstraksjon
        self._model_refs = {
            "X": X,
            "Y": Y,
            "I": I,
            "T": T,
            "YF": YF,
            "periods": periods,
            "locations": locations,
            "model_products": model_products,
        }
        return model

    # ── Løsing ────────────────────────────────────────────────────

    def solve(self, time_limit_seconds: Optional[int] = None) -> OptimizationResult:
        """Bygg modellen, løs den, og returner resultatstruktur."""
        model = self.build_model()

        import time
        start = time.time()

        if time_limit_seconds:
            model.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=time_limit_seconds))
        else:
            model.solve(pulp.PULP_CBC_CMD(msg=0))

        solve_time = time.time() - start

        refs = self._model_refs
        X = refs["X"]
        Y = refs["Y"]
        I = refs["I"]
        T = refs["T"]
        periods = refs["periods"]
        locations = refs["locations"]
        model_products = refs["model_products"]

        status = pulp.LpStatus[model.status]

        result = OptimizationResult(
            status=status,
            total_cost=pulp.value(model.objective) or 0.0,
            solve_time_seconds=solve_time,
        )

        if status != "Optimal":
            # Returner delvis/feil-status uten plan
            return result

        # Produksjonsplan
        for (p, l, t), var in sorted(X.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2])):
            val = var.value() or 0.0
            if val > 0.001:
                cost_per_unit = self.unit_costs.get((p, l), self.purchase_costs.get(p, 0.0))
                fc = self.full_costs.get((p, l), {})
                row = {
                    "product": p,
                    "product_desc": self._desc(p),
                    "location": l,
                    "period": t,
                    "quantity": round(val, 4),
                    "cost_per_unit": round(cost_per_unit, 4),
                    "total_cost": round(val * cost_per_unit, 4),
                }
                # Referanse: statisk fullkost fra CostCalculator (kostberegning.py)
                if fc:
                    row.update({
                        "full_operations_cost": round(fc.get("operations", 0.0), 4),
                        "full_material_cost": round(fc.get("material", 0.0), 4),
                        "full_setup_cost": round(fc.get("setup", 0.0), 4),
                        "full_byproduct_value": round(fc.get("byproduct", 0.0), 4),
                        "full_gross_cost": round(fc.get("gross", 0.0), 4),
                        "full_net_cost": round(fc.get("net", 0.0), 4),
                    })
                result.production_plan.append(row)

        # Batch-start (rigg) beslutninger
        for (p, l, t), var in sorted(Y.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2])):
            val = var.value() or 0.0
            if val > 0.5:
                result.batch_decisions.append({
                    "product": p,
                    "product_desc": self._desc(p),
                    "location": l,
                    "period": t,
                    "setup_cost": round(self.fixed_setup.get((p, l), 0.0), 4),
                })

        # Transportplan
        for (p, frm, to, t), var in sorted(T.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2], kv[0][3])):
            val = var.value() or 0.0
            if val > 0.001:
                tc = self.route_costs.get((p, frm, to), 0.0)
                result.transport_plan.append({
                    "product": p,
                    "from": frm,
                    "to": to,
                    "period": t,
                    "quantity": round(val, 4),
                    "cost_per_unit": round(tc, 4),
                    "total_cost": round(val * tc, 4),
                })

        # Kostnadsnedbrytning
        prod_cost = sum(
            X[k].value() * self.unit_costs.get((k[0], k[1]), self.purchase_costs.get(k[0], 0.0))
            for k in X if X[k].value() is not None
        )
        setup_cost = sum(
            Y[k].value() * self.fixed_setup.get((k[0], k[1]), 0.0)
            for k in Y if Y[k].value() is not None
        )
        trans_cost = sum(
            T[k].value() * self.route_costs.get((k[0], k[1], k[2]), 0.0)
            for k in T if T[k].value() is not None
        )
        # Lagernivåer
        for (p, l, t), var in sorted(I.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2])):
            val = var.value() or 0.0
            if val > 0.001:
                result.inventory_levels.append({
                    "product": p,
                    "product_desc": self._desc(p),
                    "location": l,
                    "period": t,
                    "quantity": round(val, 4),
                })

        hold_cost = 0.0
        if self.holding_cost_pct > 0:
            hold_cost = sum(
                I[k].value() * self.unit_costs.get((k[0], k[1]), self.purchase_costs.get(k[0], 0.0))
                * (self.holding_cost_pct / 100.0)
                for k in I if I[k].value() is not None
            )

        result.cost_breakdown = {
            "produksjon": prod_cost,
            "setup": setup_cost,
            "transport": trans_cost,
            "lagerhold": hold_cost,
        }

        return result

    def _desc(self, item_no: str) -> str:
        prod = self.data.product(item_no)
        return prod.description if prod else ""

    # ── Eksport ───────────────────────────────────────────────────

    def export_json(self, result: OptimizationResult, output_path: str):
        """Skriv resultat til JSON-fil."""
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result.to_dict(), f, ensure_ascii=False, indent=2)
        print(f"[OK] Resultat skrevet til {output_path}")


# ──────────────────────────────────────────────────────────────────────
#  Hjelpefunksjoner
# ──────────────────────────────────────────────────────────────────────

def _safe(name: str) -> str:
    """Gjør et produktnavn trygt for PuLP-variabelnavn (kun bokstaver/underscore)."""
    return "".join(ch if ch.isalnum() or ch == "_" else "_" for ch in name)


# FTI-prefikser (fra CLINE.md / FTI_Nummer.md) — 2-3 bokstaver som identifiserer
# produkttype/profilserie. Sortert etter lengde (3 før 2) for greedy matching.
_FTI_PREFIXES = [
    "AGH", "AGF", "AVF", "BL", "CC", "CD", "EGH", "ES", "HA", "IG", "IK",
    "IL", "IV", "JB", "JC", "JD", "JE", "JG", "JJ", "JK", "JL", "JM", "JO",
    "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JX", "JY", "LE", "ME", "PG",
    "SA", "YA",
]


def _fti_family(item_no: str) -> str:
    """Finn produksjonsfamilie for et FTI-nummer.

    Familie = prefiks + siffer (uten suffiks). Eksempel:
        'JD19073GH'  → 'JD19073'  (samme familie som 'JD19073' og 'JD19073TF')
        'RM_GRAN'    → 'RM_GRAN'  (ikke-FTI: fallback til hele item_no)

    Dette brukes for changeover-modellering: produkter i samme familie har
    ~0 min omstilling (kun malingsskift), mens ulike familier krever
    verktøybytte og lengre omstillingstid.
    """
    item = str(item_no).strip()
    upper = item.upper()

    # Forsøk å matche prefiks (lengste først)
    best_prefix = None
    for pref in sorted(_FTI_PREFIXES, key=len, reverse=True):
        if upper.startswith(pref):
            best_prefix = item[:len(pref)]
            break

    if best_prefix is None:
        # Ikke-FTI-produkt (f.eks. RM_GRAN, KD_TOPP) → hele varenr er familien
        return item

    rest = item[len(best_prefix):]

    # Siffer = påfølgende siffer (4-5 tegn) etter prefiks
    digits = ""
    for ch in rest:
        if ch.isdigit():
            digits += ch
        else:
            break
    if len(digits) >= 4:
        return f"{best_prefix}{digits}"

    # Ingen siffer funnet → hele varenr er familien
    return item


def _changeover_minutes(family_a: str, family_b: str,
                        family_dims: Optional[dict[str, tuple]] = None) -> float:
    """Estimert omstillingstid mellom to produktfamilier (minutter).

    Basert på FTI-nummerstrukturen (se CLINE.md — Changeover-modellering):

        Nivå 1: Suffiks-bytte (samme familie)       → 0 min
        Nivå 2: Bredde-bytte (samme tykkelse)       → 15-30 min (snitt 22.5)
        Nivå 3: Tykkelse-bytte                      → 45-60 min (snitt 52.5)
        Nivå 4: Prefiks-bytte (annen profilserie)   → ~90 min

    Args:
        family_a: Første familie (f.eks. 'JD19073')
        family_b: Andre familie (f.eks. 'JD19148')
        family_dims: Valgfri mapping familie → (tykkelse, bredde) for mer presis
                     sammenligning. Hvis None, parses dimensjoner fra sifferet.

    Returns:
        Estimert omstillingstid i minutter.
    """
    if family_a == family_b:
        return 0.0

    def _dims(fam: str):
        if family_dims and fam in family_dims:
            return family_dims[fam]
        # Pars fra FTI-siffer: JD19073 → tykkelse=19, bredde=73
        digits = "".join(ch for ch in fam if ch.isdigit())
        if len(digits) >= 4:
            thickness = int(digits[:2])
            width = int(digits[-3:])
            return (thickness, width)
        return (None, None)

    ta, wa = _dims(family_a)
    tb, wb = _dims(family_b)

    # Prefiks (profilserie) — finn prefiks-delen
    def _prefix(fam: str) -> str:
        return "".join(ch for ch in fam if not ch.isdigit())

    pa, pb = _prefix(family_a), _prefix(family_b)

    if pa != pb:
        return 90.0  # Nivå 4: full omstilling (annen profilserie)

    if ta is not None and tb is not None and ta != tb:
        return 52.5  # Nivå 3: tykkelse-bytte (snitt av 45-60)

    if wa is not None and wb is not None and wa != wb:
        return 22.5  # Nivå 2: bredde-bytte (snitt av 15-30)

    return 15.0  # Usikker/annen forskjell → konservativt estimat


def _finn_lm_per_m3_robust(data, item_no: str, visited: Optional[set[str]] = None):
    """Finn LM/M3-konvertering ved å følge BOM-kjeden til råvare med base_uom=M3.

    Følger kun trelast-komponenter (LM-baserte), og hopper over
    maling/andre råvarer med andre enheter (f.eks. LTR).

    Dette er en LOKAL robust versjon av `_finn_lm_per_m3` i kostberegning.py.
    Eksisterende kode endres ALDRI.

    Args:
        data: SqliteData-objekt
        item_no: Produktnummer å starte fra
        visited: Sett med besøkte produkter (sirkelbeskyttelse)

    Returns:
        LM/M3 (float) eller None hvis ingen M3-råvare finnes i kjeden.
    """
    if visited is None:
        visited = set()
    if item_no in visited:
        return None
    visited.add(item_no)

    for bl in data.bom_for(item_no):
        comp = data.product(bl.component_item_no)
        if comp is None:
            continue

        # Direkte M3-råvare → dette er LM/M3-forholdet
        if comp.item_type == "Raw Material" and comp.base_uom == "M3":
            if bl.quantity_per and bl.quantity_per > 0:
                return bl.quantity_per
            continue

        # Følg videre i LM-kjeden (delprodukter/ferdigvarer med LM)
        if bl.uom == "LM" or (
            comp.base_uom == "LM"
            and comp.item_type in ("Semi Finished", "Finished Good")
        ):
            result = _finn_lm_per_m3_robust(data, bl.component_item_no, visited)
            if result:
                return result

    return None


# ──────────────────────────────────────────────────────────────────────
#  Syntetisk test-case (flernivå kledning: Skurlast → Ubehandlet → Grunnet → Malt)
# ──────────────────────────────────────────────────────────────────────

def _seed_test_db(db: DataRepo):
    """Fyll temp-databasen med et flernivå-produksjonscase.

    Kunde-etterspørsel: 5000 LM Malt Kledning, uke 34, levering til KV.
    Kjedens steg:
      RM_GRAN → (høvling) → KLEDNING_UBEHANDLET → (grunning) → KLEDNING_GRUNNET
               → (maling) → KLEDNING_MALT
    """

    # Lokasjoner
    db.upsert_locations([
        {"code": "KOD", "name": "Kodal Fabrikk", "location_type": "Factory"},
        {"code": "KV", "name": "Kvaas", "location_type": "Factory"},
    ], source="seed")

    # Produkter
    db.upsert_products([
        {"item_no": "RM_GRAN", "description": "Skurlast Gran", "item_type": "Raw Material",
         "product_group": "Skurlast", "base_uom": "M3"},
        {"item_no": "KD_GRUNNING", "description": "Grunning", "item_type": "Raw Material",
         "product_group": "Maling", "base_uom": "LTR"},
        {"item_no": "KD_TOPP", "description": "Toppmaling", "item_type": "Raw Material",
         "product_group": "Maling", "base_uom": "LTR"},
        {"item_no": "KLEDNING_UBEHANDLET", "description": "Ubehandlet kledning",
         "item_type": "Semi Finished", "product_group": "Kledning", "base_uom": "LM"},
        {"item_no": "KLEDNING_GRUNNET", "description": "Grunnet kledning",
         "item_type": "Semi Finished", "product_group": "Kledning", "base_uom": "LM"},
        {"item_no": "KLEDNING_MALT", "description": "Malt kledning",
         "item_type": "Finished Good", "product_group": "Kledning", "base_uom": "LM"},
    ], source="seed")

    # Kostpriser (råvarer)
    db.upsert_item_costs([
        {"item_no": "RM_GRAN", "cost_type": "Standard Cost", "unit_cost": 3500.0,
         "currency": "NOK", "effective_date": None},
        {"item_no": "KD_GRUNNING", "cost_type": "Standard Cost", "unit_cost": 120.0,
         "currency": "NOK", "effective_date": None},
        {"item_no": "KD_TOPP", "cost_type": "Standard Cost", "unit_cost": 150.0,
         "currency": "NOK", "effective_date": None},
    ], source="seed")

    # Arbeidssentre
    db.upsert_work_centers([
        {"code": "HOVEDHOVEL", "description": "Hovedhovel Kodal", "location_code": "KOD",
         "labor_cost_hour": 800.0, "machine_cost_hour": 600.0, "overhead_cost_hour": 300.0,
         "capacity_hours_day": 16.0, "effective_capacity_pct": 90.0},
        {"code": "KVHOVEL", "description": "Hovel Kvaas", "location_code": "KV",
         "labor_cost_hour": 500.0, "machine_cost_hour": 400.0, "overhead_cost_hour": 200.0,
         "capacity_hours_day": 16.0, "effective_capacity_pct": 85.0},
        {"code": "GRUNNELINJE", "description": "Grunningslinje Kodal", "location_code": "KOD",
         "labor_cost_hour": 700.0, "machine_cost_hour": 600.0, "overhead_cost_hour": 300.0,
         "capacity_hours_day": 16.0, "effective_capacity_pct": 90.0},
        {"code": "MALTELINJE", "description": "Malingslinje Kodal", "location_code": "KOD",
         "labor_cost_hour": 900.0, "machine_cost_hour": 1000.0, "overhead_cost_hour": 400.0,
         "capacity_hours_day": 16.0, "effective_capacity_pct": 95.0},
        {"code": "MALTELINJE_KV", "description": "Malingslinje Kvaas", "location_code": "KV",
         "labor_cost_hour": 800.0, "machine_cost_hour": 900.0, "overhead_cost_hour": 350.0,
         "capacity_hours_day": 16.0, "effective_capacity_pct": 95.0},
    ], source="seed")

    # Operasjoner
    db.upsert_operations([
        {"code": "HOVLING", "description": "Høvling", "default_work_center": "HOVEDHOVEL",
         "standard_unit": "Minutes"},
        {"code": "GRUNNING", "description": "Grunning", "default_work_center": "GRUNNELINJE",
         "standard_unit": "Minutes"},
        {"code": "MALING", "description": "Maling", "default_work_center": "MALTELINJE",
         "standard_unit": "Minutes"},
    ], source="seed")

    # BOM (flernivå)
    db.upsert_bom_lines([
        {"parent_item_no": "KLEDNING_UBEHANDLET", "component_item_no": "RM_GRAN",
         "quantity_per": 400.0, "uom": "LM", "scrap_pct": 2.0, "co_product_pct": 0.0,
         "co_product_item_no": ""},
        {"parent_item_no": "KLEDNING_GRUNNET", "component_item_no": "KLEDNING_UBEHANDLET",
         "quantity_per": 1.0, "uom": "LM", "scrap_pct": 1.0, "co_product_pct": 0.0,
         "co_product_item_no": ""},
        {"parent_item_no": "KLEDNING_GRUNNET", "component_item_no": "KD_GRUNNING",
         "quantity_per": 25.0, "uom": "LTR", "scrap_pct": 5.0, "co_product_pct": 0.0,
         "co_product_item_no": ""},
        {"parent_item_no": "KLEDNING_MALT", "component_item_no": "KLEDNING_GRUNNET",
         "quantity_per": 1.0, "uom": "LM", "scrap_pct": 0.0, "co_product_pct": 0.0,
         "co_product_item_no": ""},
        {"parent_item_no": "KLEDNING_MALT", "component_item_no": "KD_TOPP",
         "quantity_per": 20.0, "uom": "LTR", "scrap_pct": 5.0, "co_product_pct": 0.0,
         "co_product_item_no": ""},
    ], source="seed")

    # Routing (produserbare produkter på flere lokasjoner)
    db.upsert_routing_lines([
        {"item_no": "KLEDNING_UBEHANDLET", "operation_no": 40, "operation_code": "HOVLING",
         "work_center_code": "HOVEDHOVEL", "setup_time_minutes": 30.0,
         "run_time_minutes": 0.010, "batch_size": 5000.0},
        {"item_no": "KLEDNING_UBEHANDLET", "operation_no": 40, "operation_code": "HOVLING",
         "work_center_code": "KVHOVEL", "setup_time_minutes": 30.0,
         "run_time_minutes": 0.012, "batch_size": 4000.0},
        {"item_no": "KLEDNING_GRUNNET", "operation_no": 40, "operation_code": "GRUNNING",
         "work_center_code": "GRUNNELINJE", "setup_time_minutes": 20.0,
         "run_time_minutes": 0.005, "batch_size": 3000.0},
        {"item_no": "KLEDNING_MALT", "operation_no": 40, "operation_code": "MALING",
         "work_center_code": "MALTELINJE", "setup_time_minutes": 30.0,
         "run_time_minutes": 0.008, "batch_size": 2000.0},
        {"item_no": "KLEDNING_MALT", "operation_no": 40, "operation_code": "MALING",
         "work_center_code": "MALTELINJE_KV", "setup_time_minutes": 25.0,
         "run_time_minutes": 0.010, "batch_size": 1800.0},
    ], source="seed")

    # Transportruter
    for frm, to, cost in [("KOD", "KV", 120.0), ("KV", "KOD", 130.0)]:
        db.conn.execute(
            """INSERT OR IGNORE INTO transport_ruter (from_loc, to_loc, cost_per_m3, distance_km, hours)
               VALUES (?, ?, ?, 35.0, 1.5)""",
            (frm, to, cost),
        )
    db.conn.commit()

    # Etterspørsel: 5000 LM Malt Kledning, uke 34, levering til KV
    db.conn.execute(
        """INSERT INTO demand (product_id, period, quantity, location_code, customer_region)
           VALUES (?, ?, ?, ?, ?)""",
        ("KLEDNING_MALT", 34, 5000.0, "KV", "Region Øst"),
    )
    db.conn.commit()


# ──────────────────────────────────────────────────────────────────────
#  Kommandolinje
# ──────────────────────────────────────────────────────────────────────

def _run_test(args) -> int:
    """Kjør mot et syntetisk flernivå-case i en temp-database."""
    tmp_dir = tempfile.mkdtemp(prefix="pk_opt_")
    db_path = os.path.join(tmp_dir, "test.db")

    db = DataRepo(db_path)
    db.initialize()
    _seed_test_db(db)
    db.close()

    print("=" * 70)
    print(" TEST-CASE: Flernivå kledning (Skurlast → Ubehandlet → Grunnet → Malt)")
    print(" Etterspørsel: 5000 LM Malt Kledning, uke 34, levert til KV")
    print("=" * 70)

    engine = OptimizationEngine(
        db_path=db_path,
        days_per_period=args.days,
        holding_cost_pct=args.holding_cost_pct,
        include_transport=not args.no_transport,
        forecast_weeks=args.forecast_weeks,
        forecast_max_setup_pct=args.forecast_max_setup_pct,
    )
    engine.load_all()

    print(f"[PROGNOSE] {args.forecast_weeks} ukers prognose fra historisk salg aktivert.")
    print(f"[PROGNOSE] Sett opp: max {args.forecast_max_setup_pct}% setup-andel, {len(engine._forecast_records())} prognose-linjer.")
    print(f"{'=' * 70}")

    # Vis bom_structure som ble auto-populert
    print("\n[FLERNIVÅ-BOM] Auto-populert fra bom_lines:")
    for bs in engine.bom_structure:
        print(f"   {bs.parent_product_id} → {bs.child_product_id}  (yield={bs.yield_factor})")

    print(f"\n[ENHETSKOST] Beregnet av CostCalculator:")
    for (p, l), cost in sorted(engine.unit_costs.items()):
        print(f"   {p:<28} {l:<5} operasjonskost/enne={cost:>10.4f}  setup/batch={engine.fixed_setup.get((p,l),0):>10.2f}")

    result = engine.solve(time_limit_seconds=args.time_limit)

    print(f"\n[STATUS] {result.status}  |  Total kost: {result.total_cost:,.2f} kr  "
          f"|  Løsningstid: {result.solve_time_seconds:.2f}s")

    if result.cost_breakdown:
        print("[KOSTNADSNEDBRYTNING]")
        for k, v in result.cost_breakdown.items():
            print(f"   {k:<12} {v:>14,.2f} kr")

    if result.production_plan:
        print("\n[PRODUKSJONSPLAN]")
        for row in result.production_plan:
            print(f"   {row['product']:<28} {row['location']:<5} uke {row['period']:<4} "
                  f"{row['quantity']:>12,.1f}  ({row['total_cost']:,.2f} kr)")

    if result.batch_decisions:
        print("\n[BATCH-START / RIGG]")
        for row in result.batch_decisions:
            print(f"   {row['product']:<28} {row['location']:<5} uke {row['period']:<4} "
                  f"setup={row['setup_cost']:,.2f} kr")

    if result.transport_plan:
        print("\n[TRANSPORTPLAN]")
        for row in result.transport_plan:
            print(f"   {row['product']:<28} {row['from']} → {row['to']:<5} "
                  f"uke {row['period']:<4} {row['quantity']:>12,.1f}  ({row['total_cost']:,.2f} kr)")

    if args.output:
        engine.export_json(result, args.output)

    return 0


def _run_real(args) -> int:
    """Kjør mot den ekte databasen — med strukturert, lesbar output."""
    db_path = args.db or str(_src / "produksjonskalkyle.db")

    print(f"[OPTIMERING] Database: {db_path}")
    print(f"[OPTIMERING] Dager per periode: {args.days}")

    engine = OptimizationEngine(
        db_path=db_path,
        days_per_period=args.days,
        holding_cost_pct=args.holding_cost_pct,
        include_transport=not args.no_transport,
        forecast_weeks=args.forecast_weeks,
        forecast_max_setup_pct=args.forecast_max_setup_pct,
    )
    engine.load_all()

    print(f"[DATA] Demand-linjer: {len(engine.demand)}")
    print(f"[DATA] Transportruter: {len(engine.transport_ruter)}")
    print(f"[PROGNOSE] {args.forecast_weeks} ukers prognose, max {args.forecast_max_setup_pct}% setup-andel ({len(engine._forecast_records())} prognose-linjer).")

    if not engine.demand:
        print("[FEIL] Ingen demand-data funnet. Fyll demand-tabellen først.")
        print("  Eksempel:")
        print("    INSERT INTO demand (product_id, period, quantity, location_code)")
        print("    VALUES ('KLEDNING_MALT_19X148', 34, 5000, 'KV');")
        return 1

    result = engine.solve(time_limit_seconds=args.time_limit)

    # ── Overskrift / STATUS ─────────────────────────────────────
    print("\n" + "=" * 74)
    print(f"  OPTIMERINGSRESULTAT  |  {result.status}")
    print("=" * 74)
    print(f"  Total kost:            {result.total_cost:>14,.2f} kr")
    print(f"  Løsningstid:           {result.solve_time_seconds:>14.2f} s")

    if result.cost_breakdown:
        print(f"  ├─ Produksjon:         {result.cost_breakdown.get('produksjon', 0):>14,.2f} kr")
        print(f"  ├─ Setup (rigg):       {result.cost_breakdown.get('setup', 0):>14,.2f} kr")
        print(f"  └─ Transport:          {result.cost_breakdown.get('transport', 0):>14,.2f} kr")
    print("=" * 74)

    if result.status != "Optimal":
        print("\n  ⚠️  Modellen fant ingen gyldig løsning innenfor kapasiteten.")
        print("      Dette betyr at demandet overstiger tilgjengelig maskintid.")
        print("      Prøv --days 10 for å se hva det koster å produsere alt.")
        return 0

    # ── Bygg oppslag: (produkt, lokasjon) → arbeidssenter ──────
    item_wc: dict[tuple[str, str], str] = {}
    for rl in engine.data.routing_lines:
        wc = engine.data.work_center(rl.work_center_code)
        if wc:
            # Første (laveste op_no) arbeidssenter vinner
            key = (rl.item_no, wc.location_code)
            if key not in item_wc:
                item_wc[key] = wc.code

    # ── Kapasitet per arbeidssenter per uke ────────────────────
    periods = sorted({r["period"] for r in result.production_plan})
    wc_capacity: dict[tuple[str, int], float] = {}
    for wc in engine.data.work_centers:
        eff = (wc.effective_capacity_pct or 100) / 100
        for t in periods:
            wc_capacity[(wc.code, t)] = wc.capacity_hours_day * eff * args.days

    # ── Grupper produksjon: (periode, arbeidssenter, produkt) ──
    # produkt_type: demand-produkt (sluttprodukt), halvfabrikat, råvare
    demand_items = {d.product_id for d in engine.demand}
    routing_items = {rl.item_no for rl in engine.data.routing_lines}
    purchasable = {ic.item_no for ic in engine.data.item_costs}

    grouped: dict[tuple[int, str], list[dict]] = {}
    for row in result.production_plan:
        p, loc, t = row["product"], row["location"], row["period"]
        wc = item_wc.get((p, loc), "KJØP/UKJENT")
        if p in demand_items:
            typ = "Ferdigvare"
        elif p in routing_items:
            typ = "Halvfabrikat"
        else:
            typ = "Råvare/kjøp"
        grouped.setdefault((t, wc), []).append({**row, "type": typ,
                                                "is_demand": p in demand_items})

    # ── Skriv strukturert: per uke → arbeidssenter ────────────
    for t in periods:
        print(f"\n" + "-" * 74)
        print(f"  UKE {t}")
        print("-" * 74)

        for wc in sorted({wc for (pt, wc) in grouped if pt == t}):
            rows = grouped[(t, wc)]
            cap = wc_capacity.get((wc, t), 0)
            # Beregn brukt tid: sum run_hours × antall (approx fra unit_costs/route)
            # Forenklet: bruk produksjonstid fra routing
            used_hours = 0.0
            for rr in rows:
                p, loc = rr["product"], rr["location"]
                total_run = 0.0
                for rl in engine.data.routing_lines:
                    if rl.item_no == p and rl.work_center_code == wc:
                        total_run += rl.run_time_minutes / 60.0
                # Co-prod-justering
                for bl in engine.data.bom_for(p):
                    if bl.co_product_pct > 0:
                        total_run *= 1 + bl.co_product_pct / 100
                used_hours += total_run * rr["quantity"]
                # Setup-tid per batch (ca)
                setup_total = 0.0
                for rl in engine.data.routing_lines:
                    if rl.item_no == p and rl.work_center_code == wc:
                        setup_total += rl.setup_time_minutes / 60.0
                used_hours += setup_total  # ca. én batch

            pct = (used_hours / cap * 100) if cap > 0 else 0
            print(f"\n  [{wc}]  Kapasitet: {used_hours:>6.1f}t / {cap:>6.1f}t ({pct:>5.1f}%)")

            # FTI-sekvensering: tykkelse → bredde → prefiks → suffiks
            def _fti_sort_key(product_id):
                fam = _fti_family(product_id)
                d = "".join(ch for ch in fam if ch.isdigit())
                if len(d) >= 4:
                    tykkelse, bredde = int(d[:2]), int(d[-3:])
                else:
                    tykkelse, bredde = 99, 9999
                pref = "".join(ch for ch in fam if not ch.isdigit())
                suff = product_id[len(fam):] if product_id.startswith(fam) else product_id
                return (tykkelse, bredde, pref, suff)

            # Ferdigvarer først, deretter halvfabrikata, så råvarer —
            # innenfor hver type sortert etter FTI-struktur
            order = {"Ferdigvare": 0, "Halvfabrikat": 1, "Råvare/kjøp": 2}
            rows_sorted = sorted(
                rows,
                key=lambda r: (order.get(r["type"], 3), _fti_sort_key(r["product"])),
            )
            prev_fam = None
            for i, rr in enumerate(rows_sorted, 1):
                fam = _fti_family(rr["product"])
                # Kun vis familiebryter innenfor samme produkttype
                if prev_fam and fam != prev_fam and rr["type"] == rows_sorted[i - 2]["type"]:
                    co = _changeover_minutes(prev_fam, fam, engine.family_dims)
                    print(f"      ── changeover {prev_fam} → {fam}: {co:.1f} min")
                marker = "►" if rr["is_demand"] else " "
                print(f"   {i:<3} {marker} {rr['product']:<26} {rr['type']:<12} "
                      f"{rr['quantity']:>10,.1f} LM")
                prev_fam = fam
        print("-" * 74)

    # ── Transportplan ───────────────────────────────────────────
    if result.transport_plan:
        print(f"\n{'=' * 74}")
        print("  TRANSPORTPLAN")
        print("=" * 74)
        for t in periods:
            trows = [r for r in result.transport_plan if r["period"] == t]
            if not trows:
                continue
            print(f"\n  UKE {t}:")
            for rr in sorted(trows, key=lambda r: (r["from"], r["to"], r["product"])):
                print(f"   {rr['product']:<28} {rr['from']} → {rr['to']:<5} "
                      f"{rr['quantity']:>10,.1f} LM  ({rr['total_cost']:,.2f} kr)")

    # ── Setup/batch-plan ────────────────────────────────────────
    if result.batch_decisions:
        print(f"\n{'=' * 74}")
        print("  BATCH-START / RIGG")
        print("=" * 74)
        for t in periods:
            brows = [r for r in result.batch_decisions if r["period"] == t]
            if not brows:
                continue
            print(f"\n  UKE {t}:")
            for rr in sorted(brows, key=lambda r: (r["location"], r["product"])):
                print(f"   {rr['product']:<28} {rr['location']:<5}  "
                      f"setup={rr['setup_cost']:,.2f} kr")

    print(f"\n{'=' * 74}")
    print("  FERDIG — Bytt til --days 10 eller høyere for å se en løsbar plan")
    print("=" * 74)

    if args.output:
        engine.export_json(result, args.output)

    return 0


def main():
    parser = argparse.ArgumentParser(
        description="MILP-produksjonsoptimering med flernivå-BOM (ProduksjonsKalkyle)"
    )
    parser.add_argument("--test", action="store_true",
                        help="Kjør syntetisk flernivå-test-case (temp-database)")
    parser.add_argument("--db", type=str, default=None,
                        help="Sti til SQLite-database (default: src/produksjonskalkyle.db)")
    parser.add_argument("--output", type=str, default=None,
                        help="Skriv resultat til JSON-fil")
    parser.add_argument("--days", type=float, default=5.0,
                        help="Antall arbeidsdager per periode (default: 5)")
    parser.add_argument("--holding-cost-pct", type=float, default=0.0,
                        help="Lagerholdskost i % av enhetskost per periode (default: 0)")
    parser.add_argument("--forecast-weeks", type=int, default=12,
                        help="Antall uker fremover med prognose fra historisk salg (default: 12)")
    parser.add_argument("--forecast-max-setup-pct", type=float, default=30.0,
                        help="Maks tillatt setup-kost som %% av operasjonskost for en prognoseperiode "
                             "(default: 30). 0 = ubegrenset. Lavere verdi gir færre, større batcher.")
    parser.add_argument("--time-limit", type=int, default=None,
                        help="Maks løsningstid i sekunder for solverb")
    parser.add_argument("--no-transport", action="store_true",
                        help="Utelat transportvariabler fra modellen")
    args = parser.parse_args()

    # Kjør mot temp test-db eller ekte db
    if args.test:
        return _run_test(args)
    return _run_real(args)


if __name__ == "__main__":
    sys.exit(main())