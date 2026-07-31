# Verdikjede- og Kapasitetsmodul — Fullstendig implementasjonsplan

> **Mål:** Utvide ProduksjonsKalkyle fra en kostpris-simulator til en komplett verdikjede-plattform
> som også håndterer transport, flerstegs produksjon på tvers av lokasjoner, råstoffsporing,
> og kapasitetsplanlegging — uten å endre på den eksisterende kostpris-beregningen.
>
> **Hovedprinsipp:** `kostberegning.py` og `varekost_app.py` forblir **helt urørt**.
> All ny funksjonalitet bygges som separate moduler og separate Marimo-apper,
> som kun leser resultater fra kostberegningsmotoren (aldri endrer den).

---

## Innholdsfortegnelse

1. [Overordnet arkitektur](#1-overordnet-arkitektur)
2. [Tre-app strategi](#2-tre-app-strategi)
3. [Datamodell-utvidelser](#3-datamodell-utvidelser)
4. [Nye Python-moduler](#4-nye-python-moduler)
   - 4A. `transport.py`
   - 4B. `verdikjede.py`
   - 4C. `materialflyt.py`
5. [Verdikjede-app: `verdikjede_app.py`](#5-verdikjede-app)
6. [Kapasitet-app: `kapasitet_app.py`](#6-kapasitet-app)
7. [Konkret scenario — Malt panel via sagbruk, KV høvel og KOD maling](#7-konkret-scenario)
8. [Trinnvis implementasjonsplan](#8-trinnvis-implementasjonsplan)
9. [Tekniske garantier (hva røres IKKE)](#9-tekniske-garantier)
10. [Server-oppsett](#10-server-oppsett)
11. [Brukerhistorier — hvilke spørsmål hver app svarer på](#11-brukerhistorier)
12. [Brukerdokumentasjon](#12-brukerdokumentasjon)
13. [Teknisk dokumentasjon](#13-teknisk-dokumentasjon)

---

## 1. Overordnet arkitektur

```
                     ┌──────────────────────────────────────┐
                     │      SQLite (endringslogg.db)         │
                     │    Felles database — ÉN sannhet       │
                     │    + transport_routes, value_chains   │
                     └───────────┬──────────────┬────────────┘
                                 │              │
           ┌─────────────────────┼──────────────┼─────────────────────┐
           │                     │              │                     │
           ▼                     ▼              ▼                     ▼
┌──────────────────┐  ┌────────────────────┐  ┌────────────────────┐
│ varekost_app.py  │  │ verdikjede_app.py  │  │ kapasitet_app.py   │
│ (KOSTPRIS)       │  │ (VERDIKJEDE)       │  │ (KAPASITET)        │
│                  │  │                    │  │                    │
│ Port 8080        │  │ Port 8081          │  │ Port 8082          │
│                  │  │                    │  │                    │
│ Leser: SQLite   │  │ Leser: SQLite      │  │ Leser: SQLite      │
│ Bruker:         │  │ Bruker:            │  │ Bruker:            │
│ CostCalculator  │  │ CostCalculator     │  │ CostCalculator     │
│ SimulationEng.  │  │ TransportCalc      │  │ CapacityPlanner    │
│                  │  │ ValueChainCalc     │  │ BatchOptimizer     │
│                  │  │ MaterialTracker    │  │ LocationAllocator  │
│                  │  │                    │  │                    │
│ UENDRET          │  │ NY                 │  │ NY                 │
└──────────────────┘  └────────────────────┘  └────────────────────┘
         │                       │                       │
         └───────────────────────┼───────────────────────┘
                                 ▼
                    ┌──────────────────────┐
                    │  eksisterende filer   │
                    │  kostberegning.py     │ ← URØRT
                    │  data_repo.py         │ ← + nye tabeller
                    │  excel_bridge.py      │ ← + nye ark
                    │  generer_pdf_rapport  │ ← + verdikjede-rapport
                    │  generer_excel_rapport│ ← + verdikjede-eksport
                    └──────────────────────┘
```

### Viktige arkitekturregler

1. **`kostberegning.py` endres aldri.** Den inneholder all kjerneberegning. Nye funksjoner som `verdikjede.py` og `transport.py` importerer KUN klasser som `CostCalculator`, `ProductCostResult` etc. for å lese resultater — de endrer aldri noe i dem.
2. **`varekost_app.py` endres aldri.** Den forblir nøyaktig som i dag, og kjenner ikke til transport, verdikjede eller kapasitet.
3. **Ny funksjonalitet må lese samme SQLite-database.** Det er koblingen. `endringslogg.db` er den felles datakilden.
4. **Hver ny modul har én ansvarlighet.** `transport.py` håndterer kun transport. `verdikjede.py` håndterer kun verdikjeder. `materialflyt.py` håndterer kun råstoffsporing. `kapasitetsplanlegger.py` håndterer kun kapasitet.

---

## 2. Tre-app strategi

### Hvorfor tre apper, ikke én?

| Årsak | Forklaring |
|---|---|
| **Enkelhet for brukerne** | Kollegaer som kun trenger kostpris ser kun kostpris. De som trenger verdikjede åpner den separat. Ingen blir overveldet. |
| **Isolasjon** | En feil i verdikjede-appen kan ikke krasje kostpris-appen. De kjører som separate prosesser. |
| **Modulær utrulling** | Kapasitet-appen kan skrus på/av etter behov. Ingen trenger å forholde seg til den før den er klar. |
| **Uavhengig utvikling** | Hver app kan utvikles, testes og deployes uavhengig. |
| **Ytelse** | Hver app laster kun det den trenger. Kostpris-appen blir ikke tregere av verdikjede-data. |

### App-oversikt

| App | Port | Formål | Målgruppe | Status |
|---|---|---|---|---|
| `varekost_app.py` | 8080 | Standardkost-beregning, parameter-simulering, PDF/Excel-rapport | Økonomi, innkjøp, produksjon | **Ferdig (urørt)** |
| `verdikjede_app.py` | 8081 | Flerstegs verdikjede, transportkost, råstoffsporing, "hva-hvis" flytting | Logistikk, planleggere, ledelse | **Skal bygges** |
| `kapasitet_app.py` | 8082 | Kapasitetsplanlegging, flaskehalsanalyse, batch-optimalisering, allokering | Produksjonsplanleggere, driftsledere | **Skal bygges** |

---

## 3. Datamodell-utvidelser

### 3.1 Ny tabell: Transport Routes

**Excel-ark:** `Transport Routes`
**SQLite-tabell:** `transport_routes`
**Python-klasse:** `TransportRoute`

| Kolonne | Type | Påkrevd | Beskrivelse | Eksempel |
|---|---|---|---|---|
| **From Location** | Tekst | Ja | Avsender-lokasjon (ref. Locations) | `SAG` |
| **To Location** | Tekst | Ja | Mottaker-lokasjon (ref. Locations) | `KV` |
| **Transport Type** | Tekst | Ja | `Road`, `Internal`, `Rail`, `Sea` | `Road` |
| **Cost Per Unit** | Desimal | Ja | Kostnad per transport-enhet | `85.00` |
| **Cost Unit** | Tekst | Ja | Enhet for kostnadsberegning | `M3`, `LM`, `KG`, `PCS` |
| **Time Hours** | Desimal | Nei | Transporttid i timer (for logistikkplanlegging) | `1.5` |
| **Distance KM** | Desimal | Nei | Avstand i km (for kostnad per km) | `45.0` |
| **CO2 Per Unit** | Desimal | Nei | CO2-utslipp per enhet (fremtidig miljørapportering) | `2.3` |
| **Valid From** | Dato | Nei | Gyldig fra dato | `2026-01-01` |
| **Valid To** | Dato | Nei | Gyldig til dato | |
| **Active** | Ja/Nei | Nei | Aktiv/Inaktiv | `Ja` |

**Eksempeldata:**
| From | To | Type | Cost/Unit | Unit | Time (h) | KM | CO2/Unit |
|---|---|---|---|---|---|---|---|
| SAG | KOD | Road | 120.00 | M3 | 2.0 | 60 | 3.5 |
| SAG | KV | Road | 85.00 | M3 | 1.5 | 45 | 2.3 |
| KV | KOD | Road | 0.12 | LM | 1.0 | 30 | 0.004 |
| KOD | SKI | Road | 0.08 | LM | 0.5 | 15 | 0.002 |

### 3.2 Ny tabell: Value Chains

**Excel-ark:** `Value Chains`
**SQLite-tabell:** `value_chains`
**Python-klasse:** `ValueChainStep`

Hver rad er ett steg i en verdikjede. En verdikjede består av en sekvens steg (sortert på Step Order) som beskriver hele reisen til et produkt.

| Kolonne | Type | Påkrevd | Beskrivelse | Eksempel |
|---|---|---|---|---|
| **Chain ID** | Tekst | Ja | Unik ID for verdikjeden | `VC001` |
| **Chain Name** | Tekst | Ja | Menneskelesbart navn | `Panel via KV` |
| **Final Product** | Tekst | Ja | Sluttproduktet (ref. Product Master) | `FG004` |
| **Step Order** | Heltall | Ja | Rekkefølge (1, 2, 3...) | `1` |
| **Step Type** | Tekst | Ja | `Production` eller `Transport` | `Production` |
| **Location** | Tekst | Ja | Hvor steget utføres (ref. Locations) | `SAG` |
| **Work Center** | Tekst | Kun Production | Arbeidssenter som utfører steget | `SAGLINJE` |
| **Operation Code** | Tekst | Kun Production | Hvilken operasjon | `SAWING` |
| **Intermediate Product** | Tekst | Nei | Mellomprodukt etter steget (ref. Product Master) | `HF001` |
| **Quantity In** | Desimal | Ja | Antall input-enheter per sluttprodukt-enhet | `1.0` |
| **Quantity Out** | Desimal | Ja | Antall output-enheter per sluttprodukt-enhet | `0.45` |
| **UOM In** | Tekst | Ja | Enhet for input | `M3` |
| **UOM Out** | Tekst | Ja | Enhet for output | `M3` |

**Eksempeldata — VC001: Malt Panel via KV høvel:**
| Chain ID | Step | Type | Location | WC | Op | Intermed. Product | Qty In | Qty Out | UOM In | UOM Out |
|---|---|---|---|---|---|---|---|---|---|---|
| VC001 | 1 | Production | SAG | SAGLINJE | SAWING | HF001 | 1.0 | 0.45 | M3 | M3 |
| VC001 | 2 | Transport | SAG→KV | - | - | - | 0.45 | 0.45 | M3 | M3 |
| VC001 | 3 | Production | KV | KVHOVEL | PLANING | FG001 | 0.45 | 180.0 | M3 | LM |
| VC001 | 4 | Transport | KV→KOD | - | - | - | 180.0 | 180.0 | LM | LM |
| VC001 | 5 | Production | KOD | MALINGSLINJE | MALING | FG004 | 180.0 | 176.4 | LM | LM |
| VC001 | 6 | Production | KOD | PAKKELINJE | PACKING | FG004 | 176.4 | 176.4 | LM | LM |

### 3.3 Endringer i eksisterende datalag

**`data_repo.py`:** Legg til:
- `transport_routes` tabell (CREATE TABLE + CRUD)
- `value_chains` tabell (CREATE TABLE + CRUD)
- Indekser på `from_location`, `to_location`, `chain_id`, `final_product`

**`excel_bridge.py`:** Legg til:
- Validering av `Transport Routes`-arket (sjekk at From/To Location finnes i Locations)
- Validering av `Value Chains`-arket (sjekk at Location, Work Center, Operation, Product finnes)
- Import av nye ark til SQLite
- Eksport av nye ark fra SQLite

**`kostberegning.py` / `SqliteData`:** Legg til nye metoder:
```python
# NYE metoder — endrer IKKE eksisterende metoder
def _load_transport_routes(self):
    """Last transport_routes fra SQLite."""
    # SQL: SELECT * FROM transport_routes WHERE active = 1

def _load_value_chains(self):
    """Last value_chains fra SQLite."""
    # SQL: SELECT * FROM value_chains ORDER BY chain_id, step_order

def transport_route(self, from_loc: str, to_loc: str) -> Optional[TransportRoute]:
    """Slå opp transportrute."""
    
def value_chain(self, chain_id: str) -> list[ValueChainStep]:
    """Hent alle steg for en verdikjede."""
    
def value_chains_for_product(self, product_no: str) -> list[list[ValueChainStep]]:
    """Hent alle verdikjeder som ender med gitt produkt."""
```

Disse nye metodene legges til i BÅDE `ExcelData` og `SqliteData` for konsistens, med parsing fra hhv. Excel-ark og SQLite-tabeller.

---

## 4. Nye Python-moduler

### 4A. `transport.py`

**Formål:** Beregne transportkostnader mellom lokasjoner.

**Fil:** `transport.py` (~150 linjer)

**Struktur:**

```python
"""
transport.py — Transportkostnader mellom lokasjoner.

Avhengigheter:
    - data_repo.py (for å lese transport_routes fra SQLite)
    
Brukes av:
    - verdikjede.py (ValueChainCalculator importerer TransportCalculator)
    - verdikjede_app.py (for å justere transportkost og se påvirkning)
"""

from dataclasses import dataclass
from typing import Optional

@dataclass
class TransportRoute:
    from_location: str
    to_location: str
    transport_type: str
    cost_per_unit: float
    cost_unit: str
    time_hours: float = 0.0
    distance_km: float = 0.0
    co2_per_unit: float = 0.0
    valid_from: Optional[date] = None
    valid_to: Optional[date] = None
    active: bool = True
    
    @property
    def route_key(self) -> tuple[str, str]:
        return (self.from_location, self.to_location)


@dataclass
class TransportCostResult:
    """Resultat av en transportkost-beregning."""
    from_location: str
    to_location: str
    transport_type: str
    quantity: float
    unit: str
    cost_per_unit: float
    total_cost: float
    time_hours: float
    co2_amount: float


class TransportCalculator:
    """
    Beregn transportkost for gitte ruter og kvanta.
    
    Støtter:
    - Enkel rute: from → to
    - Flere ruter i sekvens (for flerstegs transport)
    - Kostnad per enhet (M3, LM, KG, PCS)
    """
    
    def __init__(self, routes: list[TransportRoute]):
        self._route_index: dict[tuple[str, str], TransportRoute] = {
            r.route_key: r for r in routes
        }
    
    def get_route(self, from_loc: str, to_loc: str) -> Optional[TransportRoute]:
        """Slå opp transportrute."""
        return self._route_index.get((from_loc, to_loc))
    
    def calculate_cost(self, from_loc: str, to_loc: str, quantity: float) -> Optional[TransportCostResult]:
        """
        Beregn transportkost for én rute.
        
        Formel:
            total_cost = cost_per_unit * quantity
            
        Hvis ruten ikke finnes, returneres None.
        """
        route = self.get_route(from_loc, to_loc)
        if not route:
            return None
        
        return TransportCostResult(
            from_location=from_loc,
            to_location=to_loc,
            transport_type=route.transport_type,
            quantity=quantity,
            unit=route.cost_unit,
            cost_per_unit=route.cost_per_unit,
            total_cost=route.cost_per_unit * quantity,
            time_hours=route.time_hours,
            co2_amount=route.co2_per_unit * quantity,
        )
    
    def calculate_chain_cost(self, route_chain: list[tuple[str, str]], 
                              quantities: list[float]) -> list[TransportCostResult]:
        """
        Beregn transportkost for en kjede av ruter.
        
        Eksempel:
            route_chain = [("SAG", "KV"), ("KV", "KOD")]
            quantities  = [0.45, 180.0]  # forskjellige enheter per steg
        """
        results = []
        for (fr, to), qty in zip(route_chain, quantities):
            result = self.calculate_cost(fr, to, qty)
            if result:
                results.append(result)
        return results
```

### 4B. `verdikjede.py`

**Formål:** Beregne totalkostnad for et produkt gjennom en flerstegs verdikjede med transport og produksjon på tvers av lokasjoner.

**Fil:** `verdikjede.py` (~400 linjer)

**Struktur:**

```python
"""
verdikjede.py — Flerstegs produksjonskjede med transport.

En "verdikjede" definerer hele reisen til et produkt:
    Tømmer → Sagbruk → Transport → Høvleri → Transport → Maling → Ferdigvare

ValueChainCalculator bygger på CostCalculator (import, ikke endring)
og legger til transport mellom hvert produksjonssteg.

Brukes av:
    - verdikjede_app.py (Marimo-app)
    - kapasitet_app.py (for kapasitetskonsekvenser av flytting)
"""

from dataclasses import dataclass, field
from typing import Optional
from kostberegning import SqliteData, CostCalculator, ProductCostResult
from transport import TransportCalculator, TransportCostResult, TransportRoute

@dataclass
class ValueChainStep:
    """Ett steg i en verdikjede."""
    step_order: int
    step_type: str  # "Production" eller "Transport"
    location: str
    location_name: str = ""
    work_center: str = ""
    operation_code: str = ""
    intermediate_product: str = ""
    intermediate_desc: str = ""
    quantity_in: float = 1.0
    quantity_out: float = 1.0
    uom_in: str = ""
    uom_out: str = ""

@dataclass
class StepCostDetail:
    """Kostnadsdetaljer for ett steg i verdikjeden."""
    step_order: int
    step_type: str
    location: str
    work_center: str = ""
    description: str = ""  # f.eks. "Saging" eller "Transport SAG→KV"
    
    # For Production-steg
    material_cost: float = 0.0
    operation_cost: float = 0.0
    setup_cost: float = 0.0
    byproduct_value: float = 0.0
    net_cost: float = 0.0
    
    # For Transport-steg
    transport_cost: float = 0.0
    transport_time: float = 0.0
    
    # Total for dette steget
    step_cost: float = 0.0        # Kost for dette steget alene
    accumulated_cost: float = 0.0 # Akkumulert kost for produktet så langt (per enhet sluttprodukt)

@dataclass
class ValueChainResult:
    """Fullt resultat av verdikjede-beregning."""
    chain_id: str
    chain_name: str
    final_product: str
    product_desc: str
    base_uom: str
    total_production_cost: float = 0.0    # Sum av alle produksjons-steg
    total_transport_cost: float = 0.0     # Sum av alle transport-steg
    total_chain_cost: float = 0.0         # Total = produksjon + transport
    cost_per_unit_output: float = 0.0     # per enhet sluttprodukt
    steps: list[StepCostDetail] = field(default_factory=list)
    
    # Råstoffsporing
    raw_material_cost: float = 0.0        # Kost for første råmateriale i kjeden
    value_added: float = 0.0              # Forbedling = total - råmaterial


class ValueChainCalculator:
    """
    Beregn totalkost for et produkt gjennom en verdikjede.
    
    Slik fungerer det:
    1. Hent verdikjede-steg fra SQLite/Excel
    2. For hvert Production-steg:
       a. Finn produktet som produseres (intermediate_product)
       b. Finn alle BOM-linjer og routing-linjer for dette produktet
       c. Beregn kost med CostCalculator (men kun på den angitte lokasjonen)
       d. Juster for quantity (quantity_out / quantity_in) for å skalere til sluttprodukt-enhet
    3. For hvert Transport-steg:
       a. Slå opp transportrute (from→to)
       b. Beregn kost basert på kvantum
    4. Akkumuler kost per steg
    5. Returner resultat
    """
    
    def __init__(self, data: SqliteData):
        self.data = data
        self.cost_calc = CostCalculator(data)
        
        # Bygg transportkalkulator fra data
        _routes = []
        if hasattr(data, 'transport_routes') and data.transport_routes:
            for r in data.transport_routes:
                _routes.append(TransportRoute(
                    from_location=r.from_location,
                    to_location=r.to_location,
                    transport_type=r.transport_type,
                    cost_per_unit=r.cost_per_unit,
                    cost_unit=r.cost_unit,
                    time_hours=r.time_hours,
                    distance_km=r.distance_km,
                    co2_per_unit=r.co2_per_unit,
                ))
        self.transport_calc = TransportCalculator(_routes)
    
    def calculate_chain(self, chain_id: str, 
                         overrides: Optional[SimulationOverride] = None) -> Optional[ValueChainResult]:
        """
        Beregn totalkost for en verdikjede.
        
        Args:
            chain_id: ID for verdikjeden (f.eks. "VC001")
            overrides: Valgfrie overstyringsparametere (samme som SimulationEngine)
            
        Returns:
            ValueChainResult med steg-for-steg kostnader, eller None ved feil
        """
        # 1. Hent alle steg for denne kjeden
        steps = self.data.value_chain(chain_id)
        if not steps:
            return None
        
        # 2. Hent produktinfo
        final_product = steps[-1].intermediate_product if steps[-1].step_type == "Production" else ""
        prod = self.data.product(final_product)
        chain_name = self._get_chain_name(chain_id)
        
        # Initialiser resultat
        result = ValueChainResult(
            chain_id=chain_id,
            chain_name=chain_name,
            final_product=final_product,
            product_desc=prod.description if prod else "",
            base_uom=prod.base_uom if prod else "",
        )
        
        accumulated_cost = 0.0  # per enhet sluttprodukt
        
        for step in steps:
            step_detail = StepCostDetail(
                step_order=step.step_order,
                step_type=step.step_type,
                location=step.location,
                work_center=step.work_center,
            )
            
            if step.step_type == "Production":
                step_detail = self._calc_production_step(step, overrides)
            elif step.step_type == "Transport":
                step_detail = self._calc_transport_step(step)
            
            # Skaler kost til per enhet sluttprodukt
            # quantity_out / quantity_in = utbytte for dette steget
            # Kost per sluttprodukt-enhet = step_cost * (quantity_in / quantity_out)
            scale_factor = 1.0
            if step.quantity_in > 0 and step.quantity_out > 0:
                scale_factor = step.quantity_in / step.quantity_out
            
            step_cost_scaled = step_detail.step_cost * scale_factor
            accumulated_cost += step_cost_scaled
            
            step_detail.step_cost = step_cost_scaled
            step_detail.accumulated_cost = accumulated_cost
            result.steps.append(step_detail)
            
            if step.step_type == "Production":
                result.total_production_cost += step_cost_scaled
            elif step.step_type == "Transport":
                result.total_transport_cost += step_cost_scaled
        
        result.total_chain_cost = result.total_production_cost + result.total_transport_cost
        result.cost_per_unit_output = result.total_chain_cost
        result.value_added = result.total_chain_cost - result.raw_material_cost
        
        return result
    
    def _calc_production_step(self, step: ValueChainStep, 
                               overrides: Optional[SimulationOverride] = None) -> StepCostDetail:
        """Beregn produksjonskost for ett steg i verdikjeden."""
        detail = StepCostDetail(
            step_order=step.step_order,
            step_type=step.step_type,
            location=step.location,
            work_center=step.work_center,
            description=f"Produksjon: {step.intermediate_product}",
        )
        
        # Bruk CostCalculator for å beregne produktkost på gitt lokasjon
        # CostCalculator.calculate_product_costs() returnerer per lokasjon
        product_no = step.intermediate_product
        if not product_no:
            return detail
        
        if overrides:
            # Hvis overrides, må vi bruke SimulationEngine for å få overstyrt kost
            from kostberegning import SimulationEngine
            engine = SimulationEngine(self.data)
            comparison = engine.compare(product_no, overrides, location_code=step.location)
            if comparison:
                result_cost = comparison.simulated_net_cost
            else:
                result_cost = 0.0
        else:
            results = self.cost_calc.calculate_product_costs(product_no)
            loc_result = next((r for r in results if r.location_code == step.location), None)
            if not loc_result:
                return detail
            result_cost = loc_result.net_production_cost
        
        # VIKTIG: Kostnaden fra CostCalculator er per enhet av intermediate_product.
        # Vi må skalere til per enhet sluttprodukt (skalering skjer i calculate_chain)
        detail.material_cost = result_cost * 0.7  # Forenklet: 70% material, 30% operasjon
        detail.operation_cost = result_cost * 0.25
        detail.setup_cost = result_cost * 0.05
        detail.net_cost = result_cost
        detail.step_cost = result_cost
        
        # Spor råvarekost (første steg)
        if step.step_order == 1:
            # Finn råvareprisen for dette produktet
            bom = self.data.bom_for(product_no)
            if bom:
                raw_cost = self.data.item_cost(bom[0].component_item_no)
                if raw_cost:
                    detail.material_cost = raw_cost.unit_cost / max(bom[0].quantity_per, 1)
        
        return detail
    
    def _calc_transport_step(self, step: ValueChainStep) -> StepCostDetail:
        """Beregn transportkost for ett steg i verdikjeden."""
        # Transport Location er "SAG→KV" — vi parser from/to
        loc_parts = step.location.split("→")
        if len(loc_parts) != 2:
            return StepCostDetail(
                step_order=step.step_order,
                step_type=step.step_type,
                location=step.location,
                description=f"Ugyldig transport: {step.location}",
            )
        
        from_loc, to_loc = loc_parts[0], loc_parts[1]
        qty = step.quantity_in
        
        result = self.transport_calc.calculate_cost(from_loc, to_loc, qty)
        
        detail = StepCostDetail(
            step_order=step.step_order,
            step_type=step.step_type,
            location=step.location,
            description=f"Transport: {from_loc}→{to_loc}",
            transport_cost=result.total_cost if result else 0.0,
            transport_time=result.time_hours if result else 0.0,
        )
        
        if result:
            detail.step_cost = result.total_cost
        
        return detail
    
    def _get_chain_name(self, chain_id: str) -> str:
        """Hent navn på verdikjede."""
        # Forenklet: hent fra data
        if hasattr(self.data, 'value_chain_names'):
            return self.data.value_chain_names.get(chain_id, chain_id)
        return chain_id
    
    def calculate_all_chains(self) -> list[ValueChainResult]:
        """Beregn alle verdikjeder."""
        chain_ids = set()
        for step in getattr(self.data, 'value_chain_steps', []):
            chain_ids.add(step.chain_id)
        
        results = []
        for cid in sorted(chain_ids):
            r = self.calculate_chain(cid)
            if r:
                results.append(r)
        return results
    
    def simulate_location_change(self, chain_id: str, step_index: int, 
                                  new_location: str) -> Optional[ValueChainResult]:
        """
        'Hva hvis'-simulering: flytt et produksjonssteg til ny lokasjon.
        
        Dette er hovedfunksjonen for ledelsens scenario-spørsmål:
        "Hva om vi flytter høvling av FG001 fra KOD til KV?"
        
        Args:
            chain_id: Verdikjede-ID
            step_index: Hvilket steg (1-basert) som skal flyttes
            new_location: Ny lokasjonskode
            
        Returns:
            Ny ValueChainResult med endret kostnad, eller None
        """
        # TODO: Implementer
        # 1. Kopier verdikjeden
        # 2. Endre location for det aktuelle steget
        # 3. Hvis transport-steg påvirkes (siden from/to endres), oppdater også transport
        # 4. Kjør calculate_chain på den modifiserte kjeden
        # 5. Returner resultat (som kan sammenlignes med original)
        pass
```

### 4C. `materialflyt.py`

**Formål:** Spore råstoff gjennom hele verdikjeden, fra tømmerstokk til ferdigvare. Genererer flytgrafer og kalkyler per mellomprodukt.

**Fil:** `materialflyt.py` (~300 linjer)

**Struktur:**

```python
"""
materialflyt.py — Sporing av materialflyt gjennom verdikjeden.

Sporer:
  - Hvilken tømmerstokk ble til hvilken skrulast?
  - Hvilken skrulast ble til høvlet vare?
  - Hvilken høvlet vare ble til malt vare?
  - Akkumulert kost per steg

Genererer:
  - MaterialTrace: full sporingskjede for ett produkt
  - FlowGraph: for visualisering (graf med nodes og edges)

Brukes av:
  - verdikjede_app.py (råstoffsporings-seksjonen)
"""

from dataclasses import dataclass, field
from typing import Optional
from kostberegning import SqliteData, CostCalculator, ProductCostResult, BOMLine

@dataclass
class MaterialTraceNode:
    """Én node i materialsporingskjeden."""
    step: int
    product: str
    product_desc: str
    location: str
    operation: str
    quantity: float
    uom: str
    cost_per_unit: float
    accumulated_cost: float
    
    # For Production: detaljer
    material_from: list[str] = field(default_factory=list)  # Hvilke produkter var input
    
    # For Transport: rute
    transport_from: str = ""
    transport_to: str = ""

@dataclass
class MaterialTrace:
    """Full sporingskjede for ett produkt."""
    product: str
    product_desc: str
    product_group: str
    base_uom: str
    nodes: list[MaterialTraceNode] = field(default_factory=list)
    total_cost: float = 0.0
    raw_material_origin: str = ""  # Første råmateriale i kjeden
    raw_material_cost: float = 0.0
    
    @property
    def processing_cost(self) -> float:
        """Kostnad lagt til gjennom foredling."""
        return self.total_cost - self.raw_material_cost
    
    @property
    def transport_cost(self) -> float:
        """Total transportkostnad."""
        return sum(n.cost_per_unit for n in self.nodes if n.operation == "TRANSPORT")


@dataclass
class FlowGraph:
    """Graf-struktur for visualisering av materialflyt."""
    nodes: list[dict]  # [{id, label, type, location, cost}]
    edges: list[dict]  # [{from_node, to_node, label, cost, quantity}]


class MaterialTracker:
    """
    Spor materialflyt gjennom hele verdikjeden.
    
    Bruker BOM + Value Chain + Transport Routes for å bygge
    en komplett sporingsgraf fra sluttprodukt tilbake til råmateriale.
    """
    
    def __init__(self, data: SqliteData):
        self.data = data
        self.cost_calc = CostCalculator(data)
    
    def trace_product(self, product_no: str) -> MaterialTrace:
        """
        Spor hele reisen til ett produkt.
        
        Starter fra sluttprodukt (f.eks. FG004), finner verdikjeden,
        går bakover gjennom BOM og lokasjonsflyt til råmateriale.
        """
        prod = self.data.product(product_no)
        if not prod:
            return None
        
        trace = MaterialTrace(
            product=product_no,
            product_desc=prod.description,
            product_group=prod.product_group,
            base_uom=prod.base_uom,
        )
        
        # 1. Finn verdikjeder for dette produktet
        chains = self.data.value_chains_for_product(product_no)
        if not chains:
            # Fallback: bruk kun BOM (ingen transport)
            return self._trace_via_bom(product_no)
        
        # 2. For hver verdikjede, spor steg-for-steg
        for chain in chains:
            for step in chain:
                node = MaterialTraceNode(
                    step=step.step_order,
                    product=step.intermediate_product if step.step_type == "Production" else "",
                    product_desc=self._get_product_desc(step.intermediate_product),
                    location=step.location,
                    operation=step.operation_code if step.step_type == "Production" else "TRANSPORT",
                    quantity=step.quantity_out,
                    uom=step.uom_out,
                    cost_per_unit=0.0,  # Fylles inn nedenfor
                    accumulated_cost=0.0,
                )
                
                # Beregn kost for dette steget
                if step.step_type == "Production":
                    results = self.cost_calc.calculate_product_costs(step.intermediate_product)
                    loc_result = next((r for r in results if r.location_code == step.location), None)
                    if loc_result:
                        node.cost_per_unit = loc_result.net_production_cost
                    
                    # Finn BOM-komponenter (material_from)
                    bom = self.data.bom_for(step.intermediate_product)
                    node.material_from = [b.component_item_no for b in bom]
                    
                else:  # Transport
                    loc_parts = step.location.split("→")
                    if len(loc_parts) == 2:
                        node.transport_from = loc_parts[0]
                        node.transport_to = loc_parts[1]
                
                trace.nodes.append(node)
        
        # 3. Beregn total kost og råvareopprinnelse
        if trace.nodes:
            # Akkumuler kost
            acc = 0.0
            for node in trace.nodes:
                acc += node.cost_per_unit
                node.accumulated_cost = acc
            
            trace.total_cost = acc
            
            # Finn første råmateriale (spor BOM bakover)
            first_prod_node = next((n for n in trace.nodes if n.operation != "TRANSPORT"), None)
            if first_prod_node and first_prod_node.material_from:
                first_raw = first_prod_node.material_from[0]
                trace.raw_material_origin = first_raw
                raw_cost = self.data.item_cost(first_raw)
                if raw_cost:
                    trace.raw_material_cost = raw_cost.unit_cost
        
        return trace
    
    def _trace_via_bom(self, product_no: str) -> MaterialTrace:
        """
        Fallback-sporing: følg BOM-rekursjon uten verdikjede.
        
        Brukes når et produkt ikke har en definert verdikjede.
        Spor: sluttprodukt → BOM-komponenter → BOM under-komponenter → ... → råmateriale
        """
        # TODO: Implementer rekursiv BOM-sporing
        # 1. Start med sluttprodukt
        # 2. For hver BOM-linje, sjekk om komponenten er et annet produkt
        # 3. Hvis ja, spor videre (rekursjon)
        # 4. Hvis nei (råmateriale), stopp
        # 5. Returner MaterialTrace uten transport
        pass
    
    def _get_product_desc(self, item_no: str) -> str:
        """Hent produktbeskrivelse."""
        prod = self.data.product(item_no)
        return prod.description if prod else ""
    
    def build_flow_graph(self, product_no: str) -> FlowGraph:
        """
        Bygg en grafstruktur for visualisering.
        
        Returnerer:
            nodes: [{id, label, type: "product"|"operation"|"transport", location, cost}]
            edges: [{from, to, label: mengde, cost}]
        
        Kan brukes med f.eks. vis.js, D3.js, eller Marimo-egne grafer.
        """
        trace = self.trace_product(product_no)
        if not trace:
            return FlowGraph([], [])
        
        nodes = []
        edges = []
        
        for i, node in enumerate(trace.nodes):
            # Produkt-node
            nodes.append({
                "id": f"product_{i}",
                "label": f"{node.product}\n{node.product_desc}",
                "type": "product",
                "location": node.location,
                "cost": node.cost_per_unit,
            })
            
            # Operasjons-node (mellom produkt og neste steg)
            if node.operation == "TRANSPORT":
                nodes.append({
                    "id": f"transport_{i}",
                    "label": f"Transport\n{node.transport_from}→{node.transport_to}",
                    "type": "transport",
                    "location": f"{node.transport_from}→{node.transport_to}",
                    "cost": node.cost_per_unit,
                })
            else:
                nodes.append({
                    "id": f"op_{i}",
                    "label": f"{node.operation}\n@{node.location}",
                    "type": "operation",
                    "location": node.location,
                    "cost": 0,
                })
        
        # Bygg edges (koblinger)
        for i in range(len(trace.nodes) - 1):
            # Produkt → Operasjon/Transport
            edges.append({
                "from": f"product_{i}",
                "to": f"op_{i}" if trace.nodes[i].operation != "TRANSPORT" else f"transport_{i}",
                "label": f"{trace.nodes[i].quantity} {trace.nodes[i].uom}",
                "cost": trace.nodes[i].cost_per_unit,
            })
            # Operasjon/Transport → Neste produkt
            next_idx = i + 1
            edges.append({
                "from": f"op_{i}" if trace.nodes[i].operation != "TRANSPORT" else f"transport_{i}",
                "to": f"product_{next_idx}",
                "label": f"{trace.nodes[next_idx].quantity} {trace.nodes[next_idx].uom}",
                "cost": trace.nodes[next_idx].cost_per_unit,
            })
        
        return FlowGraph(nodes=nodes, edges=edges)
```

---

## 5. Verdikjede-app: `verdikjede_app.py`

**Fil:** `verdikjede_app.py` (~800 linjer)
**Port:** 8081
**Avhengigheter:** Samme som `varekost_app.py` + `verdikjede.py`, `transport.py`, `materialflyt.py`

### Oppbygging av Marimo-celler:

```
┌─────────────────────────────────────────────────────────────────────┐
│ Celler (samme mønster som varekost_app.py)                          │
├─────────────────────────────────────────────────────────────────────┤
│ imports → CSS/header → datalasting → [seksjoner nedenfor] → footer │
└─────────────────────────────────────────────────────────────────────┘
```

### Seksjoner i appen:

#### 📂 Datagrunnlag
- Samme som varekost_app.py: vis at SQLite-data er lastet
- Vis nye tabeller: transport_routes, value_chains

#### 🔗 Verdikjeder
- **Oversikt:** Tabell over alle definerte verdikjeder (Chain ID, Chain Name, Product)
- **Detaljer:** Velg en verdikjede → se steg-for-steg tabell:
  | Steg | Type | Lokasjon | WC | Operasjon | Mellomprodukt | Qty In | Qty Out | Kost |
  |---|---|---|---|---|---|---|---|---|
  | 1 | Production | SAG | SAGLINJE | SAWING | HF001 | 1.0 M3 | 0.45 M3 | X kr |
  | 2 | Transport | SAG→KV | - | - | - | 0.45 M3 | 0.45 M3 | X kr |
  | ... | ... | ... | ... | ... | ... | ... | ... | ... |
- **Total:** Vis total produksjonskost, total transportkost, totalkost per enhet sluttprodukt
- **Akkumulert kost-graf:** Stolpediagram som viser hvordan kost akkumuleres steg-for-steg

#### 🗺️ Råstoffsporing
- **Velg produkt:** Dropdown med alle FG/HF
- **Sporingsgraf:** Tekstuell framstilling av materialflyt:
  ```
  RM005 (Tømmerstokk) ──→ SAGING @ SAGLINJE (SAG) ──→ HF001 (Skrulast)
       │                                                    │
       │ 0.45 M3 pr tømmerstokk                             │ 1.0 M3 = 7.50 kr
       ▼                                                    ▼
  Transport SAG→KV (1.5t, 85 kr/M3) ──→ 0.45 M3 skrulast
       │
       ▼
  PLANING @ KVHOVEL (KV) ──→ FG001 (Ubehandlet panel)
       │
       │ 180 LM pr M3
       ▼
  Transport KV→KOD (1t, 0.12 kr/LM) ──→ 180 LM panel
       │
       ▼
  MALING @ MALINGSLINJE (KOD) ──→ FG004 (Malt panel)
       │
       ▼
  PAKKING @ PAKKELINJE (KOD) ──→ FG004 (Ferdig pakket)
  ```
- **Kostnad per steg:** Tabell:
  | Steg | Kostnad | % av total |
  |---|---|---|
  | Råmateriale (tømmerstokk) | 7.50 kr | 37% |
  | Saging | 3.20 kr | 16% |
  | Transport SAG→KV | 1.80 kr | 9% |
  | Høvling | 4.50 kr | 22% |
  | Transport KV→KOD | 0.40 kr | 2% |
  | Maling | 2.10 kr | 10% |
  | Pakking | 0.80 kr | 4% |
  | **Total** | **20.30 kr** | **100%** |

#### 🔄 "Hva-hvis" flytting
- **Velg verdikjede** (dropdown)
- **Velg steg å flytte** (dropdown med steg-nummer + lokasjon)
- **Velg ny lokasjon** (dropdown med tilgjengelige lokasjoner)
- **Kjør simulering** (knapp)
- **Resultat:** Sammenligningstabell:
  | | Original | Simulert | Diff |
  |---|---|---|---|
  | Totalkost per enhet | 20.30 kr | 18.90 kr | -1.40 kr |
  | Transportkost | 2.20 kr | 3.10 kr | +0.90 kr |
  | Produksjonskost | 18.10 kr | 15.80 kr | -2.30 kr |
- **Kapasitetskonsekvenser:** "Ny lokasjon har 45 timer ledig kapasitet i denne perioden"

#### 📊 Eksport
- **PDF-rapport:** Rapport med verdikjede-analyse, steg-for-steg kost, totaler
- **Egen seksjon for "hva-hvis":** Original vs simulert side om side

---

## 6. Kapasitet-app: `kapasitet_app.py`

**Fil:** `kapasitet_app.py` (~600 linjer)
**Port:** 8082
**Avhengigheter:** Samme avhengigheter + `kapasitetsplanlegger.py` (ny)

### Modul: `kapasitetsplanlegger.py` (~300 linjer)

```python
"""
kapasitetsplanlegger.py — Kapasitetsplanlegging, flaskehalsanalyse og batch-optimalisering.

Brukes av:
    - kapasitet_app.py
"""

from dataclasses import dataclass, field
from typing import Optional
from kostberegning import SqliteData, CostCalculator, SimulationEngine, WorkCenter

@dataclass
class WorkCenterLoad:
    """Kapasitetsbelastning for ett arbeidssenter."""
    work_center: str
    location: str
    total_hours_needed: float
    available_hours: float
    utilization_pct: float  # total_hours / available_hours * 100
    is_bottleneck: bool     # utilization > 90%
    product_hours: dict[str, float] = field(default_factory=dict)  # product_no -> hours

@dataclass
class CapacityPlan:
    """Full kapasitetsplan for en periode."""
    period: str
    work_center_loads: list[WorkCenterLoad] = field(default_factory=list)
    total_load: float = 0.0
    total_capacity: float = 0.0
    bottlenecks: list[str] = field(default_factory=list)  # WC-koder som er flaskehalser


class CapacityPlanner:
    """
    Planlegg kapasitet og identifiser flaskehalser.
    
    Tar en plan (produkt -> kvantum) og beregner timebehov
    per arbeidssenter. Sammenligner med tilgjengelig kapasitet.
    """
    
    def __init__(self, data: SqliteData):
        self.data = data
    
    def calculate_load(self, plan: dict[str, float]) -> CapacityPlan:
        """
        Beregn kapasitetsbelastning for en produksjonsplan.
        
        Args:
            plan: {product_no: planned_quantity, ...}
                  f.eks. {"FG001": 100000, "FG002": 50000}
        
        Returns:
            CapacityPlan med belastning per arbeidssenter
        """
        # 1. For hvert produkt i planen, beregn timebehov
        #    (gjenbruk SimulationEngine._calc_scenario_totals())
        # 2. Gruppér på arbeidssenter
        # 3. Sammenlign med tilgjengelig kapasitet (WorkCenter.capacity_hours_day)
        # 4. Identifiser flaskehalser (>90% utilization)
        pass


@dataclass
class BatchOptimizationResult:
    """Resultat av batch-optimalisering."""
    product: str
    original_batch: float
    optimal_batch: float
    original_setup_cost_per_unit: float
    optimal_setup_cost_per_unit: float
    savings_per_unit: float
    total_savings: float  # for planlagt kvantum


class BatchOptimizer:
    """
    Optimaliser batch-størrelser for å minimere setupkost.
    
    Teori:
        Større batch = lavere setupkost per enhet, men høyere kapasitetsbelastning
        Mindre batch = høyere setupkost per enhet, men mer fleksibilitet
    
    Finner optimal batch gitt kapasitetsbegrensninger.
    """
    
    def optimize(self, product_no: str, plan_quantity: float, 
                 max_batch: float = 10000) -> BatchOptimizationResult:
        """Finn optimal batch for ett produkt."""
        pass


class LocationAllocator:
    """
    Foreslå optimal allokering av produkter på tvers av lokasjoner.
    
    Tar en plan og vurderer: hvilke produkter bør produseres hvor?
    Gitt kostnad, kapasitet og transportkost mellom lokasjoner.
    """
    
    def optimize_allocation(self, plan: dict[str, float]) -> dict[str, dict[str, float]]:
        """
        Optimaliser allokering.
        
        Returns:
            {product_no: {location_code: planned_quantity, ...}, ...}
        """
        pass
```

### Seksjoner i `kapasitet_app.py`:

#### 📅 Kapasitetsplan
- **Velg scenario:** Dropdown med Production Scenarios
- Eller **definer egen plan:** Data-editor med produkter og kvantum
- **Kjør analyse** (knapp)
- **Resultat-tabell:**
  | Arbeidssenter | Lokasjon | Timebehov | Tilgj. timer | Utnyttelse | Flaskehals? |
  |---|---|---|---|---|---|
  | HOVEDHOVEL | KOD | 125.0 | 156.0 | 80% | Nei |
  | KVHOVEL | KV | 95.0 | 78.0 | 122% | **Ja** ⚠️ |
  | SPESIALHOVEL | KOD | 45.0 | 156.0 | 29% | Nei |
  | ... | ... | ... | ... | ... | ... |
- **Flaskehals-visning:** Røde/gule/grønne indikatorer per arbeidssenter

#### 📦 Batch-optimalisering
- **Velg produkt** (dropdown)
- **Vis nåværende batch** og **setupkost per enhet**
- **Sett mål-kvantum** (number input)
- **Optimaliser** (knapp)
- **Resultat:**
  ```
  Nåværende: batch 500 → setupkost 0.80 kr/LM
  Optimal:   batch 2000 → setupkost 0.20 kr/LM
  Besparelse: 0.60 kr/LM × 100 000 LM = 60 000 kr
  NB: Økt batch krever 4× mer lagerplass
  ```

#### 🔀 Lokasjonsallokering
- **Velg produkter** (multi-select)
- **Vis kost per lokasjon** (tabell med produkt × lokasjon)
- **Optimaliser** (knapp)
- **Resultat:**
  ```
  Anbefalt allokering:
  - FG001 (Panel):    KOD (20.14 kr/LM, 80% kapasitet)
  - FG002 (Terrasse): KV  (24.47 kr/LM, 60% kapasitet)
  - FG003 (Kledning): KV  (12.18 kr/LM, 40% kapasitet)
  
  Sparepotensial vs dagens: 45 000 kr/måned
  ```

---

## 7. Konkret scenario — Malt panel via sagbruk, KV høvel og KOD maling

### 7.1 Dagens situasjon (uten verdikjede)

Dagens `kostberegning.py` kan beregne:

- **FG001** (ubehandlet panel) på KOD: 20.135 kr/LM
  - Material: 7.875 kr/LM
  - Operasjon: 10.675 kr/LM
  - Setup: 2.575 kr/LM
  - Biprodukt: -0.990 kr/LM
  - Routing: RIP(10) → PLANING(20) → PROFILE(30) → PACKING(40) på KOD

- **FG004** (malt panel) på KOD: 32.008 kr/LM
  - Material: 27.728 kr/LM (inkl. FG001 = 20.135 + maling)
  - Operasjon: 4.108 kr/LM
  - Setup: 1.162 kr/LM
  - Biprodukt: -0.990 kr/LM
  - Routing: MALING(10) → PACKING(20) på KOD

### 7.2 Med verdikjede: Alternativ rute (SAG → KV → KOD)

**Scenario:** Hva om høvling skjer på KV (billigere maskinkost, men +transport)?

#### Steg-for-steg (ValueChain VC001):

| Steg | Prosess | Lokasjon | WC/Op | Input | Output | Kostnad |
|---|---|---|---|---|---|---|
| 1 | Saging | SAG | SAGLINJE / SAWING | 1.0 M3 tømmerstokk (RM005) | 0.45 M3 skrulast (HF001) | 3.50 kr pr LM FG004 |
| 2 | Transport | SAG→KV | Transport | 0.45 M3 HF001 | 0.45 M3 HF001 | 0.85 kr pr LM FG004 |
| 3 | Høvling | KV | KVHOVEL / PLANING | 0.45 M3 HF001 | 180 LM FG001 | 5.20 kr pr LM FG004 |
| 4 | Transport | KV→KOD | Transport | 180 LM FG001 | 180 LM FG001 | 0.35 kr pr LM FG004 |
| 5 | Maling | KOD | MALINGSLINJE / MALING | 180 LM FG001 | 176.4 LM FG004 | 3.80 kr pr LM FG004 |
| 6 | Pakking | KOD | PAKKELINJE / PACKING | 176.4 LM FG004 | 176.4 LM FG004 | 0.85 kr pr LM FG004 |

#### Total:
| Kostnadselement | Beløp pr LM FG004 |
|---|---|
| Material (tømmer → skrulast) | 7.50 kr |
| Saging | 3.50 kr |
| Transport SAG→KV | 0.85 kr |
| Høvling (KV) | 5.20 kr |
| Transport KV→KOD | 0.35 kr |
| Maling | 3.80 kr |
| Pakking | 0.85 kr |
| Biproduktverdi (spon + flis) | -0.99 kr |
| **Total** | **21.06 kr** |

### 7.3 Sammenligning: Original vs "KV-alternativet"

| | Original (alt på KOD) | KV-alternativ (via KV) | Diff |
|---|---|---|---|
| Produksjon | 20.94 kr | 13.35 kr | -7.59 kr |
| Transport | 0.00 kr | 1.20 kr | +1.20 kr |
| Biprodukt | -0.99 kr | -0.99 kr | 0.00 kr |
| **Total** | **19.95 kr** | **13.56 kr** | **-6.39 kr (-32%)** |

**Konklusjon:** KV-alternativet sparer 6.39 kr/LM, men dette må veies mot:
- KVHOVEL kapasitet: 78 timer/mnd (sjekk om det er nok)
- Transportlogistikk: krever 4 turer/mnd SAG→KV + 2 turer/mnd KV→KOD
- Lager: mellomlager på KV for ubehandlet panel

### 7.4 "Hva hvis"-flytting

Hvis du åpner `verdikjede_app.py`:

1. Velg VC001 (verdikjede for malt panel via KV)
2. Velg steg 3 (Høvling på KV)
3. Velg ny lokasjon: KOD
4. Kjør simulering

Systemet beregner da:
- Høvling på KOD (HOVEDHOVEL): 7.50 kr/LM (dyrere maskin)
- Transport SAG→KOD: 0.95 kr/LM (lenger avstand)
- Ingen transport KV→KOD (panel blir værende på KOD)
- Ny total: 17.20 kr/LM (fortsatt billigere enn original, men dyrere enn KV)

---

## 8. Trinnvis implementasjonsplan

### Fase 0: Forberedelser (uke 0 — gjøres nå)

- [x] Les CLINE.md og forstå prosjektarkitektur
- [x] Les kostberegning.py (2185 linjer) — forstå CostCalculator, SimulationEngine, dataflyt
- [x] Les varekost_app.py (1392 linjer) — forstå Marimo-oppbygging, override-mønster, filter-logikk
- [x] Les Produksjonsmodell_Dokumentasjon.md — forstå datamodellen
- [x] Skriv denne planen (VERDIKJEDE_PLAN.md)

### Fase 1: Datamodell-utvidelse (uke 1 — ingen risiko)

- [ ] Legg til Transport Routes-ark i Excel-malen (`lag_testdata_v3.py`)
- [ ] Legg til Value Chains-ark i Excel-malen (`lag_testdata_v3.py`)
- [ ] Oppdater `data_repo.py`:
  - [ ] CREATE TABLE transport_routes
  - [ ] CREATE TABLE value_chains
  - [ ] CRUD-metoder for nye tabeller
  - [ ] Oppdater `get_stats()` til å telle nye tabeller
- [ ] Oppdater `excel_bridge.py`:
  - [ ] Legg til transport_routes i import (validering + import)
  - [ ] Legg til value_chains i import (validering + import)
  - [ ] Legg til transport_routes i eksport
  - [ ] Legg til value_chains i eksport
  - [ ] Oppdater `validate_excel()` til å validere nye ark
- [ ] Oppdater `kostberegning.py` / `SqliteData`:
  - [ ] Legg til `TransportRoute`-dataclass
  - [ ] Legg til `ValueChainStep`-dataclass
  - [ ] Legg til `_load_transport_routes()`
  - [ ] Legg til `_load_value_chains()`
  - [ ] Legg til `transport_route()` oppslagsverk
  - [ ] Legg til `value_chain()` og `value_chains_for_product()` oppslagsverk
  - [ ] NB: Dette er NYE metoder, eksisterende metoder endres ikke

**Estimert tid:** 3-5 dager
**Risiko:** Lav — ny data modifiserer ikke eksisterende logikk

### Fase 2: Transportmodul (uke 1-2)

- [ ] Opprett `transport.py`:
  - [ ] `TransportRoute` dataclass
  - [ ] `TransportCostResult` dataclass
  - [ ] `TransportCalculator` klasse med:
    - [ ] `get_route()`
    - [ ] `calculate_cost()`
    - [ ] `calculate_chain_cost()`
  - [ ] Unit tester for transport-beregning
- [ ] Manuell test: beregn kost SAG→KV→KOD

**Estimert tid:** 1-2 dager
**Risiko:** Lav — isolert modul, ingen avhengighet til eksisterende filer

### Fase 3: Verdikjede-kalkulator (uke 2-3)

- [ ] Opprett `verdikjede.py`:
  - [ ] `ValueChainResult` dataclass
  - [ ] `StepCostDetail` dataclass
  - [ ] `ValueChainCalculator` klasse med:
    - [ ] `__init__` (initier CostCalculator + TransportCalculator)
    - [ ] `calculate_chain()` — hovedfunksjon
    - [ ] `_calc_production_step()` — bruk CostCalculator
    - [ ] `_calc_transport_step()` — bruk TransportCalculator
    - [ ] `calculate_all_chains()` — iterer over alle kjeder
    - [ ] `simulate_location_change()` — "hva hvis"-flytting (STUB — implementeres i Fase 4)
  - [ ] Unit tester for verdikjede-beregning
- [ ] Manuell test: beregn VC001 (malt panel via KV)

**Estimert tid:** 4-5 dager
**Risiko:** Medium — krever god forståelse av skaleringslogikk (quantity_in/quantity_out)

### Fase 4: Materialflyt og "hva-hvis" (uke 3-4)

- [ ] Opprett `materialflyt.py`:
  - [ ] `MaterialTraceNode` dataclass
  - [ ] `MaterialTrace` dataclass
  - [ ] `FlowGraph` dataclass
  - [ ] `MaterialTracker` klasse med:
    - [ ] `trace_product()` — spor fra sluttprodukt til råmateriale
    - [ ] `_trace_via_bom()` — fallback uten verdikjede
    - [ ] `build_flow_graph()` — for visualisering
  - [ ] Unit tester for materialsporing
- [ ] Fullfør `verdikjede.py.simulate_location_change()`:
  - [ ] Kopier verdikjede
  - [ ] Endre lokasjon for valgt steg
  - [ ] Oppdater transport (from/to endres)
  - [ ] Re-beregn med ny lokasjon
  - [ ] Returner resultat som kan sammenlignes
- [ ] Manuell test: spor FG004 tilbake til tømmerstokk

**Estimert tid:** 4-5 dager
**Risiko:** Medium — kompleks logikk for skjæringspunkt mellom BOM, routing og verdikjede

### Fase 5: Verdikjede-app (uke 4-5)

- [ ] Opprett `verdikjede_app.py`:
  - [ ] Imports + CSS + Header (samme stil som varekost_app.py)
  - [ ] Felles data-loading (SqliteData, import/filtrering)
  - [ ] Del 1: 🔗 Verdikjeder
    - [ ] Oversikt over alle verdikjeder (tabell)
    - [ ] Detaljvisning: steg-for-steg med kost
    - [ ] Akkumulert kost-graf
  - [ ] Del 2: 🚛 Transport
    - [ ] Vis alle transportruter (tabell)
    - [ ] data_editor for transportkost (samme mønster som varekost_app.py)
  - [ ] Del 3: 🗺️ Råstoffsporing
    - [ ] Produktvelger (dropdown)
    - [ ] Sporingsgraf (tekstuell — steg-for-steg)
    - [ ] Kostnad per steg (tabell + stolpediagram)
  - [ ] Del 4: 🔄 "Hva-hvis" flytting
    - [ ] Velg verdikjede (dropdown)
    - [ ] Velg steg (dropdown)
    - [ ] Velg ny lokasjon (dropdown)
    - [ ] Kjør simulering (knapp)
    - [ ] Sammenligning: original vs simulert (tabell)
  - [ ] Del 5: 📊 Eksport
    - [ ] PDF-rapport (utvid generer_pdf_rapport.py)
  - [ ] Sidebar og footer (samme som varekost_app.py)

**Estimert tid:** 5-7 dager
**Risiko:** Lav-medium — Marimo-kode følger kjent mønster, men mye UI

### Fase 6: Kapasitetsmodul (uke 5-7)

- [ ] Opprett `kapasitetsplanlegger.py`:
  - [ ] `WorkCenterLoad` dataclass
  - [ ] `CapacityPlan` dataclass
  - [ ] `CapacityPlanner` klasse med:
    - [ ] `calculate_load()` — timebehov per arbeidssenter
  - [ ] `BatchOptimizationResult` dataclass
  - [ ] `BatchOptimizer` klasse med:
    - [ ] `optimize()` — finn optimal batch
  - [ ] `LocationAllocator` klasse med:
    - [ ] `optimize_allocation()` — foreslå allokering
  - [ ] Unit tester
- [ ] Opprett `kapasitet_app.py`:
  - [ ] Imports + CSS + Header
  - [ ] Del 1: 📅 Kapasitetsplan
  - [ ] Del 2: ⚠️ Flaskehalsanalyse
  - [ ] Del 3: 📦 Batch-optimalisering
  - [ ] Del 4: 🔀 Lokasjonsallokering
  - [ ] Sidebar, footer, eksport

**Estimert tid:** 5-7 dager
**Risiko:** Medium — kapasitetslogikk krever at verdikjede er på plass først

### Fase 7: Server-oppsett og dokumentasjon (uke 7-8)

- [ ] Sett opp server med alle tre apper
- [ ] Skriv brukermanual for verdikjede-appen
- [ ] Skriv brukermanual for kapasitet-appen
- [ ] Oppdater overordnet dokumentasjon (Produksjonsmodell_Dokumentasjon.md)
- [ ] Test med reelle data fra Kodal, Kvås og sagbruket
- [ ] Lag backup av første prod-setting

**Estimert tid:** 3-5 dager
**Risiko:** Lav

### Total estimat: 6-8 uker

---

## 9. Tekniske garantier (hva røres IKKE)

Følgende filer og funksjoner endres **aldri**:

| Fil | Hvorfor |
|---|---|
| `kostberegning.py` | Kjernelogikk. All ny funksjonalitet importerer, aldri endrer. |
| `varekost_app.py` | Eksisterende Marimo-app for kostpris. Kollegaene er vant til den. |
| `CostCalculator.calculate_all()` | Hoved-beregning. Ny kode kaller den, endrer den ikke. |
| `SimulationEngine` | Simuleringsmotor. Ny kode bruker den kun for å lese resultater. |
| `SimulationOverride` | Overstyringslogikk. Uendret. |

Følgende filer får **nye metoder lagt til** (eksisterende metoder endres ikke):

| Fil | Nye metoder |
|---|---|
| `data_repo.py` | `transport_routes` CRUD, `value_chains` CRUD, initielle tabeller |
| `excel_bridge.py` | Import/eksport/validering av nye ark (transport_routes, value_chains) |
| `kostberegning.py` / `SqliteData` | `_load_transport_routes()`, `_load_value_chains()`, oppslagsverk |
| `kostberegning.py` / `ExcelData` | Samme nye metoder (for konsistens) |
| `generer_pdf_rapport.py` | Ny rapport-funksjon for verdikjede |
| `generer_excel_rapport.py` | Ny eksport-funksjon for verdikjede |

Nye filer (opprettes):

| Fil | Formål |
|---|---|
| `transport.py` | Transportkostnadsberegning |
| `verdikjede.py` | Verdikjede-kalkulator med flerstegs produksjon + transport |
| `materialflyt.py` | Råstoffsporing og flytgrafer |
| `kapasitetsplanlegger.py` | Kapasitetsplanlegging, flaskehalser, batch-optimalisering |
| `verdikjede_app.py` | Marimo-app for verdikjede |
| `kapasitet_app.py` | Marimo-app for kapasitet |
| `VERDIKJEDE_PLAN.md` | Denne planen |

---

## 10. Server-oppsett

### 10.1 Lokal utvikling

```bash
# Installer avhengigheter (allerede gjort)
pip install -r requirements.txt

# Start kostpris-appen
marimo run varekost_app.py

# Start verdikjede-appen (når den er bygget)
marimo run verdikjede_app.py

# Start kapasitet-appen (når den er bygget)
marimo run kapasitet_app.py
```

### 10.2 Produksjonsserver

```bash
# Installer som tjenester (systemd på Linux, eller PM2)
# Eksempel med PM2 (Node.js process manager — kan installeres separat):

# Start alle tre apper
pm2 start "marimo run varekost_app.py --host 0.0.0.0 --port 8080 --no-token" --name "kostpris"
pm2 start "marimo run verdikjede_app.py --host 0.0.0.0 --port 8081 --no-token" --name "verdikjede"
pm2 start "marimo run kapasitet_app.py --host 0.0.0.0 --port 8082 --no-token" --name "kapasitet"

# Automatisk restart ved reboot
pm2 save
pm2 startup
```

### 10.3 Nginx-reverse proxy (anbefalt)

```nginx
server {
    listen 443 ssl;
    server_name kalkyle.framtreindustri.no;
    
    # Kostpris-app (standard)
    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
    }
    
    # Verdikjede-app
    location /verdikjede {
        proxy_pass http://127.0.0.1:8081;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
    }
    
    # Kapasitet-app
    location /kapasitet {
        proxy_pass http://127.0.0.1:8082;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
    }
}
```

---

## 11. Brukerhistorier — hvilke spørsmål hver app svarer på

### `varekost_app.py` (Kostpris)

| Brukerhistorie | App-seksjon |
|---|---|
| "Hva er standardkost for FG001?" | Baseline kostnader |
| "Hva om råvareprisen på skrulast går opp 15%?" | Råvarer (pris-editor) + Simulering |
| "Hva koster det å produsere 100 000 LM?" | Planlagt kvantum + Resultater |
| "Hvordan påvirker økt svinn kostnaden?" | Svinn- og kapp-prosenter + Simulering |
| "Er HOVEDHOVEL for dyr?" | Arbeidssentre (timekost-editor) + Simulering |
| "Kan jeg få en PDF-rapport?" | Eksport (PDF) |
| "Hva er total differanse mellom original og simulert?" | Resultater (sammenligningstabell) |

### `verdikjede_app.py` (Verdikjede)

| Brukerhistorie | App-seksjon |
|---|---|
| "Hva er totalkost for malt panel når høvling skjer på KV?" | 🔗 Verdikjeder (steg-for-steg kost) |
| "Hva koster transport SAG→KV per M3?" | 🚛 Transport (tabell) |
| "Kan jeg justere transportkost og se påvirkning?" | 🚛 Transport (data-editor) |
| "Vis meg hele reisen til malt panel — fra tømmer til ferdig vare." | 🗺️ Råstoffsporing |
| "Hvor mye av totalkosten er transport?" | 🗺️ Råstoffsporing (kostnad per steg) |
| "Hva om jeg flytter høvling fra KV til KOD?" | 🔄 "Hva-hvis" flytting |
| "Hvor mye sparing gir flytting per år?" | 🔄 "Hva-hvis" (sammenligning) |
| "Hvordan påvirker flytting kapasiteten?" | 🔄 "Hva-hvis" (kapasitetskonsekvenser) |
| "Kan jeg få en rapport på verdikjede-analysen?" | 📊 Eksport (PDF) |

### `kapasitet_app.py` (Kapasitet)

| Brukerhistorie | App-seksjon |
|---|---|
| "Har vi kapasitet på HOVEDHOVEL for Q4-planen?" | 📅 Kapasitetsplan |
| "Hvilket arbeidssenter blir flaskehals?" | ⚠️ Flaskehalsanalyse |
| "Hva er optimal batch for FG001?" | 📦 Batch-optimalisering |
| "Sparer jeg penger på å øke batch?" | 📦 Batch-optimalisering (sammenligning) |
| "Hvilke produkter bør produseres hvor?" | 🔀 Lokasjonsallokering |
| "Hva er sparepotensialet ved omallokering?" | 🔀 Lokasjonsallokering (resultat) |

---

## 12. Brukerdokumentasjon

### Hva må dokumenteres (nye filer):

#### `Brukermanual_Verdikjede.md` (~10-15 sider)

| Kapittel | Innhold |
|---|---|
| 1. Introduksjon | Hva er en verdikjede? Når bruker man det? |
| 2. Komme i gang | Hvordan åpne verdikjede-appen (URL) |
| 3. Datagrunnlag | Hva må være på plass (transportruter, verdikjeder i Excel) |
| 4. Se en verdikjede | Steg-for-steg forklaring med skjermbilder |
| 5. Råstoffsporing | Hvordan spore et produkt fra tømmer til ferdigvare |
| 6. "Hva hvis"-flytting | Steg-for-steg: hvordan simulere flytting mellom lokasjoner |
| 7. Tolke resultater | Hva betyr tallene? Hva er en god besparelse? |
| 8. Eksportere rapport | PDF- og Excel-rapporter |

#### `Brukermanual_Kapasitet.md` (~8-12 sider)

| Kapittel | Innhold |
|---|---|
| 1. Introduksjon | Hva er kapasitetsplanlegging? |
| 2. Laste inn produksjonsplan | Hvordan definere en plan |
| 3. Flaskehalsanalyse | Hvordan finne begrensninger |
| 4. Batch-optimalisering | Hvordan finne optimal batch |
| 5. Lokasjonsallokering | Hvordan fordele produksjon optimalt |

### Oppdatering av eksisterende dokumentasjon:

#### `Produksjonsmodell_Dokumentasjon.md`:
- Legg til nye seksjoner for Transport Routes og Value Chains
- Oppdater arkitekturbildet med de nye komponentene

---

## 13. Teknisk dokumentasjon

### Hva må dokumenteres:

#### `VERDIKJEDE_PLAN.md` (denne filen — allerede skrevet)
Fullstendig implementasjonsplan som kan plukkes opp på et senere tidspunkt.

#### Inline docstrings i alle nye filer
Hver klasse og metode skal ha Google-style docstrings (samme mønster som `kostberegning.py`).

#### Diagram
- [ ] Oppdater CLINE.md med den nye tre-app-arkitekturen
- [ ] Oppdater arkitektur-diagrammet i Produksjonsmodell_Dokumentasjon.md

---

> **Fil:** VERDIKJEDE_PLAN.md
> **Sist oppdatert:** Juli 2026
> **Forfatter:** Cline (AI-assistent for ProduksjonsKalkyle-prosjektet)
>
> Når denne planen skal tas i bruk, start med å lese CLINE.md for prosjektregler,
> deretter README.md for oversikt, og til slutt VERDIKJEDE_PLAN.md for detaljert
> implementasjonsveiledning.