# Transportvare-funksjonalitet — Implementeringsplan

> **Opprettet:** 31. juli 2026  
> **Status:** Plan godkjent, klar for implementering  
> **Mål:** Automatisk generering av semi-finished produkter og transport-routing når en vare flagges som "transportvare" i SQLite.

---

## Konsept

Når en vare flagges som transportvare (f.eks. JD16073), genererer systemet automatisk:

1. Semi-finished varianter per lokasjon: JD16073-KOD, JD16073-KV, JD16073-EIK
2. BOM for hovedvaren: JD16073 → JD16073-KOD (Qty Per = 1)
3. Routing for transport: TRANSPORT @ FRAKT_KOD_KV

Når flagget fjernes, reverseres alle operasjonene og varen er tilbake til normal.

---

## Testdata (hentet fra aktiv database)

| Produkt | Høvleri | BOM | Routing | Hvorfor valgt |
|---------|---------|-----|---------|---------------|
| **BL98520** | Kodal (HOVEDHOVEL) | RM → FG | 1 operasjon | Enkleste case — ubehandlet byggelist |
| **JD16073** | Kodal (SPESIALHOVEL) | RM → FG + co-prod | 1 operasjon | Har co-product (JD16073B) |
| **JD16098TF** | Kodal (HOVEDHOVEL) | JD16098(SemiF) + maling | 1 operasjon | Eksisterende produksjonskjede |

**Database-kontekst:**
```python
from data_repo import DataRepo
db = DataRepo()
db.initialize()
# Tabeller, locations, work_centers etc.
```

**Aktive locations:** KOD (Kodal Fabrikk), KV (Kvaas), EIK (Eikås), SKI (Skien Lager)
**Aktive work centers:** HOVEDHOVEL (KOD), SPESIALHOVEL (KOD), KVHOVEL (KV), EIKHOVEL (EIK), MALINGSLINJE (KOD), IMPREGNERING (EIK), PAKKELINJE (KOD), FRAKT_KOD_KV, FRAKT_KOD_EIK, FRAKT_KV_EIK osv.

---

## Arkitektur

```
┌─────────────────────────────────────────────────────────────────────┐
│                      Excel-import                                   │
│  excel_bridge.py::import_excel_to_sqlite()                          │
│  └── sync_transport_varer(db)   ← NY — kjører etter import         │
└───────────────────────────────────┬─────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    SQLite-tabeller (nye)                            │
│                                                                     │
│  transport_flagg                    transport_routes                 │
│  ┌──────────────────────────┐      ┌───────────────────────────┐   │
│  │ item_no TEXT PRIMARY KEY  │      │ from_loc TEXT NOT NULL     │   │
│  │ is_transport INTEGER     │      │ to_loc TEXT NOT NULL       │   │
│  └──────────────────────────┘      │ work_center_code TEXT      │   │
│                                     │ distance_km REAL           │   │
│                                     │ cost_per_km REAL           │   │
│                                     │ run_time_minutes REAL      │   │
│                                     │ setup_time_minutes REAL    │   │
│                                     │ batch_size REAL            │   │
│                                     └───────────────────────────┘   │
└───────────────────────────────────┬─────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│                   sync_transport_varer(db)                           │
│                                                                     │
│  Når is_transport = 1:                                              │
│    1. For hvert produkt med flagg = 1:                              │
│       a. Les original BOM (alle component_item_no)                  │
│       b. Les original Routing (alle linjer)                         │
│       c. For hver aktiv location (KOD, KV, EIK):                    │
│          - Opprett semi-finished: {vare}-{loc}                      │
│          - Kopier original BOM til semi-finished                    │
│          - Oppdater routing til riktig arbeidssenter per loc        │
│       d. Oppdater hovedvarens BOM: {vare} → {vare}-KOD (Qty=1)      │
│       e. For hver transport-rute:                                   │
│          - Legg til TRANSPORT-operasjon i hovedvarens routing       │
│                                                                     │
│  Når is_transport = 0:                                              │
│    1. For hvert produkt med flagg = 0:                              │
│       a. Slett semi-finished: {vare}-KOD, {vare}-KV, {vare}-EIK    │
│       b. Fjern semi-finished fra hoved-BOM                          │
│       c. Fjern TRANSPORT-routing fra hovedvaren                     │
│       d. Gjenopprett original BOM (fjern semi-finished refs)        │
│    2. Endringslogg fanger opp alle CREATE/DELETE                    │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Filer som endres/opprettes

| Fil | Endring | Linjer |
|-----|---------|--------|
| `src/scripts/test_transport.py` | **NY** — Testscript med in-memory DB | ~200 |
| `src/data_repo.py` | Legg til `transport_flagg`-tabell | ~15 |
| `src/excel_bridge.py` | Legg til `sync_transport_varer()` | ~150 |
| `docs/Produksjonsmodell_Dokumentasjon.md` | Dokumentasjon av transportvare-funksjon | ~60 |

---

## Testplan (5 tester, in-memory SQLite)

### Test 1: Sett flagg på enkelt produkt (BL98520)
```
Input:  BL98520 flagges som transportvare
Verify:
  - BL98520-KOD, BL98520-KV, BL98520-EIK opprettet i products
  - Semi-finished har kopi av original BOM
  - BL98520 har BOM-linje: BL98520-KOD (Qty Per = 1)
  - BL98520 har TRANSPORT-routing lagt til
  - CostCalculator gir høyere kostnad pga frakt
  - Endringslogg inneholder CREATE for alle nye rader
```

### Test 2: Fjern flagg (BL98520 tilbake til normal)
```
Input:  BL98520 avflagges (is_transport = 0)
Verify:
  - BL98520-KOD, BL98520-KV, BL98520-EIK slettet
  - BL98520 BOM tilbake til original (kun RM)
  - BL98520 routing tilbake til original (kun HOVLING)
  - Kostnad identisk med pre-flagg-beregning
  - Endringslogg inneholder DELETE for alle fjernede rader
```

### Test 3: Co-produkt overlever sync (JD16073)
```
Input:  JD16073 flagges (har co-product JD16073B)
Verify:
  - Semi-finished får kopiert co-prod-BOM
  - JD16073-KOD har co_product_item_no = JD16073B-KOD
  - Hoved-JD16073 beholder co-prod BOM-linje
```

### Test 4: Produksjonskjede — semi-finished dobles ikke (JD16098TF)
```
Input:  JD16098TF flagges (JD16098 er allerede Semi Finished)
Verify:
  - JD16098TF-KOD/KV/EIK opprettes
  - JD16098 genereres IKKE på nytt (er allerede Semi Finished, ikke Finished Good)
  - BOM: JD16098TF-KOD → JD16098 (Qty Per = 1, kopiert fra original)
```

### Test 5: Alle tre samtidig — ingen kollisjoner
```
Input:  BL98520, JD16073, JD16098TF flagges samtidig
Verify:
  - Totalt 9 semi-finished produkter (3 × 3)
  - Ingen duplikate varenummer
  - CostCalculator kjører uten feil for alle
  - Alle produkter har ulik kostnad per lokasjon (pga ulike WC timepriser)
```

---

## Kjøring

```bash
# Testtransport (in-memory, påvirker IKKE aktiv database)
python src/scripts/test_transport.py

# Forventet output:
#   TEST 1: Sett flagg BL98520 .......... OK
#   TEST 2: Fjern flagg BL98520 ......... OK
#   TEST 3: Co-produkt (JD16073) ........ OK
#   TEST 4: Produksjonskjede (JD16098TF)  OK
#   TEST 5: Alle tre samtidig ........... OK
#   =========================================
#   5/5 tester bestått — transportvaremodul klar
```

---

## Eksempel: JD16073 fra Kodal til Kvås

```
JD16073 (Finished Good — selges)
├── BOM:
│   ├── JD16073-KOD (Qty Per = 1)          ← Semi-finished, generert
│   └── RM_50x75_US_V_Gran (Qty Per = 533.05, co_prod_pct = 0)
├── Routing:
│   ├── Op 10: HOVLING @ SPESIALHOVEL (KOD)  ← Original (kopiert til semi-finished)
│   └── Op 20: TRANSPORT @ FRAKT_KOD_KV      ← NY — fraktkostnad
└── Kostnad = JD16073-KODs produksjonskost + fraktkost

JD16073-KOD (Semi Finished — intern, generert)
├── BOM:
│   └── RM_50x75_US_V_Gran (Qty Per = 533.05, co_prod_pct = 0.5, co_prod_item_no = JD16073B-KOD)
├── Routing:
│   └── Op 40: HOVLING @ SPESIALHOVEL (KOD)
└── Kostnad = RM-kost + SPESIALHOVEL operasjonskost

JD16073B-KOD (Semi Finished — intern, generert fra co-prod)
├── BOM:
│   └── RM_50x75_US_V_Gran (samme som A-vare)
└── Routing:
    └── Op 40: HOVLING @ SPESIALHOVEL (KOD) (forholdsmessig andel)
```

---

## Kostnadseksempel (med hypotetiske fraktkostnader)

```
JD16073-KOD (produsert på Kodal):
  Materialkost:    1 600,00 kr  (RM × 533,05 m/M3)
  Operasjonskost:    280,00 kr  (SPESIALHOVEL: 0.029 min/LM × 1 605 kr/t / 0.94 eff)
  Setupkost:          15,00 kr  (30 min setup / 3 200 batch)
  Biproduktverdi:    -48,00 kr  (JD16073B-KOD)
  Netto Kodal:     1 847,00 kr  ← DETTE blir JD16073s materialkost

JD16073 (solgt fra Kvås):
  Materialkost:    1 847,00 kr  (JD16073-KODs netto kost)
  Transportkost:     112,50 kr  (FRAKT_KOD_KV: 45 min × 800 kr/t / batch 3 200)
  Netto Kvås:      1 959,50 kr  ← KUNDEN BETALER
```

---

## Dokumentasjon som må oppdateres

Når implementasjonen er ferdig og testene kjører grønt:

1. **`docs/Produksjonsmodell_Dokumentasjon.md`** — Nytt kapittel: "Transportvarer og fler-høvleri-produksjon"
2. **`docs/Brukermanual_Produksjonsmodell.md`** — Forklare transport-flagg i Excel
3. **`CLINE.md`** — Legge til `test_transport.py` og transport-flagg-tabell

---

## Rollback-plan

Hvis noe feiler:
1. `sync_transport_varer()` kan kjøres i revers: alle genererte semi-finished slettes, flagg fjernes
2. Endringsloggen viser nøyaktig hva som ble endret og når
3. `test_transport.py` bruker `:memory:` database — null risiko for produksjonsdata

---

> **Plan versjon 1.0**  
> **Neste steg:** Bytt til ACT MODE for å implementere