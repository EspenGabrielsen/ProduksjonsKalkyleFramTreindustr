# TODO: Manglende filtrering i Python-scriptene

## `active`-flagg (ignoreres i dag)
- [ ] **Product.active** — filtrer bort inaktive produkter i `calculate_all()` og `calculate_product_costs()`
- [ ] **WorkCenter.active** — filtrer bort inaktive arbeidssentre ved oppslag og routing
- [ ] **Location.active** — filtrer bort inaktive lokasjoner
- [ ] **Operation.active** — filtrer bort inaktive operasjoner

## Dato-felter (ignoreres i dag)
- [ ] **ItemCost.effective_date** — velg riktig pris basert på en angitt analysedato (ikke bare nyeste)
- [ ] **BOMLine.valid_from / valid_to** — filtrer BOM-linjer på datoperiode
- [ ] **RoutingLine.valid_from / valid_to** — filtrer routing-linjer på datoperiode
- [ ] **Scenario.start_date / end_date** — bruk datoene i scenario-simulering

## Duplikate priser
- [ ] Hvis samme `Effective Date`: siste rad vinner?
- [ ] Hvis ulike datoer: velg riktig pris basert på analysedato (f.eks. dagens dato)