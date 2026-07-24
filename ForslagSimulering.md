
# Prosjektkonsept: Volumsimulering og lokasjonsallokering i ProduksjonsKalkyle

## Bakgrunn og mål

* **Nåværende system:** En Python/Marimo-basert kalkulasjonsmodell (`ProduksjonsKalkyle`) som beregner stykkliste (BOM) og ruting-kostnader per lokasjon. Logikken speiler Microsoft Dynamics 365 Business Central.


* **Mål:** Utvide modellen fra å kun vise statisk enhetskostnad per fabrikk (x og y), til å simulere en **vektet årlig gjennomsnittskostpris** basert på hvordan årlig volum fordeles mellom fabrikkene. Dette skal brukes som et taktisk beslutningsverktøy før fremtidig ERP-migrering.



## Funksjonelle krav til utvidelsen

### 1. Datagrunnlag (Årlig forecast)

* Det må opprettes en ny tabell/datakilde (`demand_forecast` eller `volume_allocation`) som holder styr på totalt forventet årsforbruk per vare, samt en foreslått fordeling per lokasjon/fabrikk.

### 2. Dynamisk grensesnitt basert på ruting (Lokasjons-matrise)

Systemet må sjekke hvilke fabrikker som faktisk har ruting på den valgte varen, og tilpasse grensesnittet dynamisk:

* **2 fabrikker:** Én standard slider (0–100 %).
* **3 fabrikker:** En visuell eller logisk **trekant-slider (Ternary/Simplex-logikk)**.
* **4+ fabrikker:** En redigerbar tabell (`mo.ui.data_editor`) med validering på at summen er 100 %.


### 3. Simuleringstilstander (De tre punktene)

I grensesnittet (spesielt ved 3 fabrikker) skal tre scenarier visualiseres og sammenlignes:

1. **Dagens fordeling (Baseline):** Hentet fra historiske data eller gjeldende årsplan.
2. **Optimal fordeling (Best-case):** En algoritmisk beregnet fordeling som gir lavest mulig vektet kostnad basert på kapasitet.
3. **Simulert fordeling (What-if):** Brukerens interaktive justeringer i sanntid.
4. **simulering** uavhengig av antall fabrikker må antall timer totalt beregnes løpende per fabrikk slik at vi ikke legger mer til en fabrikk enn tilgjengelig kapasitet.

### 4. Beregningslogikk og KPI-er

Når volumet flyttes, skal `SimulationEngine` regne ut den vektede effekten:

* **Formel:** $Gjennomsnittskost = \frac{\sum (Kostnad_{lokasjon} \times Volum_{lokasjon})}{Totalt\ Volum}$
* **Output i UI:** Vise simulert gjennomsnittskost mot baseline, samt beregne årlig total besparelse/merkostnad i kroner.