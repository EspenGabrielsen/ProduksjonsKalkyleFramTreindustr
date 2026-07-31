# Brukermanual - Produksjonsmodellen
## For produksjonsledere, økonomi og innkjøp

> **Formål:** Denne manualen forklarer hvilke tall du skal legge inn i Excel-arket,
> og hvordan du bruker Marimo web-appen til å simulere, analysere og eksportere rapporter.

---

## Innhold

1. [Hva er produksjonsmodellen?](#1-hva-er-produksjonsmodellen)
2. [Oversikt over arkene](#2-oversikt-over-arkene)
3. [Product Master - Vareregisteret](#3-product-master---vareregisteret)
4. [Locations - Fabrikker og lagre](#4-locations---fabrikker-og-lagre)
5. [Work Centers - Maskiner og arbeidsplasser](#5-work-centers---maskiner-og-arbeidsplasser)
6. [Operation Master - Standardoperasjoner](#6-operation-master---standardoperasjoner)
7. [Item Costs - Kostpriser](#7-item-costs---kostpriser)
8. [BOM - Stykklisten (hva består produktet av)](#8-bom---stykklisten)
9. [Routing - Produksjonsflyten](#9-routing---produksjonsflyten)
10. [By Product Rules - Biprodukter](#10-by-product-rules---biprodukter)
11. [Capacity Calendar - Kapasitetskalender](#11-capacity-calendar---kapasitetskalender)
12. [Production Scenario - Produksjonsscenario](#12-production-scenario---produksjonsscenario)
13. [Slik kommer du i gang](#13-slik-kommer-du-i-gang)
14. [Vanlige feil og tips](#14-vanlige-feil-og-tips)

---

## 1. Hva er produksjonsmodellen?

Produksjonsmodellen er et system for kostnadsberegning og simulering.
Den består av tre deler som jobber sammen:

1. **Excel-ark** — der du fyller inn alle data (produkter, priser, maskiner, tider)
2. **Marimo web-app** — der du laster opp Excel, simulerer "what-if"-scenarioer og eksporterer rapporter
3. **Python-beregningsmotor** — som regner ut standardkost for hvert produkt

Systemet gjør fire ting:
- **Beregner standardkost** per produkt (hva koster det å lage én enhet?)
- **Analyserer lønnsomhet** (hvilke produkter tjener vi penger på?)
- **Simulerer endringer** (hva skjer hvis råvareprisen går opp 10%? eller hvis vi reduserer svinnet?)
- **Eksporterer rapporter** til PDF (for ledelsen) og Excel (for videre analyse i Power BI)

### Hvem gjør hva?

| Rolle | Ansvarsområde | I Excel-arket | I Marimo web-appen |
|-------|---------------|--------------|-------------------|
| **Produksjonsleder** | Maskiner, operasjonstider, produksjonsflyt | Fyller inn i Work Centers, Routing, Capacity Calendar | Justerer parametere, kjører simulering, eksporterer PDF-rapport |
| **Innkjøp** | Råvarepriser, leverandørdata | Oppdaterer Unit Cost i Item Costs | Laster opp Excel, ser konsekvens av prisendringer i simulering |
| **Økonomi / Controller** | Timekostnader, biproduktverdi, produktregister | Fyller inn i Product Master, Item Costs, By Product Rules | Laster opp Excel, eksporterer rapporter til PDF og Excel |
| **Produksjonsteknikk** | Stykkliste, operasjonsrekkefølge | Fyller inn i BOM, Routing, Operation Master | Verifiserer data i "Datamodell (Innsyn)"-fanen |
| **IT / Superbruker** | Database, versjonshistorikk, feilsøking | - | Gjeninnlaster tidligere versjoner, overvåker endringslogg |

---

## 2. Oversikt over arkene

Excel-filen har **10 ark** som må fylles ut. Her er en kort forklaring:

| Ark | Hva det er | Hvem fyller ut |
|-----|------------|----------------|
| **Product Master** | Register over alle varer (råvarer, ferdigvarer, biprodukter) | Økonomi |
| **Locations** | Fabrikker og lagre | Produksjonsleder |
| **Work Centers** | Maskiner og arbeidsplasser med timekostnad | Produksjonsleder + Økonomi |
| **Operation Master** | Standardoperasjoner (oppdeling, hovling, pakking osv.) | Produksjonsteknikk |
| **Item Costs** | Kostpriser for råvarer og markedsverdi for biprodukter | Innkjøp + Økonomi |
| **BOM** | Stykkliste - hva består produktet av? | Produksjonsteknikk |
| **Routing** | Produksjonsflyt - hvilke operasjoner, i hvilken rekkefølge, hvor lang tid? | Produksjonsleder |
| **By Product Rules** | Biprodukter som oppstår (spon, flis, bark) og hva de er verdt | Økonomi |
| **Capacity Calendar** | Kapasitetskalender per arbeidssenter | Produksjonsleder |
| **Production Scenario** | Forhåndsdefinerte produksjonsscenarioer | Produksjonsleder + Økonomi |

Når Excel-arket er fylt ut, **laster du det opp i Marimo-appen** — da blir alle data tilgjengelige for simulering og analyse.

---

## 3. Product Master - Vareregisteret

**Dette arket er "telefonkatalogen" over alle varer.** Alt som finnes i
virksomheten må være registrert her.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Item No** | En unik kode for varen. Du bestemmer selv koden, men den må være unik. | `RM001` (råvare), `FG001` (ferdigvare), `BP001` (biprodukt) |
| **Description** | Navnet på varen slik alle kjenner den | `Skrulast 48x198`, `Utvendig Panel 21x95` |
| **Item Type** | Hva slags vare er dette? | `Raw Material` (råvare), `Semi Finished` (halvfabrikat), `Finished Good` (ferdigvare), `By Product` (biprodukt), `Trading Item` (handelsvare) |
| **Product Group** | Hvilken gruppe tilhører varen? | `Skrulast`, `Panel`, `Kledning`, `Spon` |
| **Base Unit of Measure** | Hva måler vi varen i? | `M3` (kubikkmeter), `LM` (løpemeter), `KG` (kilo), `PCS` (stykker) |
| **Active** | Er varen fortsatt i bruk? | `Ja` eller `Nei` |

### Viktig å huske

- **Alle** varer må være registrert: råvarer, ferdigvarer, biprodukter og
  handelsvarer.
- Hvis du legger til et nytt produkt, må du også huske å legge det inn i
  **Item Costs**, **BOM** og **Routing** (se lengre ned).
- Item No er koden som brukes i alle andre ark for å referere til varen.

---

## 4. Locations - Fabrikker og lagre

**Dette arket er en liste over hvor dere holder til.** Hvert arbeidssenter
(maskin) må være knyttet til en lokasjon.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Location Code** | En kort kode for stedet | `KOD` (Kodal), `SKI` (Skien) |
| **Location Name** | Navnet på stedet | `Kodal Fabrikk` |
| **Location Type** | Hva slags sted er dette? | `Factory` (fabrikk), `Warehouse` (lager) |
| **Active** | Er stedet aktivt? | `Ja` eller `Nei` |

### Tips

- Du trenger bare én rad per fabrikk/lager.
- Hvis dere bare har én fabrikk, holder det med én rad.

---

## 5. Work Centers - Maskiner og arbeidsplasser

**Dette er kanskje det viktigste arket for produksjonslederen.** Her
registrerer du alle maskiner og arbeidsplasser, og hvor mye det koster å
bruke dem per time.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem finner tallet? | Eksempel |
|---------|-------------------|---------------------|----------|
| **Work Center Code** | En kort kode for maskinen | Produksjonsleder | `HOVEDHOVEL` |
| **Description** | Hva heter maskinen? | Produksjonsleder | `Hovedhovel` |
| **Location Code** | Hvilken fabrikk står maskinen i? | Produksjonsleder | `KOD` |
| **Labor Cost per Hour** | Hva koster én time arbeid ved maskinen? (lønn + arbeidsgiveravgift + pensjon + feriepenger) | Økonomi | `550` |
| **Machine Cost per Hour** | Hva koster maskinen per time? (avskrivninger, service, leasing, vedlikehold, strøm) | Økonomi | `900` |
| **Overhead Cost per Hour** | Indirekte kostnader per time (produksjonsledelse, kvalitetskontroll, intern logistikk) | Økonomi | `150` |
| **Capacity Hours per Day** | Hvor mange timer per dag kan maskinen kjøre? | Produksjonsleder | `16` (to skift) |
| **Effective Capacity %** | Hvor stor andel av tiden er maskinen faktisk i produksjon? (trekker fra stopp, vedlikehold, feil) | Produksjonsleder | `85` (betyr 85%) |
| **Active** | Er maskinen i bruk? | Produksjonsleder | `Ja` |

### Hvordan finne timekostnaden?

```
Total kost per time = Labor + Machine + Overhead

Eksempel - Hovedhovel:
  550 kr (lønn) + 900 kr (maskin) + 150 kr (overhead) = 1 600 kr/time
```

**Tips til økonomi:**
- **Labor Cost**: Ta årslønn inkl. feriepenger, pensjon og arbeidsgiveravgift,
  del på 1950 timer (normal årsverk).
- **Machine Cost**: Årlige kostnader (avskrivning + service + strøm) delt på
  antall produksjonstimer per år.
- **Overhead**: Totale indirekte produksjonskostnader delt på totale
  maskintimer.

**Tips til produksjonsleder:**
- **Effective Capacity %**: Hvis maskinen er planlagt å kjøre 16 timer, men
  i snitt står 2,4 timer pga. vedlikehold, omstilling og feil, blir effektiv
  kapasitet (16-2,4)/16 = 85%.

---

## 6. Operation Master - Standardoperasjoner

**Dette er "ordboken" over hva slags operasjoner dere utfører.** Her lister
du opp alle typer operasjoner som finnes: oppdeling, hovling, profilering,
maling, pakking osv.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Operation Code** | En kort kode for operasjonen | `RIP` (oppdeling), `PLANING` (hovling) |
| **Description** | Hva heter operasjonen? | `Oppdeling`, `Hovling` |
| **Default Work Center** | Hvilken maskin brukes vanligvis? | `HOVEDHOVEL` |
| **Standard Unit** | Måles operasjonstiden i minutter eller timer? | `Minutes` |
| **Active** | Er operasjonen i bruk? | `Ja` |

### Tips

- Dette arket trenger du bare å fylle ut én gang. Det er en standardliste
  over hva dere gjør.
- Når du skal sette opp produksjonsflyten for et produkt (i Routing-arket),
  velger du fra denne listen.

---

## 7. Item Costs - Kostpriser

**Dette arket er for innkjøp og økonomi.** Her setter dere priser på råvarer
og markedsverdi på biprodukter.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem fyller ut? | Eksempel |
|---------|-------------------|-----------------|----------|
| **Item No** | Varekoden (samme som i Product Master) | - | `RM001` |
| **Cost Type** | Hva slags kostpris er dette? | Økonomi | `Standard Cost` (anbefalt) |
| **Unit Cost** | Pris per enhet | Innkjøp | `3000.00` (for RM001, prisen per M3) |
| **Currency** | Hvilken valuta? | Økonomi | `NOK` |
| **Effective Date** | Fra hvilken dato gjelder prisen? | Innkjøp | `2026-01-01` |

### Hvem fyller ut hva?

| Varetype | Hvem setter prisen? | Forklaring |
|----------|---------------------|------------|
| **Råvarer (RM001, RM002...)** | **Innkjøp** | Sett inn faktisk innkjøpspris. Eksempel: Skrulast koster 3 000 kr per M3 |
| **Maling (RM003)** | **Innkjøp** | Pris per liter. Eksempel: 120 kr per liter |
| **Ferdigvarer (FG001, FG002...)** | **Skal være 0** | Kostnaden beregnes automatisk av modellen. La stå som 0. |
| **Biprodukter (BP001, BP002...)** | **Økonomi** | Hva kan dere selge biproduktet for? Eksempel: Hovelspon 1,50 kr/kg |

### Viktig

- **Råvarer**: Sett inn den prisen dere faktisk betaler. Oppdater når prisen
  endrer seg.
- **Ferdigvarer**: Skal ALLTID ha 0 i Unit Cost. Modellen regner ut
  kostnaden automatisk basert på hva produktet består av (BOM) og
  produksjonsprosessen (Routing).
- **Biprodukter**: Sett inn markedsverdi - hva kan dere selge det for?

---

## 8. BOM - Stykklisten

**Dette arket forteller modellen hva hvert produkt består av.** For
produksjonsteknikeren: dette er den samme stykklisten dere kjenner fra
ERP-systemet.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Parent Item No** | Produktet du skal lage | `FG001` (Utvendig Panel) |
| **Component Item No** | Hva trenger du for å lage det? | `RM001` (Skrulast) |
| **Quantity Per** | Hvor mange enheter får du ut av én enhet inn? | `400` (betyr: 400 LM panel per M3 skrulast) |
| **Unit of Measure** | Måleenhet | `LM` |
| **Scrap %** | Hvor mye går til spille? (svinn) | `5.0` (betyr 5% svinn) |
| **Co-Prod %** | Andel av produksjonen som blir samprodukt (co-product). F.eks. 6% B-vare | `6.0` |
| **Co-Prod Item No** | Varenummer for samproduktet (f.eks. B-vare) | `JD16073-B` |
| **Valid From** | Fra hvilken dato gjelder dette? | `2026-01-01` |
| **Valid To** | Til hvilken dato gjelder dette? | `2026-12-31` |

### Co-Produkt (samprodukt / A- og B-vare)

Co-Prod % brukes når en andel av produksjonen blir et sekundært produkt
(f.eks. B-vare). B-varen har samme materialkost per enhet som A-varen, men
får kun en forholdsmessig andel av operasjonskostnaden allokert.

```
Eksempel - Vare JD16073 (A-vare) med 0,5% B-vare (JD16073-B):
  Teoretisk utbytte: 533,34 meter per M3
  Co-Prod: 0,5%
  A-vare kvantum: 533,34 x (1 - 0,005) = 530,67 LM
  B-vare kvantum: 533,34 x 0,005 = 2,67 LM
  Materialkost per LM: samme for A og B
  Operasjonskost: B får 0,5% av A-varens operasjonskost
```

### Slik fungerer Quantity Per

```
Quantity Per = hvor mye du får ut av én inn-enhet

Eksempel - Panel (FG001) fra Skrulast (RM001):
  Quantity Per = 400 LM per M3
  Det betyr: 1 M3 skrulast gir 400 LM ferdig panel
  
  Forbruk per LM panel = 1 / 400 = 0,0025 M3 per LM
```

### Eksempel på stykkliste

| Produkt | Består av | Hvor mye? |
|---------|-----------|-----------|
| FG001 - Utvendig Panel | RM001 - Skrulast | 400 LM per M3 |
| FG002 - Terrassebord | RM002 - Gran | 250 LM per M3 |
| FG003 - Kledning | RM001 - Skrulast | 420 LM per M3 |
| FG004 - Malt Panel | FG001 - Ubehandlet panel + RM003 - Maling | 1 LM + 0,05 LTR |

### Tips

- **FG004 (malt panel)** er et eksempel på en **produksjonskjede**: Det
  bruker FG001 (ferdig panel) som komponent. Modellen forstår dette og
  beregner kostnaden riktig.
- **Scrap %** er viktig for realistiske kostnader. Hvis 5% av materialet går
  til spille, må du kjøpe inn 5% mer.

---

## 9. Routing - Produksjonsflyten

**Dette arket er for produksjonslederen.** Her beskriver du nøyaktig hvordan
hvert produkt blir produsert: hvilke operasjoner, i hvilken rekkefølge, på
hvilken maskin, og hvor lang tid det tar.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem finner tallet? | Eksempel |
|---------|-------------------|---------------------|----------|
| **Item No** | Produktet som skal produseres | - | `FG001` |
| **Operation No** | Rekkefølgen (10, 20, 30...) | Produksjonsleder | `10` |
| **Operation Code** | Hva skal gjøres? (fra Operation Master) | Produksjonsleder | `RIP` (oppdeling) |
| **Work Center Code** | Hvilken maskin? (fra Work Centers) | Produksjonsleder | `HOVEDHOVEL` |
| **Setup Time Minutes** | Hvor lang tid tar det å rigge til? (omstilling, knivbytte, innkjøring) | Produksjonsleder | `15.0` |
| **Run Time Minutes** | Hvor lang tid tar det å produsere ÉN enhet? | Produksjonsleder | `0.15` (minutter per LM) |
| **Batch Size** | Hvor mange enheter lager dere per ordre? | Produksjonsleder | `500` |
| **Valid From** | Fra hvilken dato gjelder denne routingen? | Produksjonsleder | `2026-01-01` |
| **Valid To** | Til hvilken dato gjelder denne routingen? | Produksjonsleder | `2026-12-31` |

> **Merk:** `Changeover Time Minutes` er fjernet fra datamodellen.
> Omstillingskost håndteres gjennom `Setup Time Minutes`.

### Slik finner du tidene

**Run Time Minutes** er den viktigste kolonnen. Den sier hvor lang tid
maskinen bruker på å produsere **én enhet**.

```
Eksempel - Oppdeling av panel (FG001):
  Maskinen kjører 400 LM per M3
  Hastighet: ca. 6,7 LM per minutt
  Run Time = 1 / 6,7 = 0,15 minutter per LM
```

**Setup Time Minutes** er tiden det tar å klargjøre maskinen FØR
produksjonen starter. Dette inkluderer:
- Bytte kniver/profiler
- Justere maskinen
- Kjøre inn og kontrollmåle
- Rydde opp etter forrige ordre

**Batch Size** er hvor mange enheter dere vanligvis produserer per ordre.
Denne brukes til å fordele setup-kostnaden.

### Eksempel på produksjonsflyt

```
FG001 - Utvendig Panel 21x95:
  Op 10: Oppdeling  @ Hovedhovel  (15 min oppsett + 0,15 min per LM)
  Op 20: Hovling    @ Hovedhovel  (10 min oppsett + 0,10 min per LM)
  Op 30: Profilering @ Spesialhovel (20 min oppsett + 0,12 min per LM)
  Op 40: Pakking    @ Pakkelinje   (5 min oppsett + 0,05 min per LM)
```

### Viktig for FG004 (malt panel)

FG004 er malt panel. Det betyr at **ubehandlet panel (FG001) allerede er
ferdig produsert** med oppdeling, hovling, profilering og pakking. FG004
trenger derfor bare:

```
FG004 - Utvendig Panel 21x95 - Malt:
  Op 10: Maling     @ Malingslinje (30 min oppsett + 0,20 min per LM)
  Op 20: Pakking    @ Pakkelinje   (5 min oppsett + 0,05 min per LM)
```

---

## 10. By Product Rules - Biprodukter

**Dette arket er for økonomi.** I trelastproduksjon oppstår det alltid
biprodukter som spon, flis og bark. Disse har en verdi som skal trekkes fra
produksjonskostnaden.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Parent Item No** | Hvilket produkt skaper biproduktet? | `FG001` |
| **By Product Item No** | Hvilket biprodukt? | `BP001` (Hovelspon) |
| **Expected Quantity** | Hvor mye biprodukt per enhet hovedprodukt? | `0.5` (0,5 kg spon per LM panel) |
| **Unit of Measure** | Måleenhet | `KG` |
| **Market Value** | Hva kan dere selge biproduktet for? | `1.50` (kr per kg) |
| **Allocation Method** | Hvordan skal verdien brukes? | `Reduce Main Product Cost` (anbefalt) |

### Slik fungerer det

```
Biproduktverdi per enhet = Forventet mengde x Markedsverdi

Eksempel - FG001 (Panel):
  Hovelspon: 0,5 kg x 1,50 kr/kg = 0,75 kr per LM
  Flis:      0,3 kg x 0,80 kr/kg = 0,24 kr per LM
  Total:     0,99 kr per LM

  Denne verdien trekkes FRA produktets kostnad:
  Netto kost = Brutto kost - Biproduktverdi
```

### Hvem gjør hva?

| Oppgave | Ansvarlig |
|---------|-----------|
| Anslå hvor mye biprodukt som oppstår | Produksjonsleder |
| Sett markedsverdi (hva kan dere selge det for?) | Økonomi/Innkjøp |

---

## 11. Capacity Calendar - Kapasitetskalender

**Dette arket er for produksjonslederen.** Her registrerer du tilgjengelig
kapasitet per arbeidssenter per dag. Kalenderen brukes til å analysere
flaskehalser og planlegge produksjon.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem fyller ut? | Eksempel |
|---------|-------------------|-----------------|----------|
| **Work Center** | Arbeidssenteret (samme som i Work Centers) | Produksjonsleder | `HOVEDHOVEL` |
| **Date** | Dato | Produksjonsleder | `2026-01-05` |
| **Available Hours** | Tilgjengelige timer denne dagen | Produksjonsleder | `16` |
| **Planned Downtime** | Planlagte stopp (vedlikehold, ferie, ombygging) | Produksjonsleder | `0` |

### Beregning

```
Available Production Hours = Available Hours - Planned Downtime
```

---

## 12. Production Scenario - Produksjonsscenario

**Dette arket er for produksjonslederen og økonomi.** Her definerer du
forhåndsdefinerte produksjonsscenarioer med planlagt kvantum per produkt.
Scenarioene brukes i Marimo-appen til å simulere produksjon og beregne
totalt ressursbehov.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem fyller ut? | Eksempel |
|---------|-------------------|-----------------|----------|
| **Scenario Name** | Navn på scenario | Produksjonsleder | `Normal Produksjon` |
| **Product** | Produkt som skal produseres (Item No) | Produksjonsleder | `FG001` |
| **Planned Quantity** | Planlagt antall enheter | Produksjonsleder | `50000` |
| **Start Date** | Startdato for produksjon | Produksjonsleder | `2026-01-01` |
| **End Date** | Sluttdato for produksjon | Produksjonsleder | `2026-12-31` |

### Tips

- Scenarioer brukes i Marimo-appen til å simulere "what-if" analyser.
- Du kan ha flere produkter per scenario (én rad per produkt).
- Kvantumet påvirker hvor mye setupkost som fordeles per enhet.

---

## 13. Slik kommer du i gang

### Første gang — oppsett

1. **Start med testdataene** som følger med. Åpne
   `Produksjonsmodell_Testdata_v3.xlsx` for å se hvordan et ferdig oppsett
   ser ut.
2. **Erstatt testdataene** med dine egne produkter og priser.
3. **Start Marimo web-appen:**
   ```bash
   marimo run varekost_app.py
   ```
4. **Last opp Excel-filen din:**
   - Gå til fanen **📁 Dataimport & Versjoner**
   - Klikk **📄 Velg Excel-fil** og velg din .xlsx-fil
   - Skriv en kommentar (f.eks. "Første import av egne data")
   - Systemet validerer automatisk og importerer
5. **Simulér og analyser:**
   - Gå til fanen **📊 Simulering & Analyse**
   - Juster parametere (råvarepriser, svinn, timekostnader, produksjonstider)
   - Angi planlagt kvantum
   - Trykk **⚡ Start simulering**
6. **Eksporter rapport:**
   - Under **💾 Eksport**: last ned PDF-rapport eller Excel-fil

### Fremgangsmåte for å legge til et nytt produkt

| Steg | Ark | Hva skal gjøres? | Hvem? |
|------|-----|------------------|-------|
| 1 | **Product Master** | Legg til varen med Item No, navn og type | Økonomi |
| 2 | **Item Costs** | Sett pris (0 for ferdigvarer) | Innkjøp |
| 3 | **BOM** | Hva består produktet av? (stykkliste) | Produksjonsteknikk |
| 4 | **Routing** | Hvordan produseres det? (operasjoner, tider) | Produksjonsleder |
| 5 | **By Product Rules** | Oppstår det biprodukter? | Økonomi |
| 6 | **Marimo-app** | Last opp Excel-filen på nytt | Hvem som helst |

### Oppdatere priser

Når priser endrer seg (f.eks. ny innkjøpspris på skrulast):

| Hva skal oppdateres? | Ark | Hvem? |
|----------------------|-----|-------|
| Ny råvarepris | **Item Costs** | Innkjøp |
| Ny timekostnad på maskin | **Work Centers** | Økonomi |
| Ny markedsverdi på biprodukt | **Item Costs** | Økonomi |

Etter endringene: last opp Excel-filen på nytt i Marimo-appen for å få oppdaterte data.

### Versjonshistorikk

Marimo-appen lagrer alle tidligere Excel-filer du har lastet opp.
Gå til **📁 Dataimport & Versjoner** og velg en tidligere versjon for å
gjeninnlaste den. Dette er nyttig hvis du har gjort feil og vil gå tilbake.

---

## 14. Vanlige feil og tips

### ❌ Vanlige feil i Excel-arket

| Feil | Problem | Løsning |
|------|---------|---------|
| **Manglende vare** | Produktet finnes ikke i Product Master | Legg til varen i alle ark |
| **Feil Item No** | Skrivefeil i varekode | Sjekk at koden er nøyaktig lik i alle ark |
| **Unit Cost = 0 på råvare** | Modellen tror råvaren er gratis | Sett inn faktisk pris i Item Costs |
| **Unit Cost > 0 på ferdigvare** | Modellen dobler kostnaden | Sett 0 på ferdigvarer - kostnad beregnes automatisk |
| **Manglende BOM** | Produktet har ingen stykkliste | Legg til BOM-linjer for produktet |
| **Manglende Routing** | Produktet har ingen produksjonsflyt | Legg til operasjoner i Routing |
| **Feil Quantity Per** | Forbruket blir feil | Sjekk: Quantity Per = output per input. Hvis 1 M3 gir 400 LM, skriv 400 |
| **Scrap % for høy/lav** | Materialkost blir feil | Sjekk faktisk svinn i produksjonen |

### ❌ Vanlige feil ved import i Marimo

| Feilmelding | Årsak | Løsning |
|-------------|-------|---------|
| "Validering fant feil — ingenting importert" | Excel-filen mangler ark eller kolonner | Sjekk at filen har alle 10 ark med korrekte kolonnenavn |
| "Kryssreferanse-feil" | En varekode i BOM finnes ikke i Product Master | Sjekk at alle Item No er registrert |
| "Ingen data lastet" | Databasen er tom | Last opp en Excel-fil via "Dataimport & Versjoner" |

### ⚠️ Kjente begrensninger i dagens kalkyle

Følgende felt registreres i Excel-arket, men filtreres foreløpig ikke i
Python-beregningen:

| Felt | Status | Planlagt forbedring |
|------|--------|---------------------|
| **Active** (Product/Location/WC/Operation) | Visuell info kun — alle regnes som aktive | Filtrering kommer |
| **Valid From** / **Valid To** (BOM og Routing) | Ignoreres — alle linjer inkluderes alltid | Datofiltrering kommer |
| **Effective Date** (Item Costs) | Velger nyeste dato, ikke "gyldig per i dag" | Forbedres til å bruke en valgt analysedato |
| **Start Date** / **End Date** (Scenario) | Ignoreres i simulering | Planlegges |

Dette påvirker ikke standard bruk av modellen, men vær oppmerksom på det
hvis du har inaktive produkter eller tidsbegrensede priser i datasettet ditt.

### ✅ Gode råd

1. **Start enkelt.** Legg inn 2-3 produkter først, sjekk at tallene gir
   mening, så utvider du.
2. **Sjekk at summen stemmer.** Beregn for hånd et enkelt produkt og
   sammenlign med modellens resultat i Marimo.
3. **Oppdater jevnlig.** Priser endrer seg — sett av tid til å oppdatere
   modellen hvert kvartal.
4. **Bruk kommentarfeltet.** Når du laster opp Excel i Marimo, skriv hva
   som er endret — da kan du senere se i endringsloggen hva som skjedde.
5. **Bruk versjonshistorikken.** Hvis noe går galt, kan du alltid gå
   tilbake til en tidligere versjon.
6. **Spør om hjelp.** Hvis tallene ser rare ut, sjekk om alle arkene er
   fylt ut riktig.

---

> **Trenger du hjelp?** Se `Produksjonsmodell_Dokumentasjon.pdf` for
> teknisk dokumentasjon, eller kontakt systemansvarlig.
>
> *Sist oppdatert: juli 2026*