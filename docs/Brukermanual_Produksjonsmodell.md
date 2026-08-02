# Brukermanual - Produksjonsmodellen

## For produksjonsledere, økonomi og innkjøp

> **Formål:** Denne manualen forklarer hvilke tall du skal legge inn i Excel-arket,
> og hvordan du bruker Marimo web-appen til å simulere, analysere og eksportere rapporter.

---

## Innhold

1. Hva er produksjonsmodellen?
2. Oversikt over arkene
3. Produktregister - Vareregisteret
4. Lokasjoner - Fabrikker og lagre
5. Arbeidssentre - Maskiner og arbeidsplasser
6. Operasjonsregister - Standardoperasjoner
7. Varekostnader - Kostpriser
8. Stykkliste - hva består produktet av
9. Produksjonsrute - Produksjonsflyten
10. Biproduktregler - Biprodukter
11. Transportflagg - Hvilke varer transporteres
12. Transportruter - Fraktkost mellom lokasjoner
13. Kapasitetskalender - Kapasitetskalender
14. Produksjonsscenario - Produksjonsscenario
15. Slik kommer du i gang
16. Vanlige feil og tips

---

<a id="1-hva-er-produksjonsmodellen"></a>

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
| **Produksjonsleder** | Maskiner, operasjonstider, produksjonsflyt | Fyller inn i Arbeidssentre, Produksjonsrute, Kapasitetskalender | Justerer parametere, kjører simulering, eksporterer PDF-rapport |
| **Innkjøp** | Råvarepriser, leverandørdata | Oppdaterer enhetskost i Varekostnader | Laster opp Excel, ser konsekvens av prisendringer i simulering |
| **Økonomi / Controller** | Timekostnader, biproduktverdi, produktregister | Fyller inn i Produktregister, Varekostnader, Biproduktregler | Laster opp Excel, eksporterer rapporter til PDF og Excel |
| **Produksjonsteknikk** | Stykkliste, operasjonsrekkefølge | Fyller inn i Stykkliste, Produksjonsrute, Operasjonsregister | Verifiserer data i "Datamodell (Innsyn)"-fanen |
| **Logistikk / Drift** | Transport mellom høvlerier, fraktkost | Fyller inn Is Transport i Produktregister + Transportruter | Justerer fraktkost i simuleringen |
| **IT / Superbruker** | Database, versjonshistorikk, feilsøking | - | Gjeninnlaster tidligere versjoner, overvåker endringslogg |

---

<a id="2-oversikt-over-arkene"></a>

## 2. Oversikt over arkene

Excel-filen har **10 stamdata-ark + 1 transport-ark** som fylles ut. Her er en kort forklaring:

| Ark | Hva det er | Hvem fyller ut |
|-----|------------|----------------|
| **Produktregister** | Register over alle varer (råvarer, ferdigvarer, biprodukter). Inkluderer kolonnen **Is Transport** som markerer transportvarer | Økonomi |
| **Lokasjoner** | Fabrikker og lagre | Produksjonsleder |
| **Arbeidssentre** | Maskiner og arbeidsplasser med timekostnad | Produksjonsleder + Økonomi |
| **Operasjonsregister** | Standardoperasjoner (oppdeling, hovling, pakking osv.) | Produksjonsteknikk |
| **Varekostnader** | Kostpriser for råvarer og markedsverdi for biprodukter | Innkjøp + Økonomi |
| **Stykkliste** | Stykkliste - hva består produktet av? | Produksjonsteknikk |
| **Produksjonsrute** | Produksjonsflyt - hvilke operasjoner, i hvilken rekkefølge, hvor lang tid? | Produksjonsleder |
| **Biproduktregler** | Biprodukter som oppstår (spon, flis, bark) og hva de er verdt | Økonomi |
| **Kapasitetskalender** | Kapasitetskalender per arbeidssenter *(legacy)* | Produksjonsleder |
| **Produksjonsscenario** | Forhåndsdefinerte produksjonsscenarioer *(legacy)* | Produksjonsleder + Økonomi |
| **Transportruter** | Fraktkost per M3 mellom lokasjoner (f.eks. Kodal → Skien) | Logistikk / Økonomi |

Når Excel-arket er fylt ut, **laster du det opp i Marimo-appen** — da blir alle data tilgjengelige for simulering og analyse.

**Endringslogg:** I tillegg genererer appen automatisk et eget ark med **Endringslogg** når du laster ned den komplette datafilen. Dette arket fyller du ikke ut selv — det viser historikken over alle endringer i databasen.

---

<a id="3-produktregister---vareregisteret"></a>

## 3. Produktregister - Vareregisteret

**Dette arket er "telefonkatalogen" over alle varer.** Alt som finnes i
virksomheten må være registrert her.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Item No** | En unik kode for varen. Du bestemmer selv koden, men den må være unik. | Du kan f.eks. bruke prefiks for varetype (RM for råvare, FG for ferdigvare) |
| **Description** | Navnet på varen slik alle kjenner den | `Skrulast 48x198`, `Utvendig Panel 21x95` |
| **Item Type** | Hva slags vare er dette? | `Raw Material` (råvare), `Semi Finished` (halvfabrikat), `Finished Good` (ferdigvare), `By Product` (biprodukt), `Trading Item` (handelsvare) |
| **Product Group** | Hvilken gruppe tilhører varen? | `Skrulast`, `Panel`, `Kledning`, `Spon` |
| **Base Unit of Measure** | Hva måler vi varen i? | `M3` (kubikkmeter), `LM` (løpemeter), `KG` (kilo), `PCS` (stykker) |
| **Is Transport** | Er varen en transportvare som transporteres mellom høvlerier? (se kapittel 11) | `1` (ja) eller `0` (nei) |

### Viktig å huske

- **Alle** varer må være registrert: råvarer, ferdigvarer, biprodukter og handelsvarer.
- Hvis du legger til et nytt produkt, må du også huske å legge det inn i
  **Varekostnader**, **Stykkliste** og **Produksjonsrute** (se lengre ned).
- Item No er koden som brukes i alle andre ark for å referere til varen.

---

<a id="4-lokasjoner---fabrikker-og-lagre"></a>

## 4. Lokasjoner - Fabrikker og lagre

**Dette arket er en liste over hvor dere holder til.** Hvert arbeidssenter
(maskin) må være knyttet til en lokasjon.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Location Code** | En kort kode for stedet | `KOD` (Kodal), `SKI` (Skien) |
| **Location Name** | Navnet på stedet | `Kodal Fabrikk` |
| **Location Type** | Hva slags sted er dette? | `Factory` (fabrikk), `Warehouse` (lager) |

### Tips

- Du trenger bare én rad per fabrikk/lager.
- Hvis dere bare har én fabrikk, holder det med én rad.

---

<a id="5-arbeidssentre---maskiner-og-arbeidsplasser"></a>

## 5. Arbeidssentre - Maskiner og arbeidsplasser

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
- **Overhead**: Totale indirekte produksjonskostnader delt på totale maskintimer.

**Tips til produksjonsleder:**

- **Effective Capacity %**: Hvis maskinen er planlagt å kjøre 16 timer, men
  i snitt står 2,4 timer pga. vedlikehold og feil, blir effektiv
  kapasitet (16-2,4)/16 = 85%.

---

<a id="6-operasjonsregister---standardoperasjoner"></a>

## 6. Operasjonsregister - Standardoperasjoner

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

### Tips

- Dette arket trenger du bare å fylle ut én gang. Det er en standardliste over hva dere gjør.
- Når du skal sette opp produksjonsflyten for et produkt (i Produksjonsrute-arket),
  velger du fra denne listen.

---

<a id="7-varekostnader---kostpriser"></a>

## 7. Varekostnader - Kostpriser

**Dette arket er for innkjøp og økonomi.** Her setter dere priser på råvarer
og markedsverdi på biprodukter.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem fyller ut? | Eksempel |
|---------|-------------------|-----------------|----------|
| **Item No** | Varekoden (samme som i Produktregister) | - | Varekoden fra Produktregister |
| **Cost Type** | Hva slags kostpris er dette? | Økonomi | `Standard Cost` (anbefalt) |
| **Unit Cost** | Pris per enhet | Innkjøp | `3000.00` (pris per M3) |
| **Currency** | Hvilken valuta? | Økonomi | `NOK` |

### Hvem fyller ut hva?

| Varetype | Hvem setter prisen? | Forklaring |
|----------|---------------------|------------|
| **Råvarer** | **Innkjøp** | Sett inn faktisk innkjøpspris. Eksempel: Skrulast koster 3 000 kr per M3 |
| **Maling** | **Innkjøp** | Pris per liter. Eksempel: 120 kr per liter |
| **Ferdigvarer** | **Skal være 0** | Kostnaden beregnes automatisk av modellen. La stå som 0. |
| **Biprodukter** | **Økonomi** | Hva kan dere selge biproduktet for? Eksempel: Hovelspon 1,50 kr/kg |

### Viktig

- **Råvarer**: Sett inn den prisen dere faktisk betaler. Oppdater når prisen endrer seg.
- **Ferdigvarer**: Skal ALLTID ha 0 i Unit Cost. Modellen regner ut kostnaden
  automatisk basert på hva produktet består av (Stykkliste) og
  produksjonsprosessen (Produksjonsrute).
- **Biprodukter**: Sett inn markedsverdi - hva kan dere selge det for?
- **Transportvarer**: For varer som transporteres mellom høvlerier, legges
  fraktkostnaden til automatisk i simuleringen basert på Transportruter
  (se kapittel 12).

---

<a id="8-stykkliste---hva-består-produktet-av"></a>

## 8. Stykkliste - hva består produktet av

**Dette arket forteller modellen hva hvert produkt består av.** For
produksjonsteknikeren: dette er den samme stykklisten dere kjenner fra
ERP-systemet.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Parent Item No** | Produktet du skal lage | Varekoden til ferdigvaren |
| **Component Item No** | Hva trenger du for å lage det? | Varekoden til råvaren |
| **Quantity Per** | Hvor mange enheter får du ut av én enhet inn? | `400` (betyr: 400 LM panel per M3 skrulast) |
| **Unit of Measure** | Måleenhet | `LM` |
| **Scrap %** | Hvor mye går til spille? (svinn) | `5.0` (betyr 5% svinn) |
| **Co-Prod %** | Andel av produksjonen som blir samprodukt (co-product). F.eks. 6% B-vare | `6.0` |
| **Co-Prod Item No** | Varenummer for samproduktet (f.eks. B-vare) | `B`-suffiks på varekoden |

### Co-Produkt (samprodukt / A- og B-vare)

Co-Prod % brukes når en andel av produksjonen blir et sekundært produkt
(f.eks. B-vare). B-varen har samme materialkost per enhet som A-varen, men
får kun en forholdsmessig andel av operasjonskostnaden allokert.

```
Eksempel - Hovedprodukt (A-vare) med 0,5% B-vare:
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

Eksempel - Utvendig Panel (ferdigvare) fra Skrulast (råvare):
  Quantity Per = 400 LM per M3
  Det betyr: 1 M3 skrulast gir 400 LM ferdig panel

  Forbruk per LM panel = 1 / 400 = 0,0025 M3 per LM
```

### Eksempel på stykkliste

| Produkt | Består av | Hvor mye? |
|---------|-----------|-----------|
| Utvendig Panel (ferdigvare) | Skrulast (råvare) | 400 LM per M3 |
| Terrassebord (ferdigvare) | Gran (råvare) | 250 LM per M3 |
| Kledning (ferdigvare) | Skrulast (råvare) | 420 LM per M3 |
| Malt Panel (ferdigvare) | Ubehandlet panel + Maling | 1 LM + 0,05 LTR |

### Tips

- **Malt panel** er et eksempel på en **produksjonskjede**: Det bruker
  ubehandlet panel (ferdigvare) som komponent. Modellen forstår dette og
  beregner kostnaden riktig.
- **Scrap %** er viktig for realistiske kostnader. Hvis 5% av materialet går
  til spille, må du kjøpe inn 5% mer.

---

<a id="9-produksjonsrute---produksjonsflyten"></a>

## 9. Produksjonsrute - Produksjonsflyten

**Dette arket er for produksjonslederen.** Her beskriver du nøyaktig hvordan
hvert produkt blir produsert: hvilke operasjoner, i hvilken rekkefølge, på
hvilken maskin, og hvor lang tid det tar.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem finner tallet? | Eksempel |
|---------|-------------------|---------------------|----------|
| **Item No** | Produktet som skal produseres | - | Varekoden til produktet |
| **Operation No** | Rekkefølgen (10, 20, 30...) | Produksjonsleder | `10` |
| **Operation Code** | Hva skal gjøres? (fra Operasjonsregister) | Produksjonsleder | `RIP` (oppdeling) |
| **Work Center Code** | Hvilken maskin? (fra Arbeidssentre) | Produksjonsleder | `HOVEDHOVEL` |
| **Setup Time Minutes** | Hvor lang tid tar det å rigge til? (omstilling, knivbytte, innkjøring) | Produksjonsleder | `15.0` |
| **Run Time Minutes** | Hvor lang tid tar det å produsere ÉN enhet? | Produksjonsleder | `0.15` (minutter per LM) |
| **Batch Size** | Hvor mange enheter lager dere per ordre? | Produksjonsleder | `500` |

> **Merk:** `Changeover Time Minutes` er fjernet fra datamodellen.
> Omstillingskost håndteres gjennom `Setup Time Minutes`.

### Slik finner du tidene

**Run Time Minutes** er den viktigste kolonnen. Den sier hvor lang tid
maskinen bruker på å produsere **én enhet**.

```
Eksempel - Oppdeling av utvendig panel:
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
Utvendig Panel 21x95:
  Op 10: Oppdeling  @ Hovedhovel  (15 min oppsett + 0,15 min per LM)
  Op 20: Hovling    @ Hovedhovel  (10 min oppsett + 0,10 min per LM)
  Op 30: Profilering @ Spesialhovel (20 min oppsett + 0,12 min per LM)
  Op 40: Pakking    @ Pakkelinje   (5 min oppsett + 0,05 min per LM)
```

### Viktig for malt panel

Malt panel betyr at **ubehandlet panel (ferdigvare) allerede er ferdig produsert**
med oppdeling, hovling, profilering og pakking. Malt panel trenger derfor bare:

```
Malt Utvendig Panel 21x95:
  Op 10: Maling     @ Malingslinje (30 min oppsett + 0,20 min per LM)
  Op 20: Pakking    @ Pakkelinje   (5 min oppsett + 0,05 min per LM)
```

---

<a id="10-biproduktregler---biprodukter"></a>

## 10. Biproduktregler - Biprodukter

**Dette arket er for økonomi.** I trelastproduksjon oppstår det alltid
biprodukter som spon, flis og bark. Disse har en verdi som skal trekkes fra
produksjonskostnaden.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Parent Item No** | Hvilket produkt skaper biproduktet? | Varekoden til produktet |
| **By Product Item No** | Hvilket biprodukt? | Varekoden til biproduktet (f.eks. som ender på `BP`) |
| **Expected Quantity** | Hvor mye biprodukt per enhet hovedprodukt? | `0.5` (0,5 kg spon per LM panel) |
| **Unit of Measure** | Måleenhet | `KG` |
| **Market Value** | Hva kan dere selge biproduktet for? | `1.50` (kr per kg) |
| **Allocation Method** | Hvordan skal verdien brukes? | `Reduce Main Product Cost` (anbefalt) |

### Slik fungerer det

```
Biproduktverdi per enhet = Forventet mengde x Markedsverdi

Eksempel - Utvendig Panel:
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

<a id="11-transportflagg---hvilke-varer-transporteres"></a>

## 11. Transportflagg - Hvilke varer transporteres

I fler-høvleri-produksjon kan en vare produseres på ett høvleri (f.eks.
Kodal) og transporteres til et annet (f.eks. Skien) for videre produksjon
eller distribusjon. Slike varer kalles **transportvarer**, og modellen legger
automatisk til fraktkostnaden i kalkylen.

> **Merk:** Transportflagg er **ikke et eget ark** i Excel. Markeringen gjøres
> via kolonnen **Is Transport** i **Produktregister**-arket (se kapittel 3).

### Kolonne du må fylle ut

| Kolonne | Hva skal stå her? | Eksempel |
|---------|-------------------|----------|
| **Is Transport** | `1` = varen transporteres mellom høvlerier, `0` = vanlig vare | `1` |

### Slik fungerer det

Når en vare er flagget som transportvare (`Is Transport = 1`):

1. Modellen ser opp hvilke transportruter som finnes (kapittel 12).
2. For hver relevant rute legges **fraktkost per M3** til produktets kostnad.
3. I datamodellen vises varen med en fiktiv lokasjon (f.eks. `KOD→KV`) som
   inkluderer fraktkostnaden i stedet for en fysisk operasjon.

### Hvem gjør hva?

| Oppgave | Ansvarlig |
|---------|-----------|
| Bestemme hvilke varer som transporteres | Produksjonsleder / Logistikk |
| Fylle inn `Is Transport` i Produktregister | Logistikk / IT |
| Fylle inn fraktkost per M3 i Transportruter | Økonomi / Logistikk |

---

<a id="12-transportruter---fraktkost-mellom-lokasjoner"></a>

## 12. Transportruter - Fraktkost mellom lokasjoner

**Dette arket er for logistikk og økonomi.** Her registrerer du fraktkostnaden
per M3 mellom to lokasjoner. Rutene brukes sammen med transportflaggene
(kapittel 11) for å legge transportkostnaden til kalkylen.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem fyller ut? | Eksempel |
|---------|-------------------|-----------------|----------|
| **From Loc** | Fra-lokasjon (samme Location Code som i Lokasjoner) | Logistikk | `KOD` |
| **To Loc** | Til-lokasjon | Logistikk | `KV` |
| **Cost Per M3** | Fraktpris per M3 på denne ruten (eneste beregningsfelt) | Økonomi | `120` |
| **Distance Km** | Distanse i kilometer (informasjon — påvirker ikke kostnaden) | Logistikk | `45` |
| **Hours** | Kjøretid i timer (informasjon — påvirker ikke kostnaden) | Logistikk | `1.5` |

### Slik fungerer det

```
Transportkost = Fraktkost per M3 x Volum som transporteres

Eksempel - Rute KOD → KV:
  Fraktkost per M3: 120 kr
  Avstand: 45 km
  Kjøretid: 1,5 timer

  Kun Cost Per M3 påvirker kalkylen.
  Distance Km og Hours er informasjon for logistikk-planlegging.
```

### I Marimo-appen

I fanen **📊 Simulering & Analyse** finner du transportrutene under
**🚛 Transport (kr/m³)** i simuleringsparametrene. Der kan du justere
fraktkostnaden per M3 for å simulere f.eks. høyere drivstoffpriser eller
ny transportavtale:

```
Eksempel - Simulering:
  KOD → KV: Org. kost 120 kr/m³ → Ny kost 150 kr/m³
  KOD → EIK: Org. kost 95 kr/m³  → Ny kost 110 kr/m³
```

### Hvem gjør hva?

| Oppgave | Ansvarlig |
|---------|-----------|
| Sette fraktkost per M3 | Økonomi |
| Fylle inn distanse og kjøretid | Logistikk |
| Oppdatere ved ny transportavtale / prisendring | Innkjøp / Logistikk |

---

<a id="13-kapasitetskalender---kapasitetskalender"></a>

## 13. Kapasitetskalender - Kapasitetskalender *(legacy)*

**Dette arket er for produksjonslederen.** Her registrerer du tilgjengelig
kapasitet per arbeidssenter per dag. Kalenderen brukes til å analysere
flaskehalser og planlegge produksjon.

> **Merk:** Kapasitetskalender er foreløpig **ikke i aktiv beregning**.
> Dataene vises i "Datamodell (Innsyn)"-fanen, men brukes ikke i
> kostnadsberegningen eller simuleringen ennå. Dette påvirker ikke standard
> bruk av modellen.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem fyller ut? | Eksempel |
|---------|-------------------|-----------------|----------|
| **Work Center** | Arbeidssenteret (samme som i Arbeidssentre) | Produksjonsleder | `HOVEDHOVEL` |
| **Date** | Dato | Produksjonsleder | `2026-01-05` |
| **Available Hours** | Tilgjengelige timer denne dagen | Produksjonsleder | `16` |
| **Planned Downtime** | Planlagte stopp (vedlikehold, ferie, ombygging) | Produksjonsleder | `0` |

### Beregning

```
Available Production Hours = Available Hours - Planned Downtime
```

---

<a id="14-produksjonsscenario---produksjonsscenario"></a>

## 14. Produksjonsscenario - Produksjonsscenario *(legacy)*

**Dette arket er for produksjonslederen og økonomi.** Her definerer du
forhåndsdefinerte produksjonsscenarioer med planlagt kvantum per produkt.
Scenarioene brukes i Marimo-appen til å simulere produksjon og beregne
totalt ressursbehov.

> **Merk:** Produksjonsscenario er foreløpig **ikke i aktiv bruk**.
> Planlagt kvantum styres nå direkte i Marimo-appen via **📦 Planlagt
> kvantum** i "📊 Simulering & Analyse"-fanen. Arket er beholdt i
> datamodellen for framtidig bruk.

### Kolonner du må fylle ut

| Kolonne | Hva skal stå her? | Hvem fyller ut? | Eksempel |
|---------|-------------------|-----------------|----------|
| **Scenario Name** | Navn på scenario | Produksjonsleder | `Normal Produksjon` |
| **Product** | Produkt som skal produseres (Item No) | Produksjonsleder | Varekoden til produktet |
| **Planned Quantity** | Planlagt antall enheter | Produksjonsleder | `50000` |

### Tips

- Du kan ha flere produkter per scenario (én rad per produkt).
- Kvantumet påvirker hvor mye setupkost som fordeles per enhet.

---

<a id="15-slik-kommer-du-i-gang"></a>

## 15. Slik kommer du i gang

### Første gang — oppsett

1. **Start med testdataene** som følger med. Åpne
   `src/Produksjonsmodell_Testdata_v3.xlsx` for å se hvordan et ferdig oppsett ser ut.
2. **Erstatt testdataene** med dine egne produkter og priser.
3. **Start Marimo web-appen:**
   ```bash
   marimo run src/varekost_app.py
   ```
4. **Last opp Excel-filen din:**
   - Gå til fanen **📁 Dataimport & Versjoner**
   - Klikk **📄 Velg Excel-fil** og velg din .xlsx-fil
   - Skriv en kommentar (f.eks. "Første import av egne data")
   - Systemet validerer automatisk og importerer
5. **Simulér og analyser:**
   - Gå til fanen **📊 Simulering & Analyse**
   - Juster parametere (råvarepriser, svinn, timekostnader, produksjonstider, transportkost)
   - Angi planlagt kvantum
   - Trykk **⚡ Start simulering**
6. **Eksporter rapport:**
   - Under **💾 Eksport**: last ned PDF-rapport eller Excel-fil

### Fremgangsmåte for å legge til et nytt produkt

| Steg | Ark | Hva skal gjøres? | Hvem? |
|------|-----|------------------|-------|
| 1 | **Produktregister** | Legg til varen med Item No, navn og type | Økonomi |
| 2 | **Varekostnader** | Sett pris (0 for ferdigvarer) | Innkjøp |
| 3 | **Stykkliste** | Hva består produktet av? (stykkliste) | Produksjonsteknikk |
| 4 | **Produksjonsrute** | Hvordan produseres det? (operasjoner, tider) | Produksjonsleder |
| 5 | **Biproduktregler** | Oppstår det biprodukter? | Økonomi |
| 6 | **Produktregister** | Er varen en transportvare? Sett `Is Transport = 1` | Logistikk |
| 7 | **Transportruter** | Finnes fraktkost for aktuell rute? (ellers legg til) | Logistikk / Økonomi |
| 8 | **Marimo-app** | Last opp Excel-filen på nytt | Hvem som helst |

### Oppdatere priser

Når priser endrer seg (f.eks. ny innkjøpspris på skrulast):

| Hva skal oppdateres? | Ark | Hvem? |
|----------------------|-----|-------|
| Ny råvarepris | **Varekostnader** | Innkjøp |
| Ny timekostnad på maskin | **Arbeidssentre** | Økonomi |
| Ny markedsverdi på biprodukt | **Varekostnader** | Økonomi |
| Ny fraktkost / transportavtale | **Transportruter** | Innkjøp / Logistikk |

Etter endringene: last opp Excel-filen på nytt i Marimo-appen for å få oppdaterte data.

### Versjonshistorikk

Marimo-appen lagrer alle tidligere Excel-filer du har lastet opp.
Gå til **📁 Dataimport & Versjoner** og velg en tidligere versjon for å
gjeninnlaste den. Dette er nyttig hvis du har gjort feil og vil gå tilbake.

---

<a id="16-vanlige-feil-og-tips"></a>

## 16. Vanlige feil og tips

### ❌ Vanlige feil i Excel-arket

| Feil | Problem | Løsning |
|------|---------|---------|
| **Manglende vare** | Produktet finnes ikke i Produktregister | Legg til varen i alle ark |
| **Feil Item No** | Skrivefeil i varekode | Sjekk at koden er nøyaktig lik i alle ark |
| **Unit Cost = 0 på råvare** | Modellen tror råvaren er gratis | Sett inn faktisk pris i Varekostnader |
| **Unit Cost > 0 på ferdigvare** | Modellen dobler kostnaden | Sett 0 på ferdigvarer - kostnad beregnes automatisk |
| **Manglende stykkliste** | Produktet har ingen stykkliste | Legg til stykklistelinjer for produktet |
| **Manglende produksjonsrute** | Produktet har ingen produksjonsflyt | Legg til operasjoner i Produksjonsrute |
| **Feil Quantity Per** | Forbruket blir feil | Sjekk: Quantity Per = output per input. Hvis 1 M3 gir 400 LM, skriv 400 |
| **Scrap % for høy/lav** | Materialkost blir feil | Sjekk faktisk svinn i produksjonen |
| **Transportvare uten rute** | Vare er flagget som transportvare, men ingen Transportruter finnes | Legg til ruten i Transportruter-arket |

### ❌ Vanlige feil ved import i Marimo

| Feilmelding | Årsak | Løsning |
|-------------|-------|---------|
| "Validering fant feil — ingenting importert" | Excel-filen mangler ark eller kolonner | Sjekk at filen har alle nødvendige ark med korrekte kolonnenavn |
| "Kryssreferanse-feil" | En varekode i Stykkliste finnes ikke i Produktregister | Sjekk at alle Item No er registrert |
| "Ingen data lastet" | Databasen er tom | Last opp en Excel-fil via "Dataimport & Versjoner" |

### ⚠️ Kjente begrensninger i dagens kalkyle

Følgende felt registreres i Excel-arket, men utnyttes foreløpig ikke fullt i
Python-beregningen:

| Felt | Status | Planlagt forbedring |
|------|--------|---------------------|
| **Kapasitetskalender** | Visuell info kun — ikke i aktiv beregning | Flaskehalsanalyse planlegges |
| **Distance Km** / **Hours** (Transportruter) | Informasjon kun — påvirker ikke kostnaden | Planlegges |

Dette påvirker ikke standard bruk av modellen.

### ✅ Gode råd

1. **Start enkelt.** Legg inn 2-3 produkter først, sjekk at tallene gir mening, så utvider du.
2. **Sjekk at summen stemmer.** Beregn for hånd et enkelt produkt og sammenlign med modellens resultat i Marimo.
3. **Oppdater jevnlig.** Priser endrer seg — sett av tid til å oppdatere modellen hvert kvartal.
4. **Bruk kommentarfeltet.** Når du laster opp Excel i Marimo, skriv hva som er endret — da kan du senere se i endringsloggen hva som skjedde.
5. **Bruk versjonshistorikken.** Hvis noe går galt, kan du alltid gå tilbake til en tidligere versjon.
6. **Spør om hjelp.** Hvis tallene ser rare ut, sjekk om alle arkene er fylt ut riktig.

---

> **Trenger du hjelp?** Kontakt systemansvarlig.
>
> *Sist oppdatert: august 2026*