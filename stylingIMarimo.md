Å style en **Marimo**-app (en reaktiv notebook for Python) handler i stor grad om å bruke Marimo sitt eget UI-bibliotek (`marimo.ui`), markdown-formatering, og standard HTML/CSS. Siden Marimo rendrer alt som web-elementer, har du stor fleksibilitet.

Her er en oppsummering av de beste måtene å style en Marimo-app på, inkludert hvordan du legger til logo og tilpasser oppsettet:

---

## 1. Legge til logo og bilder

For å legge til en logo øverst i appen din, kan du bruke Marimo sin innebygde `mo.image`-funksjon eller standard HTML via Markdown.

* **Med `mo.image` (anbefalt):**
```python
import marimo as mo

mo.image(
    src="https://din-nettside.no/logo.png",  # Eller lokal sti f.eks. "assets/logo.png"
    alt="Min App Logo",
    width=150,
    align="center"
)

```


* **Med Markdown:**
```python
mo.md("![Logo](assets/logo.png){width=100px}")

```



---

## 2. Layout og organisering (`mo.vstack` og `mo.hstack`)

Standard oppførsel i Marimo er at elementer stables vertikalt. Du kan styre dette presist for å lage "dashboard"-følelsen:

* **`mo.hstack` (Horisontal rad):** Perfekt for å sette en logo ved siden av en tittel, eller knapper på på en rad.
```python
mo.hstack([
    mo.image("logo.png", width=50),
    mo.md("# Velkommen til min app")
], justify="start", align="center")

```


* **`mo.vstack` (Vertikal stabel):** For å gruppere elementer vertikalt med kontrollert mellomrom (`gap`).

---

## 3. Bruke paneler og kort (`mo.ui.panel` / `mo.card`)

For å gi appen struktur og ramme inn innhold (som KPI-er, grafer eller input-felter), bør du bruke kort:

```python
mo.card(
    mo.vstack([
        mo.md("### Nøkkeltall"),
        mo.md("**Salg denne måneden:** 150 000 kr")
    ])
)

```

---

## 4. Egendefinert CSS (Custom Styling)

Hvis du vil endre farger, fonter eller bakgrunner, kan du injisere standard CSS direkte i appen ved hjelp av `mo.style` eller HTML i en markdown-celle.

* **Endre stil på spesifikke elementer:**
```python
tekst = mo.md("Dette er en viktig beskjed!")
mo.style(tekst, {"color": "red", "font-weight": "bold", "background-color": "#f0f0f0"})

```


* **Global CSS (for hele appen):**
Du kan legge til en celle med rå HTML for å endre bakgrunnsfarge eller fonter globalt:
```python
mo.html("<style> body { background-color: #fafafa; font-family: 'Arial'; } </style>")

```



---

## 5. Mørkt og lyst tema (Dark/Light Mode)

Marimo har innebygd støtte for både lyst og mørkt tema. Du kan styre dette via Marimo-grensesnittet (innstillinger øverst i hjørnet), eller du kan låse appen til et bestemt tema når du kjører den fra terminalen:

```bash
marimo run app.py --theme light
# eller
marimo run app.py --theme dark

```

---

## 6. Sideoppsett: App-modus vs. Notebook-modus

Når du deler appen med andre (bruker `marimo run`), skjules all kode automatisk. Du kan velge to ulike visningsformater i Marimo-editoren (gjøres via "Layout"-menyen i editoren):

* **Vertical (Standard):** Alt flyter nedover som en vanlig nettside.
* **Grid / Dashboard (Egendefinert):** Lar deg dra og slippe elementer i et rutenett for å lage et skreddersydd dashboard-utseende.