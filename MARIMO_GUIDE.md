# Marimo Guide - Erfaringer og beste praksis

## Hva er Marimo?

Marimo er en **reaktiv Python-notatblokk** (alternativ til Jupyter). Hver celle er en funksjon som automatisk oppdateres når avhengigheter endres. I motsetning til Jupyter hvor du må kjøre celler manuelt i riktig rekkefølge, sporer Marimo avhengighetsgrafen og kjører celler automatisk.

---

## 1. Grunnleggende syntax

### Celle-dekoratør
```python
@app.cell
def _(mo, pd, data):
    # Celleinnhold her
    return (variabel1, variabel2)
```

- `@app.cell` markerer en celle
- `def _(parametere):` - parametere er **avhengigheter** som Marimo automatisk injecter
- `return (variabel1,)` - returnerte variabler blir **tilgjengelige for andre celler**
- En celle kan returnere flere variabler: `return (a, b, c)`

### Importer
```python
@app.cell
def _():
    import marimo as mo
    import pandas as pd
    return (mo, pd)
```

---

## 2. UI-komponenter

### `mo.ui.number()` - Tall-input
```python
antall = mo.ui.number(
    label="Antall (stk)",
    start=1,        # minimumsverdi
    stop=100000,    # maksimumsverdi
    step=1,         # steg-størrelse
    value=1000,     # startverdi
)
antall  # ← må stå på slutten for å vises
```

**Viktig:** For å vise UI-elementet, må det stå som siste uttrykk i cellen (uten tilordning).

### `mo.ui.text()` - Tekst-input
```python
søk = mo.ui.text(label="🔍 Søk")
søk
```

### `mo.ui.data_editor()` - Redigerbar tabell (ANBEFALT)
```python
df = pd.DataFrame({
    "Varenr": ["RM001", "RM002"],
    "Org. pris": [100, 200],
    "Ny pris": [100, 200],
})
editor = mo.ui.data_editor(df, editable_columns=["Ny pris"])
editor
```

**Viktig:** `editable_columns` bestemmer hvilke kolonner som kan redigeres. `data_editor` støtter IKKE `label` eller `column_widths` i marimo 0.23.x.

**Hente verdier:** `editor.value` returnerer en DataFrame med alle endringer.

### `mo.ui.table()` - Skrivebeskyttet tabell
```python
mo.ui.table(df, selection=None)  # selection=None = ingen radvalg
```

### `mo.ui.run_button()` - Kjøreknapp
```python
knapp = mo.ui.run_button(label="▶️ Kjør", kind="neutral")
knapp
```

**Brukes til å trigge simuleringer:** `if knapp.value: ...`

### `mo.ui.file()` - Filopplasting
```python
file_upload = mo.ui.file(
    label="Velg Excel-fil",
    filetypes=[".xlsx"],
    multiple=False,
)
file_upload
```

**Hente innhold:** `file_upload.value[0].contents` (bytes)

### `mo.download()` - Nedlastingsknapp
```python
mo.download(
    label="📥 Last ned",
    filename="rapport.pdf",
    data=pdf_bytes,
)
```
**Merk:** `mo.download()` (uten `ui.`) - `mo.ui.download()` finnes ikke i marimo 0.23.x.

### `mo.md()` - Markdown
```python
mo.md("""
# Tittel
**Fet tekst** og *kursiv*
""")
```

### `mo.vstack()` - Stable elementer vertikalt
```python
mo.vstack([mo.md("Tittel"), tabell, knapp])
```

### `mo.output.replace()` - Erstatt output
```python
mo.output.replace(mo.md("### ✅ Ferdig!"))
```

**Bruk:** Når du vil erstatte output i en celle som allerede har vist noe (f.eks. etter en knappetrykk).

---

## 3. VIKTIGE "finurligheter" og fallgruver

### ⚠️ Felle 1: Variabel på slutten av cellen

For at et UI-element skal vises, må det stå som **siste uttrykk i cellen UTEN tilordning**:

```python
# ✅ RIKTIG - vises
editor = mo.ui.data_editor(df, editable_columns=["Ny pris"])
editor  # ← på slutten, uten innrykk

# ❌ FEIL - vises ikke (inni if-blokk)
if data:
    editor = mo.ui.data_editor(df, editable_columns=["Ny pris"])
    editor  # ← inni if-blokken - Marimo ser den ikke!

# ✅ RIKTIG - vises (utenfor if-blokk)
editor = None
if data:
    editor = mo.ui.data_editor(df, editable_columns=["Ny pris"])
editor  # ← på slutten, uten innrykk
```

### ⚠️ Felle 2: `mo.output.replace()` vs return

- **`mo.output.replace()`** - brukes INNI en celle for å erstatte output (f.eks. etter knappetrykk)
- **`return (variabel,)`** - brukes for å gjøre variabler tilgjengelige for andre celler
- **`variabel` på slutten** - brukes for å vise UI-elementet i cellen

### ⚠️ Felle 3: Dobbel-rendering av UI-elementer

Hvis du oppretter `mo.ui.number()`-objekter i én celle OG viser dem i en annen celle (via `mo.output.replace()`), kan de samme objektene rendres to ganger. Dette skaper konflikt der endring i ett felt påvirker et annet.

**Løsning:** Bruk `mo.ui.data_editor()` i stedet - det er en enkelt komponent som håndterer all rendering selv.

### ⚠️ Felle 4: `mo.ui.dataframe()` vs `mo.ui.data_editor()`

| Egenskap | `mo.ui.dataframe()` | `mo.ui.data_editor()` |
|---|---|---|
| Redigerbar | Nei (kun visning) | Ja |
| `editable_columns` | Støttes ikke | ✅ Støttes |
| `label` | Støttes ikke i 0.23.x | Støttes ikke i 0.23.x |
| `column_widths` | Støttes ikke i 0.23.x | Støttes ikke i 0.23.x |

**Bruk `mo.ui.data_editor()` når du vil ha redigerbare tabeller.**

### ⚠️ Felle 5: Tuple-nøkler i dicts

Når du sender data mellom celler (f.eks. til PDF-generator), vær forsiktig med tuple-nøkler:

```python
# Dette fungerer fint i marimo:
overrides.bom_scrap[("FG001", "RM001")] = 5.0

# Men hvis du konverterer til JSON/streng, ødelegges tuplene:
{str(k): v for k, v in overrides.bom_scrap.items()}
# → "{('FG001', 'RM001')}": 5.0  ← streng, ikke tuple!

# Løsning: behold som dict (ikke str-konverter)
dict(overrides.bom_scrap)  # ← beholder tuple-nøkler
```

### ⚠️ Felle 6: Filter-logikk med kaskade

Når du filtrerer data, må du tenke på avhengighetskjeden:

```
vareFilter → filtered_products → filtered_bom_lines → filtered_work_centers
                              → filtered_routing_lines → filtered_operations
                              → filtered_byproduct_rules
```

Hver celle som viser filtrerte data må bruke de filtrerte datasettene, ikke de originale.

---

## 4. Beste praksis for struktur

### Celle-organisering (anbefalt rekkefølge)

1. **Import-celle** - alle imports
2. **Tittel/header** - `mo.md()` med tittel
3. **Data-lasting** - filopplasting og parsing
4. **Filter** - søkefelt og filtreringslogikk
5. **Data-visning** - tabeller med originale data
6. **Justering** - redigerbare tabeller (`mo.ui.data_editor`)
7. **Simulering** - kjøreknapp og beregninger
8. **Resultater** - visning av simuleringsresultater
9. **Eksport** - PDF/Excel-eksport

### Navnekonvensjoner

- **UI-elementer:** `rm_price_df`, `bom_scrap_df`, `wc_cost_df`, `routing_df`
- **Filtrerte data:** `filtered_products`, `filtered_bom_lines`
- **Knapper:** `run_button`, `export_pdf_button`
- **Interne variabler:** `_tmp`, `_rows`, `_df` (med underscore for å unngå eksport)

### Returnere vs ikke returnere

- Returner **kun** variabler som andre celler trenger
- Bruk `_`-prefiks for interne variabler som ikke skal returneres
- En celle kan returnere flere variabler: `return (a, b, c)`

---

## 5. Feilsøking

### "An ancestor raised an exception"
Betyr at en celle som denne cellen er avhengig av har kastet en feil. Sjekk cellene over i avhengighetsgrafen.

### UI-element vises ikke
Sjekk at:
1. Elementet står på slutten av cellen (uten tilordning)
2. Det ikke er inni en `if`-blokk (med mindre det også står utenfor)
3. Variabelen ikke er `None`

### "too many values to unpack"
Sannsynligvis en tuple som har blitt konvertert til streng. Sjekk `str(k)`-konverteringer.

### Filter påvirker feil verdier
Sannsynligvis dobbel-rendering av `mo.ui.number()`-objekter. Bytt til `mo.ui.data_editor()`.

---

## 6. Nyttige ressurser

- [Marimo dokumentasjon](https://docs.marimo.io/)
- [Marimo UI-komponenter](https://docs.marimo.io/api/inputs/)
- [Marimo GitHub](https://github.com/marimo-team/marimo)
