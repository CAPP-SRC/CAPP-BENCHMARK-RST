# CNC Operation Sheet — Vendor Rating Benchmark

Tool da riga di comando per confrontare due **operation sheet PDF** (generati da Fusion 360 / HSMWorks) e restituire un punteggio comparativo **0–100** stile **Vendor Rating**.

Dato un pezzo da lavorare e una libreria utensili condivisa, due gruppi possono definire cicli di lavorazione diversi. Questo tool li confronta in modo oggettivo su 13 driver raggruppati in 6 categorie, producendo una scorecard immediata e leggibile.

---

## Indice

- [Quick Start](#quick-start)
- [Installazione](#installazione)
- [Uso](#uso)
- [Output](#output)
- [Framework di Scoring](#framework-di-scoring)
- [Parsing dei PDF](#parsing-dei-pdf)
- [Personalizzazione](#personalizzazione)
- [Struttura del codice](#struttura-del-codice)
- [Requisiti dei PDF](#requisiti-dei-pdf)
- [Troubleshooting](#troubleshooting)

---

## Quick Start

```bash
pip install -r requirements.txt
python benchmark_cnc.py  gruppo_A.pdf  gruppo_B.pdf
```

---

## Installazione

### Prerequisiti

- Python 3.9 o superiore
- Sistema operativo: Windows, Linux o macOS

### Opzione 1 — pip (consigliata)

```bash
# Crea un virtual environment (opzionale ma consigliato)
python -m venv venv
source venv/bin/activate        # Linux/macOS
venv\Scripts\activate           # Windows

# Installa le dipendenze
pip install -r requirements.txt
```

### Opzione 2 — Conda

```bash
# Crea l'ambiente da file
conda env create -f environment.yml

# Attiva l'ambiente
conda activate benchmark-cnc
```

### Opzione 3 — Installazione manuale

```bash
pip install pdfplumber openpyxl
```

### Verifica installazione

```bash
python benchmark_cnc.py --help
```

---

## Uso

### Sintassi

```
python benchmark_cnc.py  <pdf_gruppo_A>  <pdf_gruppo_B>  [opzioni]
```

### Argomenti

| Argomento | Descrizione |
|-----------|-------------|
| `pdf_gruppo_A` | Percorso del PDF dell'operation sheet del primo gruppo |
| `pdf_gruppo_B` | Percorso del PDF dell'operation sheet del secondo gruppo |

### Opzioni

| Opzione | Default | Descrizione |
|---------|---------|-------------|
| `--xlsx <file.xlsx>` | — | Esporta i risultati in un file Excel formattato |
| `--tool-life <minuti>` | `20` | Soglia di vita utile massima per utensile (in minuti) |

### Esempi

**Confronto base** — output su console:

```bash
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf
```

**Con export Excel:**

```bash
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf  --xlsx report.xlsx
```

**Con soglia vita utile personalizzata** (15 minuti):

```bash
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf  --tool-life 15
```

**Combinazione completa:**

```bash
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf  --xlsx report.xlsx  --tool-life 25
```

---

## Output

### Console

Lo script produce un report testuale strutturato in 4 sezioni:

1. **Punteggio finale** — score complessivo 0–100 per ciascun gruppo e vincitore
2. **Dettaglio per categoria** — score e vincitore per ciascuna delle 6 categorie con relativo peso
3. **Dettaglio per driver** — tutti i 13 driver con valori grezzi, punteggi e indicatore `◄` sul vincitore
4. **Allarmi vita utile** — lista degli utensili che hanno superato la soglia di vita utile

Esempio di output:

```
════════════════════════════════════════════════════════════════════════════════════════
                             VENDOR RATING — BENCHMARK CNC
                                     NC02  vs  TP02
════════════════════════════════════════════════════════════════════════════════════════

         PUNTEGGIO FINALE                 NC02          TP02    Migliore
  ──────────────────────────────  ────────────  ────────────  ──────────
        Score Complessivo                89.9         84.5        NC02

  CATEGORIA                          Peso        NC02        TP02    Migliore
  ──────────────────────────────── ──────  ──────────  ──────────  ──────────
  Efficienza Temporale                30%     100.0 ◄      73.3          NC02
  Utilizzo Utensili                   20%      79.2       100.0 ◄        TP02
  Vita Utile                          20%      98.8 ◄      69.6          NC02
  ...
```

### Excel (opzionale)

Il file `.xlsx` generato con `--xlsx` contiene 2 fogli:

| Foglio | Contenuto |
|--------|-----------|
| **Scorecard** | Tabella completa dei driver con valori, punteggi, delta, vincitore per driver, riepilogo per categoria e punteggio finale pesato |
| **Vita Utile** | Dettaglio per ogni utensile (codice Product, riferimento T, tempo di impiego, percentuale di vita utilizzata, stato: OK / Moderato / Attenzione / ⚠ SUPERATO) |

Entrambi i fogli sono formattati con colori condizionali (verde = migliore, rosso = allarme, oro = vincitore).

---

## Framework di Scoring

### Architettura a 3 livelli

```
Punteggio Finale (0–100)
  └── Categoria (peso %)
        └── Driver (media dei driver nella categoria)
              └── Punteggio singolo (0–100)
```

### Pesi delle categorie

| # | Categoria | Peso | Motivazione |
|---|-----------|------|-------------|
| 1 | Efficienza Temporale | **30%** | Il tempo ciclo è il driver economico diretto: impatta produttività, costi macchina e lead time |
| 2 | Utilizzo Utensili | **20%** | Ogni utensile ha un costo e ogni cambio utensile è tempo morto non produttivo |
| 3 | Vita Utile | **20%** | Il superamento della vita utile comporta rischio di rottura utensile, scarti e danni alla macchina |
| 4 | Efficienza di Percorso | **15%** | Il rapporto taglio/rapido indica quanta parte del movimento è produttiva |
| 5 | Complessità del Ciclo | **10%** | Un ciclo più semplice è più facile da gestire, debuggare e manutenere |
| 6 | Aggressività di Taglio | **5%** | Feedrate e produttività indicano quanto il ciclo sfrutta le capacità della macchina |

### Driver per categoria

#### 1. Efficienza Temporale (30%)

| Driver | Metrica | Migliore = |
|--------|---------|------------|
| Tempo ciclo complessivo | Somma dei tempi Setup 1 + Setup 2 dall'header del PDF | Più basso |
| Tempo medio per operazione | Tempo ciclo / N° operazioni | Più basso |

#### 2. Utilizzo Utensili (20%)

| Driver | Metrica | Migliore = |
|--------|---------|------------|
| N° utensili univoci | Conteggio codici Product distinti | Più basso |
| N° cambi utensile | Transizioni tra tool diversi nel ciclo | Più basso |

#### 3. Vita Utile (20%)

| Driver | Metrica | Migliore = |
|--------|---------|------------|
| Score vita utile (non lineare) | Media dei punteggi per utensile (vedi sotto) | Più alto |
| Concentrazione utensile più impiegato | % del tempo ciclo assorbita dal singolo utensile più usato | Più basso |
| Penalità superamento vita | −50 punti per ogni utensile oltre la soglia | Più alto |

**Scoring non lineare della vita utile per singolo utensile:**

| Utilizzo vita (%) | Punteggio |
|-------------------|-----------|
| ≤ 50% | 100 |
| 50 – 75% | 80 |
| 75 – 100% | 60 |
| > 100% | max(0, 60 − (% − 100) × 2) — penalità rapida |

Questa curva premia l'uso efficiente degli utensili (non troppo poco, non troppo) e penalizza pesantemente il superamento della soglia.

#### 4. Efficienza di Percorso (15%)

| Driver | Metrica | Migliore = |
|--------|---------|------------|
| Rapporto taglio / (taglio + rapido) | Distanza di taglio / distanza totale | Più alto |
| Distanza complessiva | Taglio + rapido in mm | Più basso |

#### 5. Complessità del Ciclo (10%)

| Driver | Metrica | Migliore = |
|--------|---------|------------|
| N° operazioni totali | Conteggio operazioni su entrambi i setup | Più basso |
| Rapporto operazioni / utensile | N° operazioni / N° utensili univoci | Più basso |

#### 6. Aggressività di Taglio (5%)

| Driver | Metrica | Migliore = |
|--------|---------|------------|
| Feedrate medio ponderato | Media pesata per distanza di taglio del max feedrate per operazione | Più alto |
| Produttività | Distanza taglio totale / tempo ciclo totale [mm/min] | Più alto |

### Metodo di scoring per singolo driver

Per ogni driver, il gruppo migliore riceve **100 punti** e l'altro riceve un punteggio proporzionale:

```
Se "lower is better":
    Score_A = min(A, B) / A × 100
    Score_B = min(A, B) / B × 100

Se "higher is better":
    Score_A = A / max(A, B) × 100
    Score_B = B / max(A, B) × 100
```

Fanno eccezione i driver di **Vita Utile**, che usano scoring assoluto con soglie non lineari e penalità fisse.

### Calcolo del punteggio finale

```
Score_categoria = media(score dei driver nella categoria)
Score_finale = Σ (Score_categoria × Peso_categoria)
```

---

## Parsing dei PDF

Il parser è progettato per i **Setup Sheet** generati da **Autodesk Fusion 360** e **HSMWorks** (formato standard).

### Struttura attesa del PDF

Ogni PDF contiene uno o più setup (tipicamente 2), ciascuno con:

1. **Header del setup** — contiene i dati riepilogativi:
   - `Setup Sheet for Program XXXX`
   - `Number Of Operations: N`
   - `Number Of Tools: N`
   - `Estimated Cycle Time: XXm:XXs`
   - Lista degli utensili (T-numbers)

2. **Schede utensile** — un blocco per ogni utensile con:
   - Tipo, diametro, lunghezza, numero taglienti
   - Codice Product (identificativo univoco)
   - Distanza di taglio e rapido complessive
   - Tempo ciclo complessivo

3. **Schede operazione** — un blocco per ogni operazione con:
   - `Operation X/Y TXXXXX DXXXXX LXXXXX`
   - Strategia (Adaptive, Facing, Contour 2D, Drilling, Scallop, Bore, Flat, Contour)
   - Distanza di taglio e rapido
   - Feedrate massimo
   - Tempo ciclo stimato
   - Codice Product dell'utensile

### Dati estratti per operazione

| Campo | Fonte nel PDF | Utilizzo |
|-------|---------------|----------|
| Strategia | `Strategy:` o inferita da `Description:` | Classificazione tipo lavorazione |
| Riferimento T | `Operation X/Y TXXXXX` | Identificazione utensile |
| Codice Product | `Product:` | Identificazione univoca utensile |
| Distanza taglio | `Cutting Distance:` | Efficienza percorso |
| Distanza rapido | `Rapid Distance:` | Efficienza percorso |
| Feedrate max | `Maximum Feedrate:` | Aggressività taglio |
| Tempo ciclo | `Estimated Cycle Time:` | Efficienza temporale |

### Gestione delle anomalie

- **Strategia "Flat"**: alcuni PDF non riportano il campo `Strategy:` per le operazioni Flat — il parser la inferisce dalla `Description:`
- **Product code con suffissi**: codici come `"AQXR324SA32S con inserto QOMT1651R-M2..."` vengono troncati al codice base
- **Product code con prefissi**: codici come `"fresa a punta tonda VQ4SVBR04000"` vengono puliti automaticamente
- **Tempi ciclo**: il parser usa i tempi dall'header del setup (che includono overhead di cambio utensile) per i totali, e quelli delle singole operazioni per l'analisi per utensile

---

## Personalizzazione

### Soglia vita utile

Il parametro `--tool-life` (default: 20 minuti) definisce la durata massima di impiego per singolo utensile. Oltre questa soglia, l'utensile viene segnalato come allarme e penalizzato nello scoring.

```bash
# Soglia conservativa (15 min)
python benchmark_cnc.py  A.pdf  B.pdf  --tool-life 15

# Soglia permissiva (30 min)
python benchmark_cnc.py  A.pdf  B.pdf  --tool-life 30
```

### Pesi delle categorie

I pesi sono definiti nel dizionario `CATEGORY_WEIGHTS` nel codice sorgente:

```python
CATEGORY_WEIGHTS = {
    'Efficienza Temporale': 0.30,
    'Utilizzo Utensili': 0.20,
    'Vita Utile': 0.20,
    'Efficienza di Percorso': 0.15,
    'Complessità del Ciclo': 0.10,
    'Aggressività di Taglio': 0.05,
}
```

Per modificarli, editare il file `benchmark_cnc.py` assicurandosi che la somma dei pesi sia **1.00**.

### Aggiungere nuovi driver

Per aggiungere un driver, intervenire in 3 punti:

1. **`compute_metrics()`** — calcolare la metrica grezza dal PDF parsato
2. **`compute_scores()`** — aggiungere il driver alla categoria appropriata con `relative_score()`
3. **`print_report()` / `export_xlsx()`** — il driver appare automaticamente se aggiunto tramite `add()`

---

## Struttura del codice

```
benchmark_cnc.py          Script principale (monolitico, ~500 righe)
├── 1. PDF Parser         Estrazione dati dai PDF (pdfplumber)
│   ├── parse_cycle_time()    Conversione stringhe tempo → secondi
│   ├── extract_field()       Estrazione campi generici
│   ├── detect_strategy()     Riconoscimento strategia CAM
│   ├── extract_product_code() Estrazione codice Product
│   └── parse_pdf()           Parser principale → dict strutturato
│
├── 2. Calcolo Metriche   Aggregazione dati per gruppo
│   ├── extract_short_name()  Estrazione nome breve (es. "NC02")
│   └── compute_metrics()     Calcolo 25+ indicatori
│
├── 3. Sistema Scoring    Vendor Rating
│   ├── relative_score()      Punteggio relativo (migliore = 100)
│   ├── tool_life_score()     Scoring non lineare vita utile
│   └── compute_scores()      Orchestrazione completa → scorecard
│
├── 4. Output Console     Report testuale formattato
│   ├── fmt_time()            Formattazione secondi → "Xh XXm XXs"
│   └── print_report()        Stampa report completo
│
├── 5. Export Excel        Generazione .xlsx (opzionale)
│   └── export_xlsx()         Workbook con Scorecard + Vita Utile
│
└── 6. Main               CLI con argparse
    └── main()                Entry point
```

### Dipendenze

| Pacchetto | Versione min. | Utilizzo | Obbligatorio |
|-----------|---------------|----------|--------------|
| `pdfplumber` | 0.10.0 | Estrazione testo dai PDF | ✓ Sì |
| `openpyxl` | 3.1.0 | Generazione file Excel | Solo con `--xlsx` |

---

## Requisiti dei PDF

Il tool è stato progettato e testato per i **Setup Sheet** generati da:

- **Autodesk Fusion 360** (CAM → Setup Sheet)
- **HSMWorks** per SolidWorks

### Requisiti minimi

- Il PDF deve contenere almeno un blocco `Setup Sheet for Program XXXX`
- Ogni operazione deve avere il formato `Operation X/Y TXXXXX DXXXXX LXXXXX`
- Deve essere presente il campo `Product:` per ciascun utensile (usato come identificativo univoco)
- I tempi ciclo devono essere nel formato `Xh:XXm:XXs`, `XXm:XXs` o `XXs`

### Limitazioni note

- PDF scannerizzati (immagini) non sono supportati — serve testo estraibile
- Setup Sheet con formati personalizzati o lingua diversa dall'inglese potrebbero richiedere adattamenti al parser
- Il campo `Strategy:` potrebbe non essere presente per tutte le operazioni (es. Flat); il parser tenta di inferirla dalla `Description`

---

## Troubleshooting

| Problema | Soluzione |
|----------|----------|
| `ModuleNotFoundError: pdfplumber` | Esegui `pip install -r requirements.txt` |
| `Errore: nessuna operazione trovata` | Verifica che il PDF contenga operazioni nel formato atteso |
| Tempi ciclo a `0m 00s` | Il PDF potrebbe non contenere il campo `Estimated Cycle Time` nell'header del setup |
| Codice Product `N/A` | Il campo `Product:` potrebbe essere assente o in formato non standard |
| Nomi gruppo lunghi nell'output | Il tool estrae automaticamente nomi brevi (es. "NC02") dal Document Path |

---

## Licenza

MIT
