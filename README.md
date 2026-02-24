# CNC Operation Sheet — Vendor Rating Benchmark

Suite di tool da riga di comando per confrontare **operation sheet PDF** (generati da Fusion 360 / HSMWorks) e restituire punteggi comparativi **0–100** stile **Vendor Rating**.

Dato un pezzo da lavorare e una libreria utensili condivisa, diversi gruppi di lavoro umani e/o software di pianificazione CAPP (Computer Aided Process Planning) possono definire cicli di lavorazione diversi. Questi tool li confrontano in modo oggettivo su 13 driver raggruppati in 6 categorie, producendo una scorecard immediata e leggibile.

La suite è composta da due script:

| Script | Scopo | Input |
|--------|-------|-------|
| `benchmark_cnc.py` | Confronto **1 vs 1** tra due gruppi | 2 file PDF |
| `multi_benchmark_cnc.py` | Classifica **N gruppi** simultaneamente | N file PDF o cartella |

---

## Indice

- [Quick Start](#quick-start)
- [Installazione](#installazione)
- [benchmark\_cnc.py — Confronto 1 vs 1](#benchmark_cncpy--confronto-1-vs-1)
- [multi\_benchmark\_cnc.py — Classifica N gruppi](#multi_benchmark_cncpy--classifica-n-gruppi)
- [Framework di Scoring](#framework-di-scoring)
- [Parsing dei PDF](#parsing-dei-pdf)
- [Personalizzazione](#personalizzazione)
- [Struttura del codice](#struttura-del-codice)
- [Requisiti dei PDF](#requisiti-dei-pdf)
- [Troubleshooting](#troubleshooting)

---

## Quick Start

```bash
# Installa le dipendenze
pip install -r requirements.txt

# Confronto 1 vs 1
python benchmark_cnc.py  gruppo_A.pdf  gruppo_B.pdf

# Classifica N gruppi (da cartella)
python multi_benchmark_cnc.py  ./cartella_pdf/

# Classifica N gruppi (file singoli)
python multi_benchmark_cnc.py  gruppo_A.pdf  gruppo_B.pdf  gruppo_C.pdf
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
python multi_benchmark_cnc.py --help
```

### File forniti

```
benchmark_cnc.py          Confronto 1 vs 1
multi_benchmark_cnc.py    Classifica N gruppi
requirements.txt          Dipendenze per pip
environment.yml           Ambiente per Conda
README.md                 Questo file
```

---

## benchmark_cnc.py — Confronto 1 vs 1

Confronta due operation sheet PDF e restituisce un punteggio comparativo dettagliato.

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

```bash
# Confronto base — output su console
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf

# Con export Excel
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf  --xlsx report.xlsx

# Con soglia vita utile personalizzata (15 min)
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf  --tool-life 15

# Combinazione completa
python benchmark_cnc.py  NC02_SHEET.pdf  TP02_SHEET.pdf  --xlsx report.xlsx  --tool-life 25
```

### Output console

Il report testuale è strutturato in 4 sezioni:

1. **Punteggio finale** — score complessivo 0–100 per ciascun gruppo e vincitore
2. **Dettaglio per categoria** — score e vincitore per ciascuna delle 6 categorie con peso
3. **Dettaglio per driver** — tutti i 13 driver con valori grezzi, punteggi e indicatore `◄` sul vincitore
4. **Allarmi vita utile** — lista degli utensili che hanno superato la soglia

Esempio:

```
════════════════════════════════════════════════════════════════════════
                       VENDOR RATING — BENCHMARK CNC
                               NC02  vs  TP02
════════════════════════════════════════════════════════════════════════

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

### Output Excel (con `--xlsx`)

| Foglio | Contenuto |
|--------|-----------|
| **Scorecard** | Tabella completa dei driver con valori, punteggi, delta, vincitore, riepilogo per categoria e punteggio finale pesato |
| **Vita Utile** | Dettaglio per utensile: codice Product, riferimento T, tempo di impiego, % vita utilizzata, stato (OK / Moderato / Attenzione / ⚠ SUPERATO) |

---

## multi_benchmark_cnc.py — Classifica N gruppi

Confronta **N gruppi simultaneamente** e restituisce una classifica ordinata per punteggio. Utile per confronti ampi (es. tutti i gruppi di un corso, tutte le revisioni di un ciclo, ecc.).

### Sintassi

```
python multi_benchmark_cnc.py  <input_1> [input_2] ... [input_N]  [opzioni]
```

Ogni `<input>` può essere:
- Un **file PDF** singolo
- Una **cartella** contenente file PDF (verranno letti tutti i `.pdf` al suo interno)
- Un **mix** di file e cartelle

### Argomenti

| Argomento | Descrizione |
|-----------|-------------|
| `inputs` | Uno o più file PDF e/o cartelle (minimo 2 PDF risultanti) |

### Opzioni

| Opzione | Default | Descrizione |
|---------|---------|-------------|
| `--xlsx <file.xlsx>` | — | Esporta i risultati in un file Excel formattato |
| `--tool-life <minuti>` | `20` | Soglia di vita utile massima per utensile (in minuti) |

### Esempi

```bash
# Tutti i PDF in una cartella
python multi_benchmark_cnc.py  ./pdf_folder/

# File singoli selezionati
python multi_benchmark_cnc.py  NC01.pdf  NC02.pdf  TP01.pdf  TP02.pdf

# Mix di cartelle e file
python multi_benchmark_cnc.py  ./cartella_NC/  ./cartella_TP/  extra.pdf

# Con export Excel
python multi_benchmark_cnc.py  ./pdf_folder/  --xlsx classifica.xlsx

# Con soglia vita utile personalizzata
python multi_benchmark_cnc.py  ./pdf_folder/  --xlsx classifica.xlsx  --tool-life 15
```

### Output console

Il report multi-gruppo è strutturato in 5 sezioni:

1. **Classifica finale** — podio con medaglie (🥇🥈🥉) e punteggi
2. **Dettaglio per categoria** — tabella N colonne con punteggi per categoria, ordinate per score
3. **Dettaglio per driver** — tutti i 13 driver con valori e punteggi per ogni gruppo, indicazione del migliore
4. **Allarmi vita utile** — lista utensili fuori soglia per ogni gruppo
5. **Metodologia** — riepilogo del metodo di scoring applicato

Esempio:

```
════════════════════════════════════════════════════════════════════════
             VENDOR RATING — MULTI-GROUP BENCHMARK CNC
                        6 gruppi confrontati
════════════════════════════════════════════════════════════════════════

        CLASSIFICA FINALE
  ──────────────────────────────
  🥇 1°  NC02                88.9 / 100
  🥈 2°  NC03                84.5 / 100
  🥉 3°  TP02                83.2 / 100
     4°  TP03                81.8 / 100
     5°  TP01                81.0 / 100
     6°  NC01                76.5 / 100
```

### Output Excel (con `--xlsx`)

| Foglio | Contenuto |
|--------|-----------|
| **Classifica** | Podio con punteggi, dettaglio per categoria con evidenziazione del migliore, nota metodologica |
| **Scorecard Dettaglio** | Tutti i 13 driver con valori e punteggi per ogni gruppo, raggruppati per categoria |
| **Vita Utile** | Matrice completa utensili × gruppi con tempi, % vita e stato per ogni combinazione |
| **Dati Radar** | Tabella numerica dei punteggi per categoria, pronta per generare un grafico radar in Excel |

### Note sul naming dei gruppi

Lo script estrae automaticamente un nome breve dal campo `Document Path` del PDF o dal nome del file (es. `NC02` da `X_NC02-FORI_EDIT_12100709 v4`). Se due PDF producono lo stesso nome breve, viene aggiunto un suffisso progressivo (es. `NC02_1`, `NC02_2`).

---

## Framework di Scoring

Il framework è **identico** per entrambi gli script. L'unica differenza è che `benchmark_cnc.py` confronta 2 gruppi mentre `multi_benchmark_cnc.py` confronta N gruppi.

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
| 3 | Vita Utile | **20%** | Il superamento della vita utile comporta rischio di rottura, scarti e danni alla macchina |
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

Per ogni driver, il gruppo migliore riceve **100 punti** e gli altri ricevono un punteggio proporzionale:

```
Se "lower is better":
    Score_i = min(tutti) / valore_i × 100

Se "higher is better":
    Score_i = valore_i / max(tutti) × 100
```

Fanno eccezione i driver di **Vita Utile**, che usano scoring assoluto con soglie non lineari e penalità fisse.

### Calcolo del punteggio finale

```
Score_categoria = media(score dei driver nella categoria)
Score_finale = Σ (Score_categoria × Peso_categoria)
```

---

## Parsing dei PDF

Il parser è progettato per i **Setup Sheet** generati da **Autodesk Fusion 360** e **HSMWorks** (formato standard). È **condiviso** tra i due script.

### Struttura attesa del PDF

Ogni PDF contiene uno o più setup (tipicamente 2), ciascuno con:

1. **Header del setup** — dati riepilogativi:
   - `Setup Sheet for Program XXXX`
   - `Number Of Operations: N`
   - `Number Of Tools: N`
   - `Estimated Cycle Time: XXm:XXs`

2. **Schede utensile** — un blocco per ogni utensile con tipo, diametro, codice Product, distanze e tempi

3. **Schede operazione** — un blocco per ogni operazione con strategia, distanze, feedrate, tempo ciclo e codice Product

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

- **Strategia "Flat"**: alcuni PDF non riportano il campo `Strategy:` per queste operazioni — il parser la inferisce dalla `Description:`
- **Product code con suffissi**: codici come `"AQXR324SA32S con inserto QOMT1651R-M2..."` vengono troncati al codice base
- **Product code con prefissi**: codici come `"fresa a punta tonda VQ4SVBR04000"` vengono puliti automaticamente
- **Tempi ciclo**: il parser usa i tempi dall'header del setup (che includono overhead di cambio utensile) per i totali, e quelli delle singole operazioni per l'analisi per utensile
- **Naming gruppi**: il nome breve viene estratto dal `Document Path` nel PDF o dal nome del file, cercando pattern tipo `NC01`, `TP02`, `GR03`, ecc.

---

## Personalizzazione

### Soglia vita utile

Il parametro `--tool-life` (default: 20 minuti) è disponibile in entrambi gli script:

```bash
python benchmark_cnc.py        A.pdf  B.pdf  --tool-life 15
python multi_benchmark_cnc.py  ./pdf/         --tool-life 30
```

### Pesi delle categorie

I pesi sono definiti nel dizionario `CATEGORY_WEIGHTS` nel codice sorgente di ciascuno script:

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

Per modificarli, editare il file assicurandosi che la somma dei pesi sia **1.00**. La modifica va fatta in entrambi gli script se si vogliono risultati coerenti.

### Aggiungere nuovi driver

Per aggiungere un driver, intervenire in 3 punti:

1. **`compute_metrics()`** — calcolare la metrica grezza dal PDF parsato
2. **`compute_scores()` / `compute_all_scores()`** — aggiungere il driver alla categoria appropriata
3. I report (console e Excel) includono automaticamente i nuovi driver

---

## Struttura del codice

Entrambi gli script seguono la stessa architettura a 6 moduli:

```
benchmark_cnc.py / multi_benchmark_cnc.py
├── 1. PDF Parser               Estrazione dati dai PDF (pdfplumber)
│   ├── parse_cycle_time()          Conversione stringhe tempo → secondi
│   ├── extract_field()             Estrazione campi generici
│   ├── detect_strategy()           Riconoscimento strategia CAM
│   ├── extract_product_code()      Estrazione codice Product
│   ├── extract_short_name()        Nome breve del gruppo
│   └── parse_pdf()                 Parser principale → dict strutturato
│
├── 2. Calcolo Metriche         Aggregazione dati per gruppo
│   └── compute_metrics()           Calcolo 25+ indicatori
│
├── 3. Sistema Scoring          Vendor Rating
│   ├── relative_score[_multi]()    Punteggio relativo (2 o N gruppi)
│   ├── tool_life_score()           Scoring non lineare vita utile
│   └── compute_[all_]scores()      Orchestrazione → scorecard
│
├── 4. Output Console           Report testuale formattato
│   ├── fmt_time()                  Formattazione secondi
│   └── print_[multi_]report()      Stampa report
│
├── 5. Export Excel             Generazione .xlsx (opzionale)
│   └── export_[multi_]xlsx()       Workbook formattato
│
└── 6. Main                    CLI con argparse
    ├── collect_pdfs()              [solo multi] Raccolta PDF da input
    └── main()                      Entry point
```

### Differenze chiave tra i due script

| Aspetto | `benchmark_cnc.py` | `multi_benchmark_cnc.py` |
|---------|---------------------|--------------------------|
| Input | Esattamente 2 PDF | N PDF e/o cartelle |
| Scoring | `relative_score()` — confronto a coppie | `relative_score_multi()` — confronto a N |
| Console | Tabella 2 colonne | Tabella N colonne + classifica con podio |
| Excel | 2 fogli (Scorecard, Vita Utile) | 4 fogli (Classifica, Scorecard, Vita Utile, Dati Radar) |
| Naming | Dal `Document Path` | Dal `Document Path` o nome file, con gestione duplicati |

### Dipendenze

| Pacchetto | Versione min. | Utilizzo | Obbligatorio |
|-----------|---------------|----------|--------------|
| `pdfplumber` | 0.10.0 | Estrazione testo dai PDF | ✓ Sì |
| `openpyxl` | 3.1.0 | Generazione file Excel | Solo con `--xlsx` |

---

## Requisiti dei PDF

Entrambi i tool sono progettati e testati per i **Setup Sheet** generati da:

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
| `Errore: servono almeno 2 file PDF` | Verifica che la cartella contenga almeno 2 file `.pdf` |
| `Errore: nessuna operazione trovata` | Verifica che il PDF contenga operazioni nel formato atteso |
| Tempi ciclo a `0m 00s` | Il PDF potrebbe non contenere `Estimated Cycle Time` nell'header |
| Codice Product `N/A` | Il campo `Product:` potrebbe essere assente o in formato non standard |
| Nomi gruppo lunghi o errati | Lo script cerca pattern `NC01`, `TP02`, ecc. nel Document Path e nel nome file; se non trovati, usa la prima parola del Document Path |
| Nomi gruppo duplicati | `multi_benchmark_cnc.py` aggiunge automaticamente suffissi `_1`, `_2`, ecc. |

---

## Licenza

MIT

## Autore

[RAW](https://rawmain.github.io/) - _aka RST_
