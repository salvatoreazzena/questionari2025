# questionari2025

## Script disponibili

### 1) Analisi preliminare campione

File: `build_analisi_preliminare.py`

Esecuzione:

```bash
python build_analisi_preliminare.py
```

Input atteso: `questionari_fonte.xlsx`

Output: `analisi_preliminare_campione.xlsx`

### 2) Normalizzazione provenienza

File: `normalizza_provenienza.py`

Esecuzione:

```bash
python normalizza_provenienza.py --input questionari_fonte.xlsx --output questionari_fonte_modificato.xlsx
```

Opzione inplace:

```bash
python normalizza_provenienza.py --input questionari_fonte.xlsx --inplace
```

### 3) Outlier spese e inferenza ND (pernottanti senza pacchetto)

File: `outlier_spese_pernottanti.py`

Esecuzione default:

```bash
python outlier_spese_pernottanti.py
```

Esecuzione con percorsi espliciti:

```bash
python outlier_spese_pernottanti.py --input questionari_fonte.xlsx --output questionari_outlier_pernottanti.xlsx --output-no-imputation questionari_outlier_pernottanti_no_imputazioni.xlsx
```

#### Regole applicate

- Campione analizzato: solo questionari con `durata_soggiorno >= 1` e `pacchetto` strettamente vuoto.
- Colonne spesa trattate:
	- `spese_trasporto_viaggio`
	- `spese_trasporto_interno`
	- `spese_alloggio`
	- `spese_alimentazione`
	- `spese_ristorazione`
	- `spese_souvenir`
	- `spese_altre`
- Gruppi omogenei per outlier: `macro_provenienza (ITALIANI/STRANIERI) x motivazione_principale`.
- Metrica per outlier: `spesa_pro_capite_giornaliera = spesa / (numero_componenti * durata_soggiorno)`.
- Soglie outlier: 5° e 95° percentile della metrica pro capite giornaliera per ogni colonna spesa all'interno del gruppo omogeneo.
- Outlier: evidenziati in rosso, senza sostituzione del valore originale.
- Valori ND (o simili): inferiti con mediana da questionari simili.
- Similarita per inferenza ND: `macro_provenienza`, `numero_componenti`, `durata_soggiorno`, `motivazione_principale`.
- Fallback inferenza donor: gruppo completo, poi macro+motivazione, poi macro, poi motivazione, poi globale colonna.

#### Output prodotto

- File Excel: `questionari_outlier_pernottanti.xlsx`
- File Excel parallelo senza imputazioni ND: `questionari_outlier_pernottanti_no_imputazioni.xlsx`
- Fogli:
	- `pernottanti_senza_pacchetto`
	- `statistiche`
- Include solo i questionari analizzati dal filtro.
- Colori celle spesa:
	- Rosso: outlier
	- Giallo: valore ND imputato
	- Arancione: valore ND imputato che risulta anche outlier rispetto a p5/p95 del gruppo
- Colonne aggiuntive nel foglio principale:
	- una sola colonna per ogni spesa nel formato `imputazione_<colonna_spesa>`
	- la cella indica come e stato imputato il valore (metodo e donor ID)
	- nel file parallelo senza imputazioni queste colonne non sono presenti, i valori ND restano ND ed e aggiunta la colonna `outlier_presente` (`SI`/`NO`)
- Contenuto foglio `statistiche`:
	- numero turisti input e analizzati (conteggi pesati su `numero_componenti`)
	- numero turisti con imputazioni per colonna spesa
	- numero turisti con imputazioni che risultano outlier per colonna spesa
	- numero turisti outlier per colonna spesa
	- tabella soglie reali usate per outlier (`p5_gruppo` e `p95_gruppo`) per ogni combinazione `colonna_spesa x macro_provenienza x motivazione_grp`, con `n_turisti_gruppo` pesato su `numero_componenti`

### 4) Estrazione sottocampioni (pernottanti senza pacchetto)

File: `estrai_sottocampioni.py`

Esecuzione default:

```bash
python estrai_sottocampioni.py
```

Esecuzione con percorsi espliciti:

```bash
python estrai_sottocampioni.py --input questionari_fonte_veri_outlier_sostituiti_imputati.xlsx --output questionari_sottocampioni.xlsx
```

#### Regole applicate

- Perimetro: solo righe con `durata_soggiorno > 0`, `pacchetto` strettamente vuoto, `fascia_età` valorizzata e `numero_componenti > 0`.
- Unita statistica: `numero_componenti` (turisti), non numero questionari.
- Quote per fasce d'eta: uguali tra tutte le fasce considerate nel campione.
- Target per cella: massimo comune possibile, pari al minimo `numero_componenti` disponibile tra tutte le celle richieste dal campione.
- Selezione: questionari interi (nessun frazionamento), ordinamento deterministico.

#### Campioni prodotti

- `campione_1`: 50% italiani e 50% stranieri, distribuiti nelle fasce d'eta comuni ai due gruppi.
- `campione_1a`: turisti non balneari (`motivazione_principale != BALNEARE`), mantenendo 50/50 italiani-stranieri e fasce d'eta.
- `campione_1b`: turisti con `motivazione_secondaria` valorizzata, mantenendo 50/50 italiani-stranieri e fasce d'eta.
- `campione_1c`: turisti con `durata_soggiorno <= 30`, con doppia stratificazione per classi durata `1-3`, `4-7`, `8-14`, `15-30` e fasce d'eta (comuni tra italiani/stranieri), mantenendo 50/50 italiani-stranieri.
- `campione_1d`: turisti con top 5 `motivazione_principale` (ranking per `numero_componenti`), mantenendo 50/50 italiani-stranieri e fasce d'eta.
- `campione_2`: top 5 paesi esteri (ranking per `numero_componenti`) x fasce d'eta comuni.
- `campione_3`: top 5 regioni italiane (ranking per `numero_componenti`) x fasce d'eta comuni.

#### Output prodotto

- File Excel: `questionari_sottocampioni.xlsx`
- Fogli:
	- `campione_1`
	- `campione_1a`
	- `campione_1b`
	- `campione_1c`
	- `campione_1d`
	- `campione_2`
	- `campione_3`
	- `diagnostica_celle` (disponibili, target, selezionati e overshoot per ogni cella)
	- `meta` (riepilogo run e parametri effettivi dei campioni)
