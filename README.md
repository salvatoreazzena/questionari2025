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

### 5) Analisi qualitative Campione 1

File: `build_analisi_qualitative_campione1.py`

Esecuzione default:

```bash
python3 build_analisi_qualitative_campione1.py
```

Esecuzione con percorsi espliciti:

```bash
python3 build_analisi_qualitative_campione1.py --input questionari_sottocampioni.xlsx --sheet campione_1 --output analisi_qualitative_campione_1.xlsx
```

#### Output prodotto

- File Excel: `analisi_qualitative_campione_1.xlsx`
- Fogli principali:
	- `web_prenotazione`
	- `web_cosa_prenotata`
	- `stanziale_itinerante`
	- `arrivi_x_destinazione`
	- `alloggio_prevalente`
	- `motivazione_x_alloggio`
	- `dest_top10_provenienze`
	- `dest_top_motivazioni`
	- `spesa_futura_assoc`
	- `top20_dest_x_motivaz`
	- `motivaz_prim_second`
	- `giudizio_distrib`
	- `giudizio_top15_dest`
	- `punti_forza_temi`
	- `punti_forza_top_temi`
	- `punti_forza_top20`
	- `dissenso_temi`
	- `dissenso_priorita_intervento`
	- `dissenso_top20`
	- `dissenso_dettaglio`

#### Note metodologiche sintetiche

- Distinzione italiani/stranieri costruita da `stato_provenienza`.
- `destinazione_prevalente`: estratta con la stessa logica gia usata in `build_analisi_preliminare.py`.
- `stanziale/itinerante`: una sola destinazione con giorni positivi vs piu destinazioni con giorni positivi.
- `provenienza_associata`: regione per italiani, stato per stranieri.
- `media_giudizio`: conversione `Pessimo=1`, `Sufficiente=2`, `Buono=3`, `Ottimo=4`.
- Le categorie `NON DEFINITO` vengono mantenute nelle analisi distributive per preservare i totali del campione.
- `web_specifica` continua a essere classificato per regole testuali; `apprezzamenti` e `miglioramenti/consigli` usano invece una classificazione ML zero-shot Hugging Face.
- La logica multi-label e gestita a livello di risposta: il testo viene spezzato in frammenti distinti e ogni frammento riceve una sola categoria prevalente. In questo modo `mare` resta in una sola categoria, mentre `spiagge e cibo` genera due temi realistici nella stessa risposta.
- Prima della classificazione, i frammenti subiscono una normalizzazione di base: rimozione di articoli, introduttori valutativi e aggettivi generici, in modo da ricondurre espressioni come `il mare`, `mi è piaciuto il mare`, `bellissime spiagge` o `buon cibo` a nuclei piu stabili.
- La normalizzazione ricompone anche varianti singolare/plurale e maschile/femminile piu frequenti nel dominio, per esempio `spiaggia/spiagge`, `strada/strade`, `prezzo/prezzi`, `costoso/costosa/costosi`.
- Risposte vuote, nulle o non informative come `NULL`, `vuoto`, `tutto`, `niente`, `tutto bene`, `niente da migliorare` o `non saprei` vengono escluse dalla classificazione tematica.
- Il foglio `meta` riporta anche quanti testi sono stati considerati validi e quanti sono stati scartati come vuoti/nulli o non informativi, separatamente per `apprezzamenti` e `miglioramenti`.
- Le tabelle di sintesi restano leggibili e mostrano soprattutto `questionari`, `componenti` e `pct_su_gruppo`; i punteggi di confidenza del modello restano disponibili nell'audit tecnico.
- `punti_forza_top_temi` confronta in modo diretto italiani e stranieri sui temi piu apprezzati.
- `dissenso_priorita_intervento` isola, tra chi ha espresso giudizi `SUFFICIENTE` o `PESSIMO`, le priorita di intervento urgenti per provenienza e livello di giudizio.
- Durante la run del campione base, la classificazione ML mostra una barra di avanzamento con ETA per `apprezzamenti` e `dissenso`; `web_specifica` resta fuori da questa fase perche continua a usare regole testuali.
- La stessa logica ML multi-label viene applicata anche al blocco `campione_1d`, con ranking dei temi per `macro_provenienza x motivazione_principale`, top frammenti e audit dedicato.
