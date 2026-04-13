# Pipeline Outlier e Imputazioni - Documentazione Operativa

## Obiettivo
Questa pipeline separata produce:
1. Analisi statistica iniziale delle spese.
2. Individuazione outlier per componente di spesa con soglie p5/p95.
3. File outlier dedicato alla verifica manuale con evidenziazione gialla delle celle outlier.
4. Imputazione valori mancanti ND/NR tramite questionari equivalenti.
5. File finale con valori imputati evidenziati in verde.

## Perimetro
La pipeline e completamente separata dal resto del progetto.
Cartella dedicata: `outlier_pipeline`.
Script principale: `outlier_pipeline/run_outlier_pipeline.py`.
Output: `outlier_pipeline/output/`.

Perimetro analitico corrente (fase operativa):
- Solo pernottanti senza pacchetto.
- Definizione pernottante: `durata_soggiorno >= 1`.
- Definizione senza pacchetto: colonna `pacchetto` vuota.

## Colonne usate
- Chiavi di gruppo outlier:
  - `stato_provenienza`
  - `motivazione_principale`
- Chiavi equivalenza imputazione:
  - `stato_provenienza` (vincolo forte)
  - `motivazione_principale` (vincolo forte)
  - `tipologia_sistemazione`
  - `fascia_età`
  - `classe_numero_componenti` (derivata da `numero_componenti`)
- Variabili spesa:
  - `spese_trasporto_viaggio`
  - `spese_trasporto_interno`
  - `spese_alloggio`
  - `spese_alimentazione`
  - `spese_ristorazione`
  - `spese_souvenir`
  - `spese_altre`

## Trasformazione metrica spesa
Le colonne spesa del file sorgente sono importi totali.
Per il calcolo soglie outlier, la pipeline usa la metrica derivata pro capite giornaliera:

spesa_pcg = spesa_totale_categoria / (numero_componenti * durata_soggiorno)

Se il denominatore e nullo o non valido, la cella non e valutabile per le soglie outlier.

## Regole outlier
1. Soglie per ogni variabile spesa: percentile 5 e percentile 95 (interpolazione lineare) calcolati sulla metrica pro capite giornaliera.
2. Le soglie sono stimate solo nel perimetro `pernottanti senza pacchetto`.
3. Nel calcolo soglie sono esclusi tutti i record con spesa categoria uguale a 0.
4. Gruppo principale: `stato_provenienza x motivazione_principale`.
5. Fallback soglie quando il gruppo e piccolo (`min_n = 30`):
  - solo provenienza
  - solo motivazione
  - totale campione
6. Flag outlier (solo per record nel perimetro):
  - `OUTLIER_BASSO` se valore < p5
  - `OUTLIER_ALTO` se valore > p95
  - `OK` se p5 <= valore <= p95
7. Record fuori perimetro:
  - non valutati per outlier
  - tracciati con livello dedicato `FUORI_PERIMETRO`

## Gestione pacchetti A/B/C
Regola implementata:
- Se `pacchetto` e A/B/C e `spesa_pacchetto` > 0, allora una spesa categoria uguale a 0 e trattata come zero strutturale.

Effetto della regola:
1. I record con pacchetto non rientrano nel perimetro analitico corrente outlier/imputazione.
2. Lo zero strutturale resta tracciato nei fogli di output come eccezione.

## Regole imputazione ND/NR
1. L'imputazione e applicata solo ai record nel perimetro `pernottanti senza pacchetto`.
2. Candidati donor: stesso `stato_provenienza` e stessa `motivazione_principale` (vincolo forte) e stesso perimetro operativo.
3. Scoring equivalenza: numero di match su 5 chiavi totali.
4. Regola primaria: almeno 3 match su 5.
5. Se ci sono piu donor, si selezionano quelli con score massimo e si usa la mediana del valore pro capite giornaliero donor.
6. Il valore imputato nelle colonne originali viene ricostruito come:

imputato_totale = mediana_donor_pcg * (numero_componenti * durata_soggiorno)

7. Se il denominatore target non e disponibile, fallback alla mediana donor dell'importo totale.
8. Fallback: se non esistono donor con >=3/5, si usa donor con vincolo forte (strong only).
9. Se non esiste donor compatibile, il valore resta mancante.

## Output prodotti
### 1) `outliers_da_verificare.xlsx`
Fogli:
- `OUTLIERS`: righe con almeno un outlier.
  - Celle spesa outlier evidenziate in giallo.
  - Colonne di dettaglio: tipo outlier, p5, p95, livello soglia, flag zero strutturale.
- `STATS`: statistiche iniziali globali per variabile spesa.
- `STATS_GRUPPI`: statistiche per gruppo `stato x motivazione`.
- `SOGLIE`: soglie calcolate ai diversi livelli di fallback.
- `REGOLE`: riepilogo regole applicate.

### 2) `questionari_imputazioni.xlsx`
Fogli:
- `DATI_IMPUTATI`: dataset completo con valori imputati inseriti.
  - Celle imputate evidenziate in verde.
- `LOG_IMPUTAZIONI`: audit delle imputazioni (id riga, colonna, valore pre/post, donor, regola).
- `STATS_POST_IMPUT`: statistiche globali ricalcolate dopo imputazione.

## Esecuzione
Dalla root del progetto:

```bash
./.venv/bin/python outlier_pipeline/run_outlier_pipeline.py
```

Con parametri espliciti:

```bash
./.venv/bin/python outlier_pipeline/run_outlier_pipeline.py \
  --input questionari_fonte.xlsx \
  --output-dir outlier_pipeline/output \
  --min-n 30
```

## Note metodologiche
- La plausibilita finale degli outlier resta manuale su `outliers_da_verificare.xlsx`.
- Questa versione non applica winsorizzazione automatica.
- La pipeline e progettata per supportare il controllo successivo sui soli pernottanti senza pacchetto.