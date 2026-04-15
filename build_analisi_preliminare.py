from __future__ import annotations

from pathlib import Path
import re

import pandas as pd

INPUT_FILE = Path("questionari_fonte.xlsx")
OUTPUT_FILE = Path("analisi_preliminare_campione.xlsx")
NULL_VALUE = "NULL"
UNDEFINED_LABEL = "NON DEFINITO"

NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}
CASE_INSENSITIVE_COLUMNS = {
    "luogo_somministrazione",
    "tipologia_sistemazione",
    "pacchetto",
    "fascia_età",
    "tipologia_turistica",
    "stato_provenienza",
    "regione_provenienza",
    "motivazione_principale",
}
REQUIRED_COLUMNS = [
    "numero_componenti",
    "luogo_somministrazione",
    "tipologia_sistemazione",
    "pacchetto",
    "fascia_età",
    "tipologia_turistica",
    "stato_provenienza",
    "regione_provenienza",
    "località_visitate",
    "motivazione_principale",
]


def normalize_text_series(series: pd.Series) -> pd.Series:
    s = series.astype("string")
    s = s.fillna("").str.strip()
    upper = s.str.upper()
    mask_null = upper.isin({token.upper() for token in NULL_TOKENS})
    s = s.mask(mask_null, NULL_VALUE)
    return s


def normalize_case_insensitive_series(series: pd.Series) -> pd.Series:
    # Canonicalizza il testo in maiuscolo per evitare split di categoria dovuti al case.
    s = series.astype("string").fillna("").str.strip()
    return s.str.upper()


def add_distribution_rows(
    rows: list[dict],
    *,
    analisi_id: int,
    analisi_nome: str,
    source_series: pd.Series,
    componenti_series: pd.Series,
    total_questionari: int,
    total_componenti: int,
    dimensione_1: str,
    dimensione_2: str = "",
    sort_by: str = "questionari",
    sort_desc: bool = True,
    limit: int | None = None,
) -> None:
    stats = pd.DataFrame({"valore": source_series, "componenti": componenti_series})
    counts_df = (
        stats.groupby("valore", dropna=False, as_index=False)
        .agg(questionari=("valore", "size"), componenti=("componenti", "sum"))
    )
    if sort_desc:
        metric_col = "componenti" if sort_by == "componenti" else "questionari"
        counts_df = counts_df.sort_values(by=[metric_col, "valore"], ascending=[False, True], kind="mergesort")
    if limit is not None:
        counts_df = counts_df.head(limit)

    for _, row in counts_df.iterrows():
        value = row["valore"]
        questionari_int = int(row["questionari"])
        componenti_int = int(row["componenti"])

        pct_q_total = (questionari_int / total_questionari * 100.0) if total_questionari > 0 else None
        pct_c_total = (componenti_int / total_componenti * 100.0) if total_componenti > 0 else None

        value_str = str(value)
        if pd.isna(value) or value_str.upper() == NULL_VALUE:
            value_str = UNDEFINED_LABEL

        rows.append(
            {
                "analisi_id": analisi_id,
                "analisi_nome": analisi_nome,
                "metrica": "count",
                "dimensione_1": dimensione_1,
                "dimensione_2": dimensione_2,
                "valore_categoria": value_str,
                "questionari": questionari_int,
                "componenti": componenti_int,
                "pct_questionari_su_totale": pct_q_total,
                "pct_componenti_su_totale": pct_c_total,
            }
        )


def to_numeric_sum(series: pd.Series) -> int:
    numeric = pd.to_numeric(series, errors="coerce")
    return int(numeric.fillna(0).sum())


def extract_prevalent_destination(value: object) -> str:
    if pd.isna(value):
        return NULL_VALUE

    raw_text = str(value).strip()
    if not raw_text or raw_text == NULL_VALUE:
        return NULL_VALUE

    # Alcuni questionari racchiudono tutta la lista tra parentesi esterne.
    if raw_text.startswith("(") and raw_text.endswith(")"):
        raw_text = raw_text[1:-1].strip()

    if not raw_text:
        return NULL_VALUE

    days_by_destination: dict[str, int] = {}
    first_position: dict[str, int] = {}
    pair_pattern = re.compile(r"\s*([^,;][^,;]*?)\s*,\s*([0-9]+(?:[.,][0-9]+)?)\s*(?=;|,|$)")

    # Ogni coppia finisce sul numero giorni; accettiamo sia ';' sia ',' come separatore tra coppie.
    matches = list(pair_pattern.finditer(raw_text))
    for idx, match in enumerate(matches):
        destination_raw = match.group(1).strip()
        days_raw = match.group(2).strip()

        # Esclude il testo tra parentesi nel nome localita (frazione/area).
        destination_clean = re.sub(r"\s*\([^)]*\)\s*", " ", destination_raw)
        # Rimuove eventuali prefissi di indice dovuti a formattazioni errate (es. "0: PALAU").
        destination_clean = re.sub(r"^(?:\s*\d+\s*:\s*)+", "", destination_clean)
        destination_clean = re.sub(r"\s+", " ", destination_clean).strip()
        if not destination_clean:
            continue

        try:
            days_float = float(days_raw.replace(",", "."))
        except ValueError:
            continue

        if not days_float.is_integer():
            continue

        days_int = int(days_float)
        if days_int <= 0:
            continue

        destination_key = destination_clean.upper()
        if destination_key not in first_position:
            first_position[destination_key] = idx
        days_by_destination[destination_key] = days_by_destination.get(destination_key, 0) + days_int

    if not days_by_destination:
        return NULL_VALUE

    prevalent_destination = min(
        days_by_destination.keys(),
        key=lambda k: (-days_by_destination[k], first_position[k], k),
    )
    return prevalent_destination


def build_visual_report(out: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict] = []

    ordered = out.sort_values(by=["analisi_id", "ordinamento", "valore_categoria"], kind="mergesort")
    for analisi_id in sorted(ordered["analisi_id"].unique().tolist()):
        section = ordered[ordered["analisi_id"] == analisi_id]
        analisi_nome = str(section["analisi_nome"].iloc[0])

        rows.append(
            {
                "Analisi": f"{analisi_id}. {analisi_nome}",
                "Categoria": "",
                "Questionari": "",
                "% Questionari su totale": "",
                "Componenti": "",
                "% Componenti su totale": "",
            }
        )

        if analisi_id == 8:
            states = section[section["dimensione_2"] == "paesi_esteri"].sort_values(
                by=["ordinamento", "valore_categoria"],
                kind="mergesort",
            )
            regions = section[section["dimensione_2"] == "regioni_italiane"].sort_values(
                by=["ordinamento", "valore_categoria"],
                kind="mergesort",
            )

            for _, r in states.iterrows():
                pct_q_total = "" if pd.isna(r["pct_questionari_su_totale"]) else float(r["pct_questionari_su_totale"])
                pct_c_total = "" if pd.isna(r["pct_componenti_su_totale"]) else float(r["pct_componenti_su_totale"])
                rows.append(
                    {
                        "Analisi": "",
                        "Categoria": str(r["valore_categoria"]),
                        "Questionari": int(r["questionari"]),
                        "% Questionari su totale": pct_q_total,
                        "Componenti": int(r["componenti"]),
                        "% Componenti su totale": pct_c_total,
                    }
                )

            rows.append(
                {
                    "Analisi": "",
                    "Categoria": "",
                    "Questionari": "",
                    "% Questionari su totale": "",
                    "Componenti": "",
                    "% Componenti su totale": "",
                }
            )

            for _, r in regions.iterrows():
                pct_q_total = "" if pd.isna(r["pct_questionari_su_totale"]) else float(r["pct_questionari_su_totale"])
                pct_c_total = "" if pd.isna(r["pct_componenti_su_totale"]) else float(r["pct_componenti_su_totale"])
                rows.append(
                    {
                        "Analisi": "",
                        "Categoria": str(r["valore_categoria"]),
                        "Questionari": int(r["questionari"]),
                        "% Questionari su totale": pct_q_total,
                        "Componenti": int(r["componenti"]),
                        "% Componenti su totale": pct_c_total,
                    }
                )
        else:
            for _, r in section.iterrows():
                pct_q_total = "" if pd.isna(r["pct_questionari_su_totale"]) else float(r["pct_questionari_su_totale"])
                pct_c_total = "" if pd.isna(r["pct_componenti_su_totale"]) else float(r["pct_componenti_su_totale"])

                rows.append(
                    {
                        "Analisi": "",
                        "Categoria": str(r["valore_categoria"]),
                        "Questionari": int(r["questionari"]),
                        "% Questionari su totale": pct_q_total,
                        "Componenti": int(r["componenti"]),
                        "% Componenti su totale": pct_c_total,
                    }
                )

        rows.append(
            {
                "Analisi": "",
                "Categoria": "",
                "Questionari": "",
                "% Questionari su totale": "",
                "Componenti": "",
                "% Componenti su totale": "",
            }
        )
        rows.append(
            {
                "Analisi": "",
                "Categoria": "",
                "Questionari": "",
                "% Questionari su totale": "",
                "Componenti": "",
                "% Componenti su totale": "",
            }
        )

    return pd.DataFrame(
        rows,
        columns=[
            "Analisi",
            "Categoria",
            "Questionari",
            "% Questionari su totale",
            "Componenti",
            "% Componenti su totale",
        ],
    )


def build_analysis_table(df: pd.DataFrame, *, include_analysis_7: bool) -> tuple[pd.DataFrame, int, int]:
    total_questionari = int(len(df))
    total_componenti = int(df["numero_componenti_num"].sum())
    rows: list[dict] = []

    # 1) Numero totale questionari
    rows.append(
        {
            "analisi_id": 1,
            "analisi_nome": "Numero totale questionari",
            "metrica": "totale",
            "dimensione_1": "",
            "dimensione_2": "",
            "valore_categoria": "Totale questionari",
            "questionari": total_questionari,
            "componenti": total_componenti,
            "pct_questionari_su_totale": None,
            "pct_componenti_su_totale": None,
        }
    )

    # 2) Numero totale componenti
    rows.append(
        {
            "analisi_id": 2,
            "analisi_nome": "Numero totale componenti",
            "metrica": "totale",
            "dimensione_1": "",
            "dimensione_2": "",
            "valore_categoria": "Totale componenti",
            "questionari": total_questionari,
            "componenti": total_componenti,
            "pct_questionari_su_totale": None,
            "pct_componenti_su_totale": None,
        }
    )

    # 3) Distribuzione localita di somministrazione
    add_distribution_rows(
        rows,
        analisi_id=3,
        analisi_nome="Distribuzione localita di somministrazione",
        source_series=df["luogo_somministrazione"],
        componenti_series=df["numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="luogo_somministrazione",
    )

    # 4) Suddivisione: pernottanti (con/senza pacchetto), escursionisti, crocieristi
    sist = df["tipologia_sistemazione"]
    pacchetto = normalize_text_series(df["pacchetto"])

    is_crocierista = sist.str.upper() == "NAVE DA CROCIERA"
    is_con_pacchetto = pacchetto.str.upper().isin({"A", "B", "C"})

    n_crocieristi = int(is_crocierista.sum())
    n_pernottanti_con = int((~is_crocierista & is_con_pacchetto).sum())
    n_pernottanti_senza = int((~is_crocierista & ~is_con_pacchetto).sum())
    n_escursionisti = 0

    c_crocieristi = int(df.loc[is_crocierista, "numero_componenti_num"].sum())
    c_pernottanti_con = int(df.loc[(~is_crocierista & is_con_pacchetto), "numero_componenti_num"].sum())
    c_pernottanti_senza = int(df.loc[(~is_crocierista & ~is_con_pacchetto), "numero_componenti_num"].sum())
    c_escursionisti = 0

    cat4 = [
        ("Pernottanti - Con pacchetto", n_pernottanti_con, c_pernottanti_con),
        ("Pernottanti - Senza pacchetto", n_pernottanti_senza, c_pernottanti_senza),
        ("Escursionisti", n_escursionisti, c_escursionisti),
        ("Crocieristi", n_crocieristi, c_crocieristi),
    ]
    for label, questionari_int, componenti_int in cat4:
        rows.append(
            {
                "analisi_id": 4,
                "analisi_nome": "Suddivisione questionari per categoria",
                "metrica": "count",
                "dimensione_1": "categoria_campione",
                "dimensione_2": "",
                "valore_categoria": label,
                "questionari": questionari_int,
                "componenti": componenti_int,
                "pct_questionari_su_totale": (questionari_int / total_questionari * 100.0)
                if total_questionari > 0
                else None,
                "pct_componenti_su_totale": (componenti_int / total_componenti * 100.0)
                if total_componenti > 0
                else None,
            }
        )

    # 5) Distribuzione campione fascia d'eta
    add_distribution_rows(
        rows,
        analisi_id=5,
        analisi_nome="Distribuzione campione fascia d'eta",
        source_series=df["fascia_età"],
        componenti_series=df["numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="fascia_età",
    )

    # 6) Distribuzione campione tipologia turistica
    add_distribution_rows(
        rows,
        analisi_id=6,
        analisi_nome="Distribuzione campione tipologia turistica",
        source_series=df["tipologia_turistica"],
        componenti_series=df["numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="tipologia_turistica",
    )

    naz = df["stato_provenienza"]

    # 7) Composizione italiani/stranieri (opzionale per report filtrati)
    if include_analysis_7:
        composizione = pd.Series(
            pd.NA,
            index=df.index,
            dtype="string",
        )
        composizione = composizione.mask(naz == NULL_VALUE, NULL_VALUE)
        composizione = composizione.mask(naz.str.upper() == "ITALIA", "Italiani")
        composizione = composizione.mask((naz != NULL_VALUE) & (naz.str.upper() != "ITALIA"), "Stranieri")

        add_distribution_rows(
            rows,
            analisi_id=7,
            analisi_nome="Composizione campione italiani/stranieri",
            source_series=composizione,
            componenti_series=df["numero_componenti_num"],
            total_questionari=total_questionari,
            total_componenti=total_componenti,
            dimensione_1="composizione_campione",
        )

    # 8) Provenienza: distribuzione completa paesi esteri e regioni italiane (ranking su componenti)
    paesi_esteri = naz[naz.str.upper() != "ITALIA"]
    add_distribution_rows(
        rows,
        analisi_id=8,
        analisi_nome="Composizione campione per provenienza",
        source_series=paesi_esteri,
        componenti_series=df.loc[paesi_esteri.index, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="stato_provenienza",
        dimensione_2="paesi_esteri",
        sort_by="componenti",
    )

    regioni = df["regione_provenienza"]
    regioni_ita = regioni[naz.str.upper() == "ITALIA"]
    add_distribution_rows(
        rows,
        analisi_id=8,
        analisi_nome="Composizione campione per provenienza",
        source_series=regioni_ita,
        componenti_series=df.loc[regioni_ita.index, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="regione_provenienza",
        dimensione_2="regioni_italiane",
        sort_by="componenti",
    )

    # 9) Distribuzione motivazioni principali
    add_distribution_rows(
        rows,
        analisi_id=9,
        analisi_nome="Distribuzione motivazioni principali",
        source_series=df["motivazione_principale"],
        componenti_series=df["numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="motivazione_principale",
    )

    # 10) Distribuzione destinazione prevalente (estratta da localita_visitate)
    destinazione_prevalente = df["località_visitate"].apply(extract_prevalent_destination)
    add_distribution_rows(
        rows,
        analisi_id=10,
        analisi_nome="Distribuzione destinazione prevalente",
        source_series=destinazione_prevalente,
        componenti_series=df["numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="destinazione_prevalente",
        sort_by="componenti",
    )

    out = pd.DataFrame(rows)
    out["ordinamento"] = out.groupby(["analisi_id", "dimensione_1", "dimensione_2"]).cumcount() + 1

    out["pct_questionari_su_totale"] = out["pct_questionari_su_totale"].round(4)
    out["pct_componenti_su_totale"] = out["pct_componenti_su_totale"].round(4)

    out = out[
        [
            "analisi_id",
            "analisi_nome",
            "metrica",
            "dimensione_1",
            "dimensione_2",
            "valore_categoria",
            "questionari",
            "componenti",
            "pct_questionari_su_totale",
            "pct_componenti_su_totale",
            "ordinamento",
        ]
    ]

    return out, total_questionari, total_componenti


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"File non trovato: {INPUT_FILE}")

    xls = pd.ExcelFile(INPUT_FILE)
    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(INPUT_FILE, sheet_name=first_sheet)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel file sorgente: {missing}")

    for col in REQUIRED_COLUMNS:
        if col == "numero_componenti":
            continue
        df[col] = normalize_text_series(df[col])
        if col in CASE_INSENSITIVE_COLUMNS:
            df[col] = normalize_case_insensitive_series(df[col])

    df["numero_componenti_num"] = pd.to_numeric(df["numero_componenti"], errors="coerce").fillna(0)

    out_totale, total_questionari, total_componenti = build_analysis_table(df, include_analysis_7=True)
    report_totale = build_visual_report(out_totale)

    naz = df["stato_provenienza"]
    df_italiani = df[naz.str.upper() == "ITALIA"].copy()
    df_stranieri = df[(naz != NULL_VALUE) & (naz.str.upper() != "ITALIA")].copy()

    out_italiani, _, _ = build_analysis_table(df_italiani, include_analysis_7=False)
    out_stranieri, _, _ = build_analysis_table(df_stranieri, include_analysis_7=False)
    report_italiani = build_visual_report(out_italiani)
    report_stranieri = build_visual_report(out_stranieri)

    saved_path = OUTPUT_FILE
    try:
        with pd.ExcelWriter(saved_path, engine="openpyxl") as writer:
            report_totale.to_excel(writer, sheet_name="risultati", index=False)
            report_italiani.to_excel(writer, sheet_name="ITALIANI", index=False)
            report_stranieri.to_excel(writer, sheet_name="STRANIERI", index=False)
    except PermissionError:
        saved_path = OUTPUT_FILE.with_name(f"{OUTPUT_FILE.stem}_nuovo{OUTPUT_FILE.suffix}")
        with pd.ExcelWriter(saved_path, engine="openpyxl") as writer:
            report_totale.to_excel(writer, sheet_name="risultati", index=False)
            report_italiani.to_excel(writer, sheet_name="ITALIANI", index=False)
            report_stranieri.to_excel(writer, sheet_name="STRANIERI", index=False)
        print(
            "File di output originale in uso: salvato su file alternativo "
            f"{saved_path}"
        )

    print(f"Output generato: {saved_path}")
    print(f"Righe report totale: {len(report_totale)}")
    print(f"Righe report ITALIANI: {len(report_italiani)}")
    print(f"Righe report STRANIERI: {len(report_stranieri)}")
    print(f"Totale questionari: {total_questionari}")
    print(f"Totale componenti: {total_componenti}")


if __name__ == "__main__":
    main()
