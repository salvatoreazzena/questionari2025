from __future__ import annotations

from pathlib import Path

import pandas as pd

INPUT_FILE = Path("questionari_sardi_outlier_sostituiti.xlsx")
OUTPUT_FILE = Path("analisi_preliminare_campione_sardi.xlsx")

NULL_VALUE = "NULL"
UNDEFINED_LABEL = "NON DEFINITO"
NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}

REQUIRED_COLUMNS = [
    "luogo_somministrazione",
    "vacanza_2025",
    "si_dove",
    "fascia_età",
    "tipologia_turistica",
    "provincia_provenienza",
    "comune_provenienza",
    "numero_componenti",
]


def normalize_text_series(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("").str.strip()
    upper = s.str.upper()
    mask_null = upper.isin({t.upper() for t in NULL_TOKENS})
    return s.mask(mask_null, NULL_VALUE)


def normalize_upper_series(series: pd.Series) -> pd.Series:
    return normalize_text_series(series).str.upper()


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
    mostra_componenti: bool = True,
) -> None:
    stats = pd.DataFrame({"valore": source_series, "componenti": componenti_series})
    grouped = (
        stats.groupby("valore", dropna=False, as_index=False)
        .agg(questionari=("valore", "size"), componenti=("componenti", "sum"))
        .sort_values(by=["questionari", "valore"], ascending=[False, True], kind="mergesort")
    )

    for _, r in grouped.iterrows():
        raw = r["valore"]
        val = str(raw)
        if pd.isna(raw) or val.upper() == NULL_VALUE:
            val = UNDEFINED_LABEL

        q = int(r["questionari"])
        c = int(r["componenti"])
        rows.append(
            {
                "analisi_id": analisi_id,
                "analisi_nome": analisi_nome,
                "dimensione_1": dimensione_1,
                "mostra_componenti": mostra_componenti,
                "categoria": val,
                "questionari": q,
                "componenti": c if mostra_componenti else None,
                "pct_questionari_su_totale": (q / total_questionari * 100.0) if total_questionari > 0 else None,
                "pct_componenti_su_totale": ((c / total_componenti * 100.0) if total_componenti > 0 else None)
                if mostra_componenti
                else None,
            }
        )


def build_categoria_vacanza(vacanza_series: pd.Series, dove_series: pd.Series) -> pd.Series:
    vac = normalize_upper_series(vacanza_series)
    dove = normalize_upper_series(dove_series)
    out = pd.Series(UNDEFINED_LABEL, index=vac.index, dtype="string")

    is_si = vac.isin({"SI", "SÌ"})
    is_no = vac.isin({"NO"})
    out = out.mask(is_no, "NO VACANZA")

    out = out.mask(is_si & dove.eq("IN SARDEGNA"), "VACANZA IN SARDEGNA")
    out = out.mask(is_si & dove.eq("IN ITALIA"), "VACANZA IN ALTRE REGIONI ITALIANE")
    out = out.mask(is_si & dove.eq("ALL'ESTERO"), "VACANZA ALL'ESTERO")

    out = out.mask(is_si & out.eq(UNDEFINED_LABEL), "VACANZA (DESTINAZIONE NON DEFINITA)")
    return out


def normalize_comune_pre_virgola(series: pd.Series) -> pd.Series:
    s = normalize_text_series(series)
    s = s.str.split(",", n=1).str[0].str.strip()
    s = s.mask(s.eq(""), NULL_VALUE)
    return s


def build_visual_report(out: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict] = []
    ordered = out.sort_values(by=["analisi_id", "dimensione_1", "ordinamento"], kind="mergesort")

    section_keys = (
        ordered[["analisi_id", "dimensione_1"]]
        .drop_duplicates()
        .itertuples(index=False, name=None)
    )

    for analisi_id, dimensione_1 in section_keys:
        section = ordered[(ordered["analisi_id"] == analisi_id) & (ordered["dimensione_1"] == dimensione_1)]
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

        for _, r in section.iterrows():
            pct_q = "" if pd.isna(r["pct_questionari_su_totale"]) else float(r["pct_questionari_su_totale"])
            pct_c = "" if pd.isna(r["pct_componenti_su_totale"]) else float(r["pct_componenti_su_totale"])
            mostra_componenti = bool(r.get("mostra_componenti", True))
            rows.append(
                {
                    "Analisi": "",
                    "Categoria": str(r["categoria"]),
                    "Questionari": int(r["questionari"]),
                    "% Questionari su totale": pct_q,
                    "Componenti": ("" if not mostra_componenti or pd.isna(r["componenti"]) else int(r["componenti"])),
                    "% Componenti su totale": ("" if not mostra_componenti else pct_c),
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


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"File non trovato: {INPUT_FILE}")

    xls = pd.ExcelFile(INPUT_FILE)
    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(INPUT_FILE, sheet_name=first_sheet)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel dataset: {missing}")

    for col in REQUIRED_COLUMNS:
        if col == "numero_componenti":
            continue
        df[col] = normalize_text_series(df[col])

    df["numero_componenti_num"] = pd.to_numeric(df["numero_componenti"], errors="coerce").fillna(0)
    total_questionari = int(len(df))
    total_componenti = int(df["numero_componenti_num"].sum())
    questionari_componenti_validi = int(df["numero_componenti_num"].gt(0).sum())

    rows: list[dict] = []

    rows.append(
        {
            "analisi_id": 1,
            "analisi_nome": "Numero totale questionari",
            "dimensione_1": "",
            "mostra_componenti": True,
            "categoria": "Totale questionari",
            "questionari": total_questionari,
            "componenti": total_componenti,
            "pct_questionari_su_totale": None,
            "pct_componenti_su_totale": None,
        }
    )
    rows.append(
        {
            "analisi_id": 2,
            "analisi_nome": "Numero totale componenti",
            "dimensione_1": "",
            "mostra_componenti": True,
            "categoria": f"Totale componenti (somma su {questionari_componenti_validi} questionari con numero_componenti > 0)",
            "questionari": total_questionari,
            "componenti": total_componenti,
            "pct_questionari_su_totale": None,
            "pct_componenti_su_totale": None,
        }
    )

    vacanza_norm = normalize_upper_series(df["vacanza_2025"])
    is_vacanza = vacanza_norm.isin({"SI", "SÌ"})
    is_no_vacanza = vacanza_norm.eq("NO")

    add_distribution_rows(
        rows,
        analisi_id=3,
        analisi_nome="Distribuzione località di somministrazione (NO VACANZA)",
        source_series=df.loc[is_no_vacanza, "luogo_somministrazione"],
        componenti_series=df.loc[is_no_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="luogo_somministrazione_no_vacanza",
        mostra_componenti=False,
    )

    add_distribution_rows(
        rows,
        analisi_id=3,
        analisi_nome="Distribuzione località di somministrazione (SI VACANZA)",
        source_series=df.loc[is_vacanza, "luogo_somministrazione"],
        componenti_series=df.loc[is_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="luogo_somministrazione_si_vacanza",
        mostra_componenti=True,
    )

    categoria_vacanza = build_categoria_vacanza(df["vacanza_2025"], df["si_dove"])
    categoria_vacanza = categoria_vacanza.mask(categoria_vacanza.eq("NO VACANZA"), "NO VACANZA (solo questionari)")
    add_distribution_rows(
        rows,
        analisi_id=4,
        analisi_nome="Suddivisione questionari per categoria vacanza",
        source_series=categoria_vacanza,
        componenti_series=df["numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="categoria_vacanza",
    )

    add_distribution_rows(
        rows,
        analisi_id=5,
        analisi_nome="Distribuzione campione fascia d'età (fa vacanza)",
        source_series=df.loc[is_vacanza, "fascia_età"],
        componenti_series=df.loc[is_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="fascia_età_fa_vacanza",
        mostra_componenti=True,
    )

    add_distribution_rows(
        rows,
        analisi_id=5,
        analisi_nome="Distribuzione campione fascia d'età (non fa vacanza)",
        source_series=df.loc[is_no_vacanza, "fascia_età"],
        componenti_series=df.loc[is_no_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="fascia_età_non_vacanza",
        mostra_componenti=False,
    )

    add_distribution_rows(
        rows,
        analisi_id=6,
        analisi_nome="Distribuzione campione tipologia turistica (NO VACANZA)",
        source_series=df.loc[is_no_vacanza, "tipologia_turistica"],
        componenti_series=df.loc[is_no_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="tipologia_turistica_no_vacanza",
        mostra_componenti=False,
    )

    add_distribution_rows(
        rows,
        analisi_id=6,
        analisi_nome="Distribuzione campione tipologia turistica (SI VACANZA)",
        source_series=df.loc[is_vacanza, "tipologia_turistica"],
        componenti_series=df.loc[is_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="tipologia_turistica_si_vacanza",
        mostra_componenti=True,
    )

    add_distribution_rows(
        rows,
        analisi_id=8,
        analisi_nome="Composizione campione per provenienza provincia (NO VACANZA)",
        source_series=df.loc[is_no_vacanza, "provincia_provenienza"],
        componenti_series=df.loc[is_no_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="provincia_provenienza_no_vacanza",
        mostra_componenti=False,
    )

    add_distribution_rows(
        rows,
        analisi_id=8,
        analisi_nome="Composizione campione per provenienza provincia (SI VACANZA)",
        source_series=df.loc[is_vacanza, "provincia_provenienza"],
        componenti_series=df.loc[is_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="provincia_provenienza_si_vacanza",
        mostra_componenti=True,
    )

    comune_pre_virgola = normalize_comune_pre_virgola(df["comune_provenienza"])
    add_distribution_rows(
        rows,
        analisi_id=8,
        analisi_nome="Composizione campione per provenienza comune (NO VACANZA)",
        source_series=comune_pre_virgola.loc[is_no_vacanza],
        componenti_series=df.loc[is_no_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="comune_provenienza_no_vacanza",
        mostra_componenti=False,
    )

    add_distribution_rows(
        rows,
        analisi_id=8,
        analisi_nome="Composizione campione per provenienza comune (SI VACANZA)",
        source_series=comune_pre_virgola.loc[is_vacanza],
        componenti_series=df.loc[is_vacanza, "numero_componenti_num"],
        total_questionari=total_questionari,
        total_componenti=total_componenti,
        dimensione_1="comune_provenienza_si_vacanza",
        mostra_componenti=True,
    )

    out = pd.DataFrame(rows)
    out["pct_questionari_su_totale"] = out["pct_questionari_su_totale"].round(4)
    out["pct_componenti_su_totale"] = out["pct_componenti_su_totale"].round(4)
    out["ordinamento"] = out.groupby(["analisi_id", "dimensione_1"]).cumcount() + 1
    out = out.sort_values(by=["analisi_id", "dimensione_1", "ordinamento"], kind="mergesort")

    report = build_visual_report(out)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        report.to_excel(writer, sheet_name="risultati", index=False)

    print(f"Input letto: {INPUT_FILE} ({first_sheet})")
    print(f"Output generato: {OUTPUT_FILE}")
    print(f"Totale questionari: {total_questionari}")
    print(f"Totale componenti: {total_componenti}")


if __name__ == "__main__":
    main()
