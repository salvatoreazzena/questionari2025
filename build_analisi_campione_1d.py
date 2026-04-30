from __future__ import annotations

from pathlib import Path
import re

import pandas as pd

INPUT_FILE = Path("questionari_sottocampioni.xlsx")
INPUT_SHEET = "campione_1d"
OUTPUT_FILE = Path("analisi_campione_1d_motivazione_apprezzamento.xlsx")

NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A.", "NON DISPONIBILE", "NULL"}


def normalize_text_series(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("").str.strip()
    mask_null = s.str.upper().isin({t.upper() for t in NULL_TOKENS})
    return s.mask(mask_null, "")


def split_apprezzamenti(text: str) -> list[str]:
    parts = re.split(r"[;,/|]+", text)
    out: list[str] = []
    for p in parts:
        cleaned = re.sub(r"\s+", " ", p).strip(" .:-")
        if not cleaned:
            continue
        if cleaned.upper() in {t.upper() for t in NULL_TOKENS}:
            continue
        out.append(cleaned.title())
    return out


def build_nazionalita(stato: pd.Series) -> pd.Series:
    s = normalize_text_series(stato).str.upper()
    out = pd.Series("NON DEFINITO", index=s.index, dtype="string")
    out = out.mask(s.eq("ITALIA"), "ITALIANI")
    out = out.mask((s.ne("")) & (s.ne("ITALIA")), "STRANIERI")
    return out


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"File non trovato: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)
    required = ["stato_provenienza", "motivazione_principale", "apprezzamenti", "numero_componenti"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti: {missing}")

    df["motivazione_principale"] = normalize_text_series(df["motivazione_principale"]).str.upper()
    df["apprezzamenti"] = normalize_text_series(df["apprezzamenti"])
    df["nazionalita_grp"] = build_nazionalita(df["stato_provenienza"])
    df["numero_componenti_num"] = pd.to_numeric(df["numero_componenti"], errors="coerce").fillna(0)

    working = df[(df["motivazione_principale"] != "") & (df["apprezzamenti"] != "")].copy()

    rows: list[dict] = []
    for _, r in working.iterrows():
        elementi = split_apprezzamenti(str(r["apprezzamenti"]))
        if not elementi:
            continue
        for e in elementi:
            rows.append(
                {
                    "nazionalita": r["nazionalita_grp"],
                    "motivazione_principale": r["motivazione_principale"],
                    "elemento_apprezzamento": e,
                    "questionari": 1,
                    "componenti": float(r["numero_componenti_num"]),
                }
            )

    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("Nessun apprezzamento valido trovato nel campione_1d.")

    grouped = (
        out.groupby(["nazionalita", "motivazione_principale", "elemento_apprezzamento"], as_index=False)
        .agg(questionari=("questionari", "sum"), componenti=("componenti", "sum"))
    )
    grouped["componenti"] = grouped["componenti"].round(0).astype(int)

    grouped = grouped.sort_values(
        by=["nazionalita", "motivazione_principale", "questionari", "componenti", "elemento_apprezzamento"],
        ascending=[True, True, False, False, True],
        kind="mergesort",
    )
    grouped["rank_nel_gruppo"] = grouped.groupby(["nazionalita", "motivazione_principale"]).cumcount() + 1

    top10 = grouped[grouped["rank_nel_gruppo"] <= 10].copy()

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        top10.to_excel(writer, sheet_name="top_apprezzamenti", index=False)
        grouped.to_excel(writer, sheet_name="dettaglio_completo", index=False)

    print(f"Input letto: {INPUT_FILE} [{INPUT_SHEET}]")
    print(f"Output generato: {OUTPUT_FILE}")
    print(f"Righe top_apprezzamenti: {len(top10)}")
    print(f"Righe dettaglio_completo: {len(grouped)}")


if __name__ == "__main__":
    main()
