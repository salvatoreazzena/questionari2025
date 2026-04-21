from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

DEFAULT_INPUT = Path("questionari_fonte_veri_outlier_sostituiti_imputati.xlsx")
DEFAULT_OUTPUT = Path("questionari_sottocampioni.xlsx")

ID_COL = "ID"
DURATION_COL = "durata_soggiorno"
PACCHETTO_COL = "pacchetto"
COMPONENTI_COL = "numero_componenti"
ETA_COL = "fascia_età"
STATO_COL = "stato_provenienza"
REGIONE_COL = "regione_provenienza"

NULL_VALUE = "NULL"
NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}

REQUIRED_COLUMNS = [
    DURATION_COL,
    PACCHETTO_COL,
    COMPONENTI_COL,
    ETA_COL,
    STATO_COL,
    REGIONE_COL,
]


@dataclass(frozen=True)
class CellPlan:
    group_value: str
    fascia_eta: str
    available_components: float
    available_questionari: int


def is_strict_empty(value: object) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def normalize_text_series(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("").str.strip()
    upper = s.str.upper()
    mask_null = upper.isin({token.upper() for token in NULL_TOKENS})
    return s.mask(mask_null, NULL_VALUE)


def normalize_case_insensitive_series(series: pd.Series) -> pd.Series:
    return series.astype("string").fillna("").str.strip().str.upper()


def build_macro_provenienza(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("").str.strip().str.upper()
    out = pd.Series(pd.NA, index=s.index, dtype="string")
    out = out.mask(s == "ITALIA", "ITALIANI")
    out = out.mask((s != "") & (s != NULL_VALUE) & (s != "ITALIA"), "STRANIERI")
    return out


def validate_columns(df: pd.DataFrame) -> None:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel file sorgente: {missing}")


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    out[ETA_COL] = normalize_case_insensitive_series(normalize_text_series(out[ETA_COL]))
    out[STATO_COL] = normalize_case_insensitive_series(normalize_text_series(out[STATO_COL]))
    out[REGIONE_COL] = normalize_case_insensitive_series(normalize_text_series(out[REGIONE_COL]))

    out["_durata_num"] = pd.to_numeric(out[DURATION_COL], errors="coerce")
    out["_componenti_num"] = pd.to_numeric(out[COMPONENTI_COL], errors="coerce")

    out["_row_order"] = out.index.astype("int64")
    if ID_COL in out.columns:
        out["_id_text"] = out[ID_COL].astype("string").fillna("")
    else:
        out["_id_text"] = (out.index + 2).astype("string")

    out["macro_provenienza"] = build_macro_provenienza(out[STATO_COL])

    mask = (
        (out["_durata_num"] > 0)
        & out[PACCHETTO_COL].map(is_strict_empty)
        & (out[ETA_COL] != NULL_VALUE)
        & out["_componenti_num"].notna()
        & (out["_componenti_num"] > 0)
    )

    return out.loc[mask].copy()


def format_number(value: float) -> int | float:
    rounded = round(float(value), 6)
    as_int = int(round(rounded))
    if abs(rounded - as_int) < 1e-9:
        return as_int
    return rounded


def compute_group_ranking(df: pd.DataFrame, group_col: str, top_n: int) -> list[str]:
    stats = (
        df.groupby(group_col, dropna=False, as_index=False)
        .agg(componenti=("_componenti_num", "sum"), questionari=(group_col, "size"))
        .sort_values(by=["componenti", group_col], ascending=[False, True], kind="mergesort")
    )

    values = stats[group_col].astype("string").tolist()
    if len(values) < top_n:
        raise ValueError(
            f"Categorie insufficienti per top {top_n} su {group_col}: trovate {len(values)}."
        )
    return [str(v) for v in values[:top_n]]


def compute_age_bands_common(df: pd.DataFrame, *, group_col: str, group_values: list[str]) -> list[str]:
    age_sets: list[set[str]] = []
    for group in group_values:
        group_df = df[df[group_col] == group]
        grouped = group_df.groupby(ETA_COL, as_index=False).agg(componenti=("_componenti_num", "sum"))
        valid_ages = {str(r[ETA_COL]) for _, r in grouped.iterrows() if float(r["componenti"]) > 0}
        if not valid_ages:
            raise ValueError(f"Nessuna fascia eta disponibile per gruppo {group}.")
        age_sets.append(valid_ages)

    common = set.intersection(*age_sets)
    if not common:
        raise ValueError(
            "Nessuna fascia eta comune tra tutti i gruppi richiesti; campione non costruibile."
        )
    return sorted(common)


def build_cell_plan(
    df: pd.DataFrame,
    *,
    group_col: str,
    group_values: list[str],
    age_bands: list[str],
) -> list[CellPlan]:
    plans: list[CellPlan] = []
    for group in group_values:
        for age_band in age_bands:
            cell = df[(df[group_col] == group) & (df[ETA_COL] == age_band)]
            available_components = float(cell["_componenti_num"].sum())
            available_questionari = int(len(cell))
            plans.append(
                CellPlan(
                    group_value=group,
                    fascia_eta=age_band,
                    available_components=available_components,
                    available_questionari=available_questionari,
                )
            )
    return plans


def is_integer_like(value: float, *, tol: float = 1e-9) -> bool:
    return abs(float(value) - round(float(value))) <= tol


def select_cell_rows_with_optimal_deviation(cell_df: pd.DataFrame, target_components: float) -> tuple[pd.Index, float]:
    if target_components <= 0:
        return pd.Index([], dtype="int64"), 0.0

    ordered = cell_df.sort_values(
        by=["_componenti_num", "_id_text", "_row_order"],
        ascending=[True, True, True],
        kind="mergesort",
    )

    weights_raw = pd.to_numeric(ordered["_componenti_num"], errors="coerce")
    if weights_raw.isna().any():
        return select_cell_rows(cell_df, target_components)

    weights_list = [float(v) for v in weights_raw.tolist()]
    if not is_integer_like(target_components) or any(not is_integer_like(v) for v in weights_list):
        return select_cell_rows(cell_df, target_components)

    weights = [int(round(v)) for v in weights_list]
    target_int = int(round(float(target_components)))
    total_sum = int(sum(weights))

    # Limite di sicurezza per evitare costi eccessivi in casi patologici.
    if total_sum > 50000:
        return select_cell_rows(cell_df, target_components)

    reachable = bytearray(total_sum + 1)
    prev_sum = [-1] * (total_sum + 1)
    prev_idx = [-1] * (total_sum + 1)
    reachable[0] = 1

    for i, w in enumerate(weights):
        if w <= 0:
            continue
        for s in range(total_sum - w, -1, -1):
            if reachable[s] and not reachable[s + w]:
                reachable[s + w] = 1
                prev_sum[s + w] = s
                prev_idx[s + w] = i

    best_sum: int | None = None
    best_key: tuple[int, int, int] | None = None
    for s in range(1, total_sum + 1):
        if not reachable[s]:
            continue
        diff = s - target_int
        key = (abs(diff), max(diff, 0), s)
        if best_key is None or key < best_key:
            best_key = key
            best_sum = s

    if best_sum is None:
        return pd.Index([], dtype="int64"), 0.0

    selected_positions: list[int] = []
    cur = best_sum
    while cur > 0:
        i = prev_idx[cur]
        if i < 0:
            # Fallback robusto: in caso di stato incoerente usa selezione greedily.
            return select_cell_rows(cell_df, target_components)
        selected_positions.append(i)
        cur = prev_sum[cur]

    selected_positions.reverse()
    selected_rows = ordered.iloc[selected_positions]
    selected_idx = pd.Index(selected_rows.index.astype("int64"))
    return selected_idx, float(best_sum)


def select_cell_rows(cell_df: pd.DataFrame, target_components: float) -> tuple[pd.Index, float]:
    if target_components <= 0:
        return pd.Index([], dtype="int64"), 0.0

    ordered = cell_df.sort_values(
        by=["_componenti_num", "_id_text", "_row_order"],
        ascending=[True, True, True],
        kind="mergesort",
    )

    cumulative = 0.0
    cumulative_before = 0.0
    selected_indices: list[int] = []

    for idx, row in ordered.iterrows():
        cumulative_before = cumulative
        cumulative += float(row["_componenti_num"])
        selected_indices.append(int(idx))
        if cumulative >= target_components:
            break

    # Fallback: se il prefisso precedente e piu vicino al target, preferiscilo.
    if selected_indices:
        if abs(cumulative_before - target_components) <= abs(cumulative - target_components):
            selected_indices = selected_indices[:-1]
            cumulative = cumulative_before

    return pd.Index(selected_indices), cumulative


def extract_sample(
    df: pd.DataFrame,
    *,
    sample_name: str,
    group_col: str,
    group_values: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    age_bands = compute_age_bands_common(df, group_col=group_col, group_values=group_values)
    cell_plan = build_cell_plan(
        df,
        group_col=group_col,
        group_values=group_values,
        age_bands=age_bands,
    )

    target_components = min(cell.available_components for cell in cell_plan)
    if target_components <= 0:
        raise ValueError(f"Target non valido per {sample_name}: {target_components}")

    selected_indices_all: list[pd.Index] = []
    selection_rows: list[dict[str, object]] = []

    for cell in cell_plan:
        cell_df = df[(df[group_col] == cell.group_value) & (df[ETA_COL] == cell.fascia_eta)]
        if cell.available_components < target_components - 1e-9:
            raise ValueError(
                f"Capacita insufficiente per {sample_name} nella cella {cell.group_value} x {cell.fascia_eta}."
            )

        selected_idx, selected_components = select_cell_rows_with_optimal_deviation(cell_df, target_components)
        selected_indices_all.append(selected_idx)

        selection_rows.append(
            {
                "campione": sample_name,
                "dimensione_gruppo": group_col,
                "valore_gruppo": cell.group_value,
                "fascia_eta": cell.fascia_eta,
                "componenti_disponibili": format_number(cell.available_components),
                "questionari_disponibili": cell.available_questionari,
                "target_componenti": format_number(target_components),
                "componenti_selezionati": format_number(selected_components),
                "overshoot_componenti": format_number(selected_components - target_components),
                "questionari_selezionati": int(len(selected_idx)),
            }
        )

    if selected_indices_all:
        final_idx = pd.Index([], dtype="int64")
        for idx in selected_indices_all:
            final_idx = final_idx.append(idx)
    else:
        final_idx = pd.Index([], dtype="int64")

    selected = df.loc[final_idx].copy()
    selected["campione"] = sample_name
    selected["dimensione_gruppo"] = group_col

    selected[group_col] = selected[group_col].astype("string")
    selected["strato_gruppo"] = selected[group_col].astype("string")
    selected["strato_fascia_eta"] = selected[ETA_COL].astype("string")

    selected = selected.sort_values(by=["strato_gruppo", "strato_fascia_eta", "_id_text", "_row_order"], kind="mergesort")

    selection_df = pd.DataFrame(selection_rows)
    meta = {
        "campione": sample_name,
        "n_gruppi": len(group_values),
        "n_fasce_eta": len(age_bands),
        "target_componenti_per_cella": format_number(target_components),
        "componenti_totali_selezionati": format_number(float(selected["_componenti_num"].sum())),
        "questionari_totali_selezionati": int(len(selected)),
    }
    return selected, selection_df, meta


def build_sample_1(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[df["macro_provenienza"].isin(["ITALIANI", "STRANIERI"])].copy()
    if subset.empty:
        raise ValueError("Campione 1 non costruibile: nessuna riga ITALIANI/STRANIERI.")

    group_values = ["ITALIANI", "STRANIERI"]
    selected, cell_summary, meta = extract_sample(
        subset,
        sample_name="Campione_1_50_50",
        group_col="macro_provenienza",
        group_values=group_values,
    )

    selected_components = (
        selected.groupby("macro_provenienza", as_index=False)
        .agg(componenti=("_componenti_num", "sum"))
        .set_index("macro_provenienza")["componenti"]
    )
    italiani = float(selected_components.get("ITALIANI", 0.0))
    stranieri = float(selected_components.get("STRANIERI", 0.0))
    meta["componenti_italiani"] = format_number(italiani)
    meta["componenti_stranieri"] = format_number(stranieri)
    meta["rapporto_italiani_su_totale"] = format_number((italiani / (italiani + stranieri) * 100.0) if (italiani + stranieri) > 0 else 0.0)

    return selected, cell_summary, meta


def build_sample_2(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[(df[STATO_COL] != NULL_VALUE) & (df[STATO_COL] != "ITALIA")].copy()
    if subset.empty:
        raise ValueError("Campione 2 non costruibile: nessun paese estero disponibile.")

    top5 = compute_group_ranking(subset, STATO_COL, top_n=5)
    subset = subset[subset[STATO_COL].isin(top5)].copy()

    selected, cell_summary, meta = extract_sample(
        subset,
        sample_name="Campione_2_top5_paesi",
        group_col=STATO_COL,
        group_values=top5,
    )
    meta["top5"] = ", ".join(top5)
    return selected, cell_summary, meta


def build_sample_3(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[(df[STATO_COL] == "ITALIA") & (df[REGIONE_COL] != NULL_VALUE)].copy()
    if subset.empty:
        raise ValueError("Campione 3 non costruibile: nessuna regione italiana disponibile.")

    top5 = compute_group_ranking(subset, REGIONE_COL, top_n=5)
    subset = subset[subset[REGIONE_COL].isin(top5)].copy()

    selected, cell_summary, meta = extract_sample(
        subset,
        sample_name="Campione_3_top5_regioni",
        group_col=REGIONE_COL,
        group_values=top5,
    )
    meta["top5"] = ", ".join(top5)
    return selected, cell_summary, meta


def split_output_columns(df_selected: pd.DataFrame, original_columns: list[str]) -> pd.DataFrame:
    extra_cols = [
        "campione",
        "dimensione_gruppo",
        "strato_gruppo",
        "strato_fascia_eta",
    ]
    existing_original = [c for c in original_columns if c in df_selected.columns]
    return df_selected[extra_cols + existing_original].copy()


def run(input_file: Path, output_file: Path) -> Path:
    if not input_file.exists():
        raise FileNotFoundError(f"File input non trovato: {input_file}")

    xls = pd.ExcelFile(input_file)
    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(input_file, sheet_name=first_sheet)

    validate_columns(df)
    original_columns = df.columns.tolist()

    eligible = prepare_dataframe(df)
    if eligible.empty:
        raise ValueError("Nessun record eleggibile dopo i filtri base.")

    sample_1_df, sample_1_diag, sample_1_meta = build_sample_1(eligible)
    sample_2_df, sample_2_diag, sample_2_meta = build_sample_2(eligible)
    sample_3_df, sample_3_diag, sample_3_meta = build_sample_3(eligible)

    diagnostics = pd.concat([sample_1_diag, sample_2_diag, sample_3_diag], ignore_index=True)

    meta_rows: list[dict[str, object]] = [
        {"chiave": "input_file", "valore": str(input_file)},
        {"chiave": "sheet", "valore": first_sheet},
        {"chiave": "record_eleggibili", "valore": int(len(eligible))},
        {
            "chiave": "componenti_eleggibili",
            "valore": format_number(float(eligible["_componenti_num"].sum())),
        },
    ]

    for sample_meta in [sample_1_meta, sample_2_meta, sample_3_meta]:
        sample_name = str(sample_meta.get("campione", ""))
        for k, v in sample_meta.items():
            if k == "campione":
                continue
            meta_rows.append({"chiave": f"{sample_name}.{k}", "valore": v})

    meta_df = pd.DataFrame(meta_rows)

    out_path = output_file
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            split_output_columns(sample_1_df, original_columns).to_excel(writer, sheet_name="campione_1", index=False)
            split_output_columns(sample_2_df, original_columns).to_excel(writer, sheet_name="campione_2", index=False)
            split_output_columns(sample_3_df, original_columns).to_excel(writer, sheet_name="campione_3", index=False)
            diagnostics.to_excel(writer, sheet_name="diagnostica_celle", index=False)
            meta_df.to_excel(writer, sheet_name="meta", index=False)
    except PermissionError:
        out_path = output_file.with_name(f"{output_file.stem}_nuovo{output_file.suffix}")
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            split_output_columns(sample_1_df, original_columns).to_excel(writer, sheet_name="campione_1", index=False)
            split_output_columns(sample_2_df, original_columns).to_excel(writer, sheet_name="campione_2", index=False)
            split_output_columns(sample_3_df, original_columns).to_excel(writer, sheet_name="campione_3", index=False)
            diagnostics.to_excel(writer, sheet_name="diagnostica_celle", index=False)
            meta_df.to_excel(writer, sheet_name="meta", index=False)
        print(f"File output in uso: salvato su percorso alternativo {out_path}")

    print(f"Input letto: {input_file}")
    print(f"Record eleggibili: {len(eligible)}")
    print(f"Componenti eleggibili: {format_number(float(eligible['_componenti_num'].sum()))}")
    print(f"Campione 1 - questionari: {len(sample_1_df)} | componenti: {format_number(float(sample_1_df['_componenti_num'].sum()))}")
    print(f"Campione 2 - questionari: {len(sample_2_df)} | componenti: {format_number(float(sample_2_df['_componenti_num'].sum()))}")
    print(f"Campione 3 - questionari: {len(sample_3_df)} | componenti: {format_number(float(sample_3_df['_componenti_num'].sum()))}")
    print(f"Output generato: {out_path}")

    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Estrae 3 sottocampioni dai pernottanti senza pacchetto, con quote su numero_componenti e "
            "bilanciamento per fascia eta."
        )
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="File Excel sorgente")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="File Excel output")
    args = parser.parse_args()

    run(args.input, args.output)


if __name__ == "__main__":
    main()
