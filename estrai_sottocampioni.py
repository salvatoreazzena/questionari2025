from __future__ import annotations

import argparse
import random
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

DEFAULT_INPUT = Path("questionari_fonte_veri_outlier_sostituiti_imputati.xlsx")
DEFAULT_OUTPUT = Path("questionari_sottocampioni.xlsx")
DEFAULT_SELECTION_TRIALS = 200

ID_COL = "ID"
DURATION_COL = "durata_soggiorno"
PACCHETTO_COL = "pacchetto"
COMPONENTI_COL = "numero_componenti"
ETA_COL = "fascia_età"
STATO_COL = "stato_provenienza"
REGIONE_COL = "regione_provenienza"
MOTIV_PRINC_COL = "motivazione_principale"
MOTIV_SEC_COL = "motivazione_secondaria"

NULL_VALUE = "NULL"
NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}
BALNEARE_VALUE = "BALNEARE"

REQUIRED_COLUMNS = [
    DURATION_COL,
    PACCHETTO_COL,
    COMPONENTI_COL,
    ETA_COL,
    STATO_COL,
    REGIONE_COL,
    MOTIV_PRINC_COL,
    MOTIV_SEC_COL,
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
    out[MOTIV_PRINC_COL] = normalize_case_insensitive_series(normalize_text_series(out[MOTIV_PRINC_COL]))
    out[MOTIV_SEC_COL] = normalize_case_insensitive_series(normalize_text_series(out[MOTIV_SEC_COL]))

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


def build_duration_class(duration_num_series: pd.Series) -> pd.Series:
    s = pd.to_numeric(duration_num_series, errors="coerce")
    out = pd.Series(pd.NA, index=s.index, dtype="string")
    out = out.mask((s >= 1) & (s <= 3), "1-3")
    out = out.mask((s >= 4) & (s <= 7), "4-7")
    out = out.mask((s >= 8) & (s <= 14), "8-14")
    out = out.mask((s >= 15) & (s <= 30), "15-30")
    return out


def build_sampling_age_band(series: pd.Series, *, merge_56_65_over65: bool) -> pd.Series:
    out = series.astype("string").fillna("").str.strip().str.upper()
    if not merge_56_65_over65:
        return out

    merge_values = {"56-65", "OVER 65"}
    return out.mask(out.isin(merge_values), "56-OVER65")


def compute_age_bands_common(
    df: pd.DataFrame,
    *,
    group_col: str,
    group_values: list[str],
    age_col: str = ETA_COL,
) -> list[str]:
    age_sets: list[set[str]] = []
    for group in group_values:
        group_df = df[df[group_col] == group]
        grouped = group_df.groupby(age_col, as_index=False).agg(componenti=("_componenti_num", "sum"))
        valid_ages = {str(r[age_col]) for _, r in grouped.iterrows() if float(r["componenti"]) > 0}
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
    age_col: str = ETA_COL,
) -> list[CellPlan]:
    plans: list[CellPlan] = []
    for group in group_values:
        for age_band in age_bands:
            cell = df[(df[group_col] == group) & (df[age_col] == age_band)]
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


def _component_bucket(value: float) -> str:
    if value <= 2:
        return "LOW"
    if value <= 4:
        return "MID"
    return "HIGH"


def _build_variety_order(cell_df: pd.DataFrame, rng: random.Random) -> pd.DataFrame:
    work = cell_df.copy()
    work["_bucket"] = work["_componenti_num"].map(lambda v: _component_bucket(float(v)))

    low = work[work["_bucket"] == "LOW"].sample(frac=1, random_state=rng.randint(0, 10**9))
    mid = work[work["_bucket"] == "MID"].sample(frac=1, random_state=rng.randint(0, 10**9))
    high = work[work["_bucket"] == "HIGH"].sample(frac=1, random_state=rng.randint(0, 10**9))

    chunks = {
        "LOW": low,
        "MID": mid,
        "HIGH": high,
    }
    pointers = {k: 0 for k in chunks}

    patterns = [
        ["HIGH", "MID", "LOW"],
        ["MID", "HIGH", "LOW"],
        ["HIGH", "LOW", "MID"],
        ["MID", "LOW", "HIGH"],
    ]
    pattern = patterns[rng.randrange(len(patterns))]

    ordered_idx: list[int] = []
    remaining = int(len(work))
    while remaining > 0:
        progressed = False
        for bucket in pattern:
            ptr = pointers[bucket]
            df_b = chunks[bucket]
            if ptr < len(df_b):
                ordered_idx.append(int(df_b.index[ptr]))
                pointers[bucket] += 1
                remaining -= 1
                progressed = True
        if not progressed:
            break

    out = work.loc[pd.Index(ordered_idx)]
    return out.drop(columns=["_bucket"])


def select_cell_rows_with_optimal_deviation(
    cell_df: pd.DataFrame,
    target_components: float,
    *,
    rng: random.Random,
    selection_trials: int,
) -> tuple[pd.Index, float]:
    if target_components <= 0:
        return pd.Index([], dtype="int64"), 0.0

    best_idx = pd.Index([], dtype="int64")
    best_total = 0.0
    best_key: tuple[float, float, int, float] | None = None

    trials = max(1, int(selection_trials))
    for _ in range(trials):
        ordered = _build_variety_order(cell_df, rng)

        cumulative = 0.0
        cumulative_before = 0.0
        selected_indices: list[int] = []

        for idx, row in ordered.iterrows():
            cumulative_before = cumulative
            cumulative += float(row["_componenti_num"])
            selected_indices.append(int(idx))
            if cumulative >= target_components:
                break

        if selected_indices and abs(cumulative_before - target_components) <= abs(cumulative - target_components):
            selected_indices = selected_indices[:-1]
            cumulative = cumulative_before

        selected = cell_df.loc[pd.Index(selected_indices)] if selected_indices else cell_df.iloc[0:0]
        selected_components = selected["_componenti_num"].astype(float)
        unique_comp = int(selected_components.nunique())
        std_comp = float(selected_components.std(ddof=0)) if len(selected_components) > 1 else 0.0

        diff = abs(cumulative - target_components)
        overshoot = max(cumulative - target_components, 0.0)
        key = (diff, overshoot, -unique_comp, -std_comp)

        if best_key is None or key < best_key:
            best_key = key
            best_idx = pd.Index(selected_indices, dtype="int64")
            best_total = cumulative

    return best_idx, best_total


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
    rng: random.Random,
    selection_trials: int,
    merge_56_65_over65: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    work = df.copy()
    work["_fascia_eta_sampling"] = build_sampling_age_band(
        work[ETA_COL],
        merge_56_65_over65=merge_56_65_over65,
    )
    age_col = "_fascia_eta_sampling"

    age_bands = compute_age_bands_common(
        work,
        group_col=group_col,
        group_values=group_values,
        age_col=age_col,
    )
    cell_plan = build_cell_plan(
        work,
        group_col=group_col,
        group_values=group_values,
        age_bands=age_bands,
        age_col=age_col,
    )

    target_components = min(cell.available_components for cell in cell_plan)
    if target_components <= 0:
        raise ValueError(f"Target non valido per {sample_name}: {target_components}")

    selected_indices_all: list[pd.Index] = []
    selection_rows: list[dict[str, object]] = []

    for cell in cell_plan:
        cell_df = work[(work[group_col] == cell.group_value) & (work[age_col] == cell.fascia_eta)]
        if cell.available_components < target_components - 1e-9:
            raise ValueError(
                f"Capacita insufficiente per {sample_name} nella cella {cell.group_value} x {cell.fascia_eta}."
            )

        selected_idx, selected_components = select_cell_rows_with_optimal_deviation(
            cell_df,
            target_components,
            rng=rng,
            selection_trials=selection_trials,
        )
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

    selected = work.loc[final_idx].copy()
    selected["campione"] = sample_name
    selected["dimensione_gruppo"] = group_col

    selected[group_col] = selected[group_col].astype("string")
    selected["strato_gruppo"] = selected[group_col].astype("string")
    selected["strato_fascia_eta"] = selected[age_col].astype("string")

    selected = selected.sort_values(by=["strato_gruppo", "strato_fascia_eta", "_id_text", "_row_order"], kind="mergesort")

    selection_df = pd.DataFrame(selection_rows)
    meta = {
        "campione": sample_name,
        "n_gruppi": len(group_values),
        "n_fasce_eta": len(age_bands),
        "merge_56_65_over65": bool(merge_56_65_over65),
        "target_componenti_per_cella": format_number(target_components),
        "componenti_totali_selezionati": format_number(float(selected["_componenti_num"].sum())),
        "questionari_totali_selezionati": int(len(selected)),
    }
    return selected, selection_df, meta


def build_sample_1(
    df: pd.DataFrame,
    *,
    rng: random.Random,
    selection_trials: int,
    merge_56_65_over65: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[df["macro_provenienza"].isin(["ITALIANI", "STRANIERI"])].copy()
    if subset.empty:
        raise ValueError("Campione 1 non costruibile: nessuna riga ITALIANI/STRANIERI.")

    group_values = ["ITALIANI", "STRANIERI"]
    selected, cell_summary, meta = extract_sample(
        subset,
        sample_name="Campione_1_50_50",
        group_col="macro_provenienza",
        group_values=group_values,
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
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


def build_sample_1a(
    df: pd.DataFrame,
    *,
    rng: random.Random,
    selection_trials: int,
    merge_56_65_over65: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[
        df["macro_provenienza"].isin(["ITALIANI", "STRANIERI"])
        & (df[MOTIV_PRINC_COL] != NULL_VALUE)
        & (df[MOTIV_PRINC_COL] != BALNEARE_VALUE)
    ].copy()
    if subset.empty:
        raise ValueError("Campione 1a non costruibile: nessun turista non balneare disponibile.")

    selected, cell_summary, meta = extract_sample(
        subset,
        sample_name="Campione_1a_non_balneari",
        group_col="macro_provenienza",
        group_values=["ITALIANI", "STRANIERI"],
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    return selected, cell_summary, meta


def build_sample_1b(
    df: pd.DataFrame,
    *,
    rng: random.Random,
    selection_trials: int,
    merge_56_65_over65: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[
        df["macro_provenienza"].isin(["ITALIANI", "STRANIERI"])
        & (df[MOTIV_SEC_COL] != NULL_VALUE)
    ].copy()
    if subset.empty:
        raise ValueError("Campione 1b non costruibile: nessun turista con motivazione secondaria disponibile.")

    selected, cell_summary, meta = extract_sample(
        subset,
        sample_name="Campione_1b_motivazione_secondaria",
        group_col="macro_provenienza",
        group_values=["ITALIANI", "STRANIERI"],
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    return selected, cell_summary, meta


def build_sample_1c(
    df: pd.DataFrame,
    *,
    rng: random.Random,
    selection_trials: int,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[df["macro_provenienza"].isin(["ITALIANI", "STRANIERI"]) & (df["_durata_num"] <= 30)].copy()
    subset["classe_durata_30"] = build_duration_class(subset["_durata_num"])
    subset = subset[subset["classe_durata_30"].notna()].copy()
    if subset.empty:
        raise ValueError("Campione 1c non costruibile: nessun turista nelle classi durata 1-30 giorni.")

    duration_classes = ["1-3", "4-7", "8-14", "15-30"]
    cell_plan: list[CellPlan] = []
    for group in ["ITALIANI", "STRANIERI"]:
        for duration_class in duration_classes:
            cell = subset[
                (subset["macro_provenienza"] == group)
                & (subset["classe_durata_30"] == duration_class)
            ]
            cell_plan.append(
                CellPlan(
                    group_value=group,
                    fascia_eta=duration_class,
                    available_components=float(cell["_componenti_num"].sum()),
                    available_questionari=int(len(cell)),
                )
            )

    target_components = min(cell.available_components for cell in cell_plan)
    if target_components <= 0:
        raise ValueError(f"Target non valido per Campione_1c_durata_1_30: {target_components}")

    selected_indices_all: list[pd.Index] = []
    selection_rows: list[dict[str, object]] = []

    for cell in cell_plan:
        cell_df = subset[
            (subset["macro_provenienza"] == cell.group_value)
            & (subset["classe_durata_30"] == cell.fascia_eta)
        ]
        selected_idx, selected_components = select_cell_rows_with_optimal_deviation(
            cell_df,
            target_components,
            rng=rng,
            selection_trials=selection_trials,
        )
        selected_indices_all.append(selected_idx)

        selection_rows.append(
            {
                "campione": "Campione_1c_durata_1_30",
                "dimensione_gruppo": "macro_provenienza",
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

    final_idx = pd.Index([], dtype="int64")
    for idx in selected_indices_all:
        final_idx = final_idx.append(idx)

    selected = subset.loc[final_idx].copy()
    selected["campione"] = "Campione_1c_durata_1_30"
    selected["dimensione_gruppo"] = "macro_provenienza"
    selected["strato_gruppo"] = selected["macro_provenienza"].astype("string")
    selected["strato_fascia_eta"] = selected["classe_durata_30"].astype("string")
    selected = selected.sort_values(by=["strato_gruppo", "strato_fascia_eta", "_id_text", "_row_order"], kind="mergesort")

    cell_summary = pd.DataFrame(selection_rows)
    meta = {
        "campione": "Campione_1c_durata_1_30",
        "n_gruppi": 2,
        "n_classi_durata": len(duration_classes),
        "classi_durata": ", ".join(duration_classes),
        "target_componenti_per_cella": format_number(target_components),
        "componenti_totali_selezionati": format_number(float(selected["_componenti_num"].sum())),
        "questionari_totali_selezionati": int(len(selected)),
    }
    return selected, cell_summary, meta


def build_sample_1d(
    df: pd.DataFrame,
    *,
    rng: random.Random,
    selection_trials: int,
    merge_56_65_over65: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    subset = df[
        df["macro_provenienza"].isin(["ITALIANI", "STRANIERI"])
        & (df[MOTIV_PRINC_COL] != NULL_VALUE)
    ].copy()
    if subset.empty:
        raise ValueError("Campione 1d non costruibile: nessuna motivazione principale valorizzata.")

    top5 = compute_group_ranking(subset, MOTIV_PRINC_COL, top_n=5)
    subset = subset[subset[MOTIV_PRINC_COL].isin(top5)].copy()

    selected, cell_summary, meta = extract_sample(
        subset,
        sample_name="Campione_1d_top5_motivazioni",
        group_col="macro_provenienza",
        group_values=["ITALIANI", "STRANIERI"],
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    meta["top5_motivazioni_principali"] = ", ".join(top5)
    return selected, cell_summary, meta


def build_sample_2(
    df: pd.DataFrame,
    *,
    rng: random.Random,
    selection_trials: int,
    merge_56_65_over65: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
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
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    meta["top5"] = ", ".join(top5)
    return selected, cell_summary, meta


def build_sample_3(
    df: pd.DataFrame,
    *,
    rng: random.Random,
    selection_trials: int,
    merge_56_65_over65: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
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
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    meta["top5"] = ", ".join(top5)
    return selected, cell_summary, meta


def split_output_columns(df_selected: pd.DataFrame, original_columns: list[str]) -> pd.DataFrame:
    existing_original = [c for c in original_columns if c in df_selected.columns]
    return df_selected[existing_original].copy()


def run(
    input_file: Path,
    output_file: Path,
    *,
    seed: int | None,
    selection_trials: int,
    merge_56_65_over65: bool,
) -> Path:
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

    rng = random.Random(seed)

    sample_1_df, sample_1_diag, sample_1_meta = build_sample_1(
        eligible,
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    sample_1a_df, sample_1a_diag, sample_1a_meta = build_sample_1a(
        sample_1_df,
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    sample_1b_df, sample_1b_diag, sample_1b_meta = build_sample_1b(
        sample_1_df,
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    sample_1c_df, sample_1c_diag, sample_1c_meta = build_sample_1c(sample_1_df, rng=rng, selection_trials=selection_trials)
    sample_1d_df, sample_1d_diag, sample_1d_meta = build_sample_1d(
        sample_1_df,
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    sample_2_df, sample_2_diag, sample_2_meta = build_sample_2(
        eligible,
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )
    sample_3_df, sample_3_diag, sample_3_meta = build_sample_3(
        eligible,
        rng=rng,
        selection_trials=selection_trials,
        merge_56_65_over65=merge_56_65_over65,
    )

    diagnostics = pd.concat(
        [
            sample_1_diag,
            sample_1a_diag,
            sample_1b_diag,
            sample_1c_diag,
            sample_1d_diag,
            sample_2_diag,
            sample_3_diag,
        ],
        ignore_index=True,
    )

    meta_rows: list[dict[str, object]] = [
        {"chiave": "input_file", "valore": str(input_file)},
        {"chiave": "sheet", "valore": first_sheet},
        {"chiave": "record_eleggibili", "valore": int(len(eligible))},
        {
            "chiave": "componenti_eleggibili",
            "valore": format_number(float(eligible["_componenti_num"].sum())),
        },
        {"chiave": "seed", "valore": "AUTO" if seed is None else seed},
        {"chiave": "selection_trials", "valore": int(selection_trials)},
        {"chiave": "merge_56_65_over65", "valore": bool(merge_56_65_over65)},
    ]

    for sample_meta in [
        sample_1_meta,
        sample_1a_meta,
        sample_1b_meta,
        sample_1c_meta,
        sample_1d_meta,
        sample_2_meta,
        sample_3_meta,
    ]:
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
            split_output_columns(sample_1a_df, original_columns).to_excel(writer, sheet_name="campione_1a", index=False)
            split_output_columns(sample_1b_df, original_columns).to_excel(writer, sheet_name="campione_1b", index=False)
            split_output_columns(sample_1c_df, original_columns).to_excel(writer, sheet_name="campione_1c", index=False)
            split_output_columns(sample_1d_df, original_columns).to_excel(writer, sheet_name="campione_1d", index=False)
            split_output_columns(sample_2_df, original_columns).to_excel(writer, sheet_name="campione_2", index=False)
            split_output_columns(sample_3_df, original_columns).to_excel(writer, sheet_name="campione_3", index=False)
            diagnostics.to_excel(writer, sheet_name="diagnostica_celle", index=False)
            meta_df.to_excel(writer, sheet_name="meta", index=False)
    except PermissionError:
        out_path = output_file.with_name(f"{output_file.stem}_nuovo{output_file.suffix}")
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            split_output_columns(sample_1_df, original_columns).to_excel(writer, sheet_name="campione_1", index=False)
            split_output_columns(sample_1a_df, original_columns).to_excel(writer, sheet_name="campione_1a", index=False)
            split_output_columns(sample_1b_df, original_columns).to_excel(writer, sheet_name="campione_1b", index=False)
            split_output_columns(sample_1c_df, original_columns).to_excel(writer, sheet_name="campione_1c", index=False)
            split_output_columns(sample_1d_df, original_columns).to_excel(writer, sheet_name="campione_1d", index=False)
            split_output_columns(sample_2_df, original_columns).to_excel(writer, sheet_name="campione_2", index=False)
            split_output_columns(sample_3_df, original_columns).to_excel(writer, sheet_name="campione_3", index=False)
            diagnostics.to_excel(writer, sheet_name="diagnostica_celle", index=False)
            meta_df.to_excel(writer, sheet_name="meta", index=False)
        print(f"File output in uso: salvato su percorso alternativo {out_path}")

    print(f"Input letto: {input_file}")
    print(f"Record eleggibili: {len(eligible)}")
    print(f"Componenti eleggibili: {format_number(float(eligible['_componenti_num'].sum()))}")
    print(f"Campione 1 - questionari: {len(sample_1_df)} | componenti: {format_number(float(sample_1_df['_componenti_num'].sum()))}")
    print(f"Campione 1a - questionari: {len(sample_1a_df)} | componenti: {format_number(float(sample_1a_df['_componenti_num'].sum()))}")
    print(f"Campione 1b - questionari: {len(sample_1b_df)} | componenti: {format_number(float(sample_1b_df['_componenti_num'].sum()))}")
    print(f"Campione 1c - questionari: {len(sample_1c_df)} | componenti: {format_number(float(sample_1c_df['_componenti_num'].sum()))}")
    print(f"Campione 1d - questionari: {len(sample_1d_df)} | componenti: {format_number(float(sample_1d_df['_componenti_num'].sum()))}")
    print(f"Campione 2 - questionari: {len(sample_2_df)} | componenti: {format_number(float(sample_2_df['_componenti_num'].sum()))}")
    print(f"Campione 3 - questionari: {len(sample_3_df)} | componenti: {format_number(float(sample_3_df['_componenti_num'].sum()))}")
    print(f"Output generato: {out_path}")

    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Estrae sottocampioni dai pernottanti senza pacchetto, con quote su numero_componenti e "
            "bilanciamento per fascia eta / classi durata in base al campione."
        )
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="File Excel sorgente")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="File Excel output")
    parser.add_argument(
        "--seed",
        type=int,
        default=None,
        help="Seed random per estrazione riproducibile (default: casuale)",
    )
    parser.add_argument(
        "--selection-trials",
        type=int,
        default=DEFAULT_SELECTION_TRIALS,
        help="Numero tentativi casuali per cella (piu alto = piu varieta, piu lento)",
    )
    parser.add_argument(
        "--merge-56-65-over65",
        action="store_true",
        help="Aggrega la fascia 56-65 con OVER 65 durante la stratificazione",
    )
    args = parser.parse_args()

    run(
        args.input,
        args.output,
        seed=args.seed,
        selection_trials=args.selection_trials,
        merge_56_65_over65=args.merge_56_65_over65,
    )


if __name__ == "__main__":
    main()
