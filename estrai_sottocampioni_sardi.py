from __future__ import annotations

import argparse
import random
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

DEFAULT_INPUT = Path("questionari_sardi_outlier_sostituiti.xlsx")
DEFAULT_OUTPUT = Path("questionari_sardi_sottocampioni.xlsx")

ID_COL = "ID"
VACANZA_COL = "vacanza_2025"
PACCHETTO_COL = "pacchetto"
COMPONENTI_COL = "numero_componenti"
ETA_COL = "fascia_età"
DURATA_COL = "durata_soggiorno"
MOTIV_COL = "motivazione_principale"
DEST_COL = "si_dove"

NULL_VALUE = "NULL"
NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}
SI_VALUES = {"SI", "SÌ"}
NO_VALUES = {"NO"}
BALNEARE_VALUE = "BALNEARE"

REQUIRED_COLUMNS = [VACANZA_COL, PACCHETTO_COL, COMPONENTI_COL, ETA_COL, DURATA_COL, MOTIV_COL]
PROVINCIA_COL = "provincia_provenienza"


@dataclass(frozen=True)
class CellPlan:
    group_value: str
    fascia_eta: str
    available_components: float
    available_rows: int


def normalize_text_series(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("").str.strip()
    upper = s.str.upper()
    mask_null = upper.isin({token.upper() for token in NULL_TOKENS})
    return s.mask(mask_null, NULL_VALUE)


def normalize_upper(series: pd.Series) -> pd.Series:
    return normalize_text_series(series).astype("string").str.upper()


def is_empty_or_na(value: object) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def validate_columns(df: pd.DataFrame) -> None:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti: {missing}")


def build_age_sampling_band(series: pd.Series) -> pd.Series:
    s = normalize_upper(series)
    return s.mask(s.isin({"56-65", "OVER 65"}), "56-OVER65")


def build_duration_class(series: pd.Series) -> pd.Series:
    d = pd.to_numeric(series, errors="coerce")
    out = pd.Series(pd.NA, index=d.index, dtype="string")
    out = out.mask((d >= 1) & (d <= 3), "1-3")
    out = out.mask((d >= 4) & (d <= 7), "4-7")
    out = out.mask((d >= 8) & (d <= 14), "8-14")
    out = out.mask((d >= 15) & (d <= 30), "15-30")
    return out


def build_destination_group(series: pd.Series) -> pd.Series:
    s = normalize_upper(series)
    out = pd.Series(pd.NA, index=s.index, dtype="string")
    out = out.mask(s.eq("IN SARDEGNA"), "SARDEGNA")
    out = out.mask(s.eq("IN ITALIA"), "ITALIA")
    out = out.mask(s.eq("ALL'ESTERO"), "ESTERO")
    return out


def normalize_province_series(series: pd.Series) -> pd.Series:
    return normalize_upper(series)


def get_valid_values_by_components(df: pd.DataFrame, value_col: str) -> list[str]:
    grouped = (
        df.groupby(value_col, as_index=False)
        .agg(componenti=("_componenti_num", "sum"))
        .sort_values(by=[value_col], kind="mergesort")
    )
    return [str(r[value_col]) for _, r in grouped.iterrows() if float(r["componenti"]) > 0]


def motivation_sort_priority_sardinia(value: object) -> int:
    text = str(value).strip().upper()
    if text == "ENOGASTRONOMICO":
        return 0
    if text == "ALTRO":
        return 1
    return 2


def build_stratum_keys(df: pd.DataFrame, cols: list[str]) -> list[tuple[str, ...]]:
    if not cols:
        return [("__ALL__",)] * len(df)
    normalized_cols = [normalize_text_series(df[col]).astype("string").tolist() for col in cols]
    return list(zip(*normalized_cols, strict=False))


def format_number(value: float) -> int | float:
    rounded = round(float(value), 6)
    as_int = int(round(rounded))
    if abs(rounded - as_int) < 1e-9:
        return as_int
    return rounded


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

    chunks = {"LOW": low, "MID": mid, "HIGH": high}
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
    selection_trials: int = 200,
) -> tuple[pd.Index, float]:
    if target_components <= 0:
        return pd.Index([], dtype="int64"), 0.0

    can_meet_target = float(cell_df["_componenti_num"].sum()) >= float(target_components) - 1e-9
    component_values = [float(v) for v in cell_df["_componenti_num"].tolist()]
    if component_values and is_integer_like(target_components) and all(is_integer_like(v) for v in component_values):
        target_int = int(round(target_components))
        values_int = [int(round(v)) for v in component_values]
        max_value = max(values_int)
        upper_bound = target_int + max_value

        parent: dict[int, tuple[int, int] | None] = {0: None}
        for pos, val in enumerate(values_int):
            current_sums = sorted(parent.keys(), reverse=True)
            for s in current_sums:
                new_sum = s + val
                if new_sum > upper_bound:
                    continue
                if new_sum not in parent:
                    parent[new_sum] = (s, pos)

        if len(parent) > 1:
            best_sum = min(
                parent.keys(),
                key=lambda s: (
                    1 if can_meet_target and s < target_int else 0,
                    max(s - target_int, 0),
                    abs(s - target_int),
                    -s,
                ),
            )
            if best_sum > 0:
                chosen_pos: list[int] = []
                cursor = best_sum
                while cursor != 0:
                    prev = parent.get(cursor)
                    if prev is None:
                        break
                    prev_sum, pos = prev
                    chosen_pos.append(pos)
                    cursor = prev_sum
                chosen_pos.reverse()

                idx_values = cell_df.index.tolist()
                chosen_idx = [int(idx_values[p]) for p in chosen_pos]
                return pd.Index(chosen_idx, dtype="int64"), float(best_sum)

    best_idx = pd.Index([], dtype="int64")
    best_total = 0.0
    best_key: tuple[float, float, float, float, int, float] | None = None

    for _ in range(max(1, int(selection_trials))):
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

        if (
            selected_indices
            and (not can_meet_target or cumulative_before >= target_components - 1e-9)
            and abs(cumulative_before - target_components) <= abs(cumulative - target_components)
        ):
            selected_indices = selected_indices[:-1]
            cumulative = cumulative_before

        selected = cell_df.loc[pd.Index(selected_indices)] if selected_indices else cell_df.iloc[0:0]
        selected_components = selected["_componenti_num"].astype(float)
        unique_comp = int(selected_components.nunique())
        std_comp = float(selected_components.std(ddof=0)) if len(selected_components) > 1 else 0.0

        overshoot = max(cumulative - target_components, 0.0)
        undershoot = max(target_components - cumulative, 0.0)
        key = (
            1 if can_meet_target and undershoot > 1e-9 else 0,
            overshoot,
            undershoot,
            abs(cumulative - target_components),
            -unique_comp,
            -std_comp,
        )

        if best_key is None or key < best_key:
            best_key = key
            best_idx = pd.Index(selected_indices, dtype="int64")
            best_total = cumulative

    return best_idx, best_total


def prepare_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out[VACANZA_COL] = normalize_upper(out[VACANZA_COL])
    out[ETA_COL] = build_age_sampling_band(out[ETA_COL])
    out[MOTIV_COL] = normalize_upper(out[MOTIV_COL])
    out["_durata_num"] = pd.to_numeric(out[DURATA_COL], errors="coerce")
    out["_componenti_num"] = pd.to_numeric(out[COMPONENTI_COL], errors="coerce")

    out["_row_order"] = out.index.astype("int64")
    if ID_COL in out.columns:
        out["_id_text"] = out[ID_COL].astype("string").fillna("")
    else:
        out["_id_text"] = (out.index + 2).astype("string")

    mask = (
        out[PACCHETTO_COL].map(is_empty_or_na)
        & (out[ETA_COL] != NULL_VALUE)
        & out[VACANZA_COL].isin(SI_VALUES | NO_VALUES)
        & out["_componenti_num"].notna()
        & (out["_componenti_num"] > 0)
    )
    return out.loc[mask].copy()


def compute_common_age_bands(df: pd.DataFrame, group_col: str, group_values: list[str]) -> list[str]:
    sets: list[set[str]] = []
    for gv in group_values:
        grouped = (
            df[df[group_col] == gv]
            .groupby(ETA_COL, as_index=False)
            .agg(componenti=("_componenti_num", "sum"))
        )
        age_vals = {str(r[ETA_COL]) for _, r in grouped.iterrows() if float(r["componenti"]) > 0}
        if not age_vals:
            raise ValueError(f"Nessuna fascia eta per gruppo {gv}")
        sets.append(age_vals)

    common = set.intersection(*sets)
    if not common:
        raise ValueError("Nessuna fascia eta comune tra i gruppi richiesti")
    return sorted(common)


def select_equal_rows_by_age(
    df: pd.DataFrame,
    *,
    sample_name: str,
    group_col: str,
    group_values: list[str],
    rng: random.Random,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    age_bands = compute_common_age_bands(df, group_col, group_values)

    plans: list[CellPlan] = []
    for gv in group_values:
        for age in age_bands:
            cell = df[(df[group_col] == gv) & (df[ETA_COL] == age)]
            plans.append(
                CellPlan(
                    gv,
                    age,
                    float(cell["_componenti_num"].sum()),
                    int(len(cell)),
                )
            )

    target_components = min(p.available_components for p in plans)
    if target_components <= 0:
        raise ValueError(f"Campione {sample_name} non costruibile: target componenti = {target_components}")

    selected_indices_all: list[pd.Index] = []
    diag_rows: list[dict[str, object]] = []

    for p in plans:
        cell = df[(df[group_col] == p.group_value) & (df[ETA_COL] == p.fascia_eta)].copy()
        selected_idx, selected_components = select_cell_rows_with_optimal_deviation(
            cell,
            target_components,
            rng=rng,
        )
        selected_indices_all.append(selected_idx)
        diag_rows.append(
            {
                "campione": sample_name,
                "dimensione_gruppo": group_col,
                "valore_gruppo": p.group_value,
                "fascia_eta": p.fascia_eta,
                "componenti_disponibili": format_number(p.available_components),
                "target_componenti": format_number(target_components),
                "componenti_selezionati": format_number(selected_components),
                "overshoot_componenti": format_number(selected_components - target_components),
            }
        )

    final_idx = pd.Index([], dtype="int64")
    for idx in selected_indices_all:
        final_idx = final_idx.append(idx)

    selected = df.loc[final_idx].copy()
    selected["campione"] = sample_name
    selected = selected.sort_values(by=[group_col, ETA_COL, "_id_text", "_row_order"], kind="mergesort")

    meta = {
        "campione": sample_name,
        "dimensione_gruppo": group_col,
        "n_gruppi": len(group_values),
        "n_fasce_eta": len(age_bands),
        "target_componenti_per_cella": format_number(target_components),
        "componenti_totali_selezionati": format_number(float(selected["_componenti_num"].sum())),
        "fasce_eta": ", ".join(age_bands),
    }
    return selected, pd.DataFrame(diag_rows), meta


def select_equal_rows_by_group_components(
    df: pd.DataFrame,
    *,
    sample_name: str,
    group_col: str,
    group_values: list[str],
    rng: random.Random,
    ensure_province_coverage: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, object]]:
    plans: list[tuple[str, float]] = []
    for gv in group_values:
        cell = df[df[group_col] == gv].copy()
        available_components = float(cell["_componenti_num"].sum())
        if available_components <= 0:
            raise ValueError(f"Campione {sample_name} non costruibile: nessun componente per gruppo {gv}")
        plans.append((gv, available_components))

    target_components = min(available_components for _, available_components in plans)
    if target_components <= 0:
        raise ValueError(f"Campione {sample_name} non costruibile: target componenti = {target_components}")

    reserved_idx: set[int] = set()
    reserved_components_by_group: dict[str, float] = {str(gv): 0.0 for gv in group_values}
    reserved_provinces: list[str] = []

    if ensure_province_coverage and PROVINCIA_COL in df.columns:
        work = df.copy()
        work["_prov_norm"] = normalize_province_series(work[PROVINCIA_COL])
        valid_mask = work["_prov_norm"].notna() & (work["_prov_norm"] != NULL_VALUE)
        required_provinces = sorted(set(work.loc[valid_mask, "_prov_norm"].astype("string").tolist()))
        prov_freq = work.loc[valid_mask, "_prov_norm"].value_counts().to_dict()

        for prov in sorted(required_provinces, key=lambda p: (prov_freq.get(p, 0), str(p))):
            candidates = work[(work["_prov_norm"] == prov) & (~work.index.isin(reserved_idx))].copy()
            if candidates.empty:
                continue

            best_idx: int | None = None
            best_key: tuple[float, float, float, str] | None = None
            for idx, row in candidates.iterrows():
                group_value = str(row[group_col])
                if group_value not in reserved_components_by_group:
                    continue
                comp = float(row["_componenti_num"])
                reserved_after = reserved_components_by_group[group_value] + comp
                key = (
                    max(reserved_after - target_components, 0.0),
                    reserved_components_by_group[group_value],
                    comp,
                    str(row["_id_text"]),
                )
                if best_key is None or key < best_key:
                    best_key = key
                    best_idx = int(idx)

            if best_idx is None:
                continue

            reserved_idx.add(best_idx)
            reserved_group = str(work.loc[best_idx, group_col])
            reserved_components_by_group[reserved_group] += float(work.loc[best_idx, "_componenti_num"])
            reserved_provinces.append(str(prov))

    selected_indices_all: list[pd.Index] = []
    diag_rows: list[dict[str, object]] = []

    for gv, available_components in plans:
        cell = df[df[group_col] == gv].copy()
        reserved_cell = cell[cell.index.isin(reserved_idx)].copy()
        reserved_components = float(reserved_cell["_componenti_num"].sum())
        remaining_target = max(target_components - reserved_components, 0.0)
        pool = cell[~cell.index.isin(reserved_idx)].copy()

        extra_idx, extra_components = select_cell_rows_with_optimal_deviation(
            pool,
            remaining_target,
            rng=rng,
        )
        selected_idx = reserved_cell.index.append(extra_idx)
        selected_components = reserved_components + extra_components
        selected_indices_all.append(selected_idx)
        diag_rows.append(
            {
                "campione": sample_name,
                "dimensione_gruppo": group_col,
                "valore_gruppo": gv,
                "componenti_disponibili": format_number(available_components),
                "target_componenti": format_number(target_components),
                "componenti_riservati_province": format_number(reserved_components),
                "componenti_selezionati": format_number(selected_components),
                "overshoot_componenti": format_number(selected_components - target_components),
            }
        )

    final_idx = pd.Index([], dtype="int64")
    for idx in selected_indices_all:
        final_idx = final_idx.append(idx)

    selected = df.loc[final_idx].copy()
    selected["campione"] = sample_name
    selected = selected.sort_values(by=[group_col, "_id_text", "_row_order"], kind="mergesort")

    meta = {
        "campione": sample_name,
        "dimensione_gruppo": group_col,
        "n_gruppi": len(group_values),
        "target_componenti_per_cella": format_number(target_components),
        "componenti_totali_selezionati": format_number(float(selected["_componenti_num"].sum())),
        "province_riservate_in_estrazione": ", ".join(sorted(set(reserved_provinces))),
    }
    return selected, pd.DataFrame(diag_rows), meta


def ensure_all_provinces_covered(
    selected: pd.DataFrame,
    pool: pd.DataFrame,
    *,
    strata_cols: list[str],
    relaxed_strata_cols: list[str] | None = None,
    append_only: bool = False,
) -> tuple[pd.DataFrame, list[str], list[str], float]:
    if PROVINCIA_COL not in selected.columns or PROVINCIA_COL not in pool.columns:
        return selected, [], [], 0.0

    sel = selected.copy()
    pool2 = pool.copy()
    sel["_prov_norm"] = normalize_province_series(sel[PROVINCIA_COL])
    pool2["_prov_norm"] = normalize_province_series(pool2[PROVINCIA_COL])
    sel["_stratum_key"] = build_stratum_keys(sel, strata_cols)
    pool2["_stratum_key"] = build_stratum_keys(pool2, strata_cols)

    relaxed_cols = strata_cols if relaxed_strata_cols is None else relaxed_strata_cols
    sel["_relaxed_stratum_key"] = build_stratum_keys(sel, relaxed_cols)
    pool2["_relaxed_stratum_key"] = build_stratum_keys(pool2, relaxed_cols)

    valid_mask_pool = pool2["_prov_norm"].notna() & (pool2["_prov_norm"] != NULL_VALUE)
    required_provinces = sorted(set(pool2.loc[valid_mask_pool, "_prov_norm"].astype("string").tolist()))
    if not required_provinces:
        return selected, [], [], 0.0

    added: list[str] = []
    extra_components_added = 0.0
    valid_selected_strata = set(sel["_stratum_key"].tolist())
    valid_relaxed_strata = set(sel["_relaxed_stratum_key"].tolist())

    while True:
        current = set(
            sel.loc[sel["_prov_norm"].notna() & (sel["_prov_norm"] != NULL_VALUE), "_prov_norm"].astype("string").tolist()
        )
        missing = [p for p in required_provinces if p not in current]
        if not missing:
            break

        progress = False
        for target_prov in missing:
            candidates = pool2[(pool2["_prov_norm"] == target_prov) & (~pool2.index.isin(sel.index))].copy()
            if candidates.empty:
                continue
            candidates = candidates.sort_values(by=["_id_text", "_row_order"], kind="mergesort")

            replaced = False
            if not append_only:
                for cand_idx, cand_row in candidates.iterrows():
                    stratum_key = cand_row["_stratum_key"]
                    same_stratum = sel[sel["_stratum_key"] == stratum_key].copy()
                    if same_stratum.empty:
                        continue

                    prov_counts = same_stratum["_prov_norm"].value_counts(dropna=False)
                    removable = same_stratum[
                        same_stratum["_prov_norm"].map(lambda p: prov_counts.get(p, 0) > 1 or pd.isna(p) or p == NULL_VALUE)
                    ].copy()
                    if removable.empty:
                        continue

                    removable = removable.sort_values(
                        by=["_id_text", "_row_order"], ascending=[False, False], kind="mergesort"
                    )
                    drop_idx = int(removable.index[0])

                    sel = sel.drop(index=drop_idx)
                    row_add = pool2.loc[[cand_idx]].copy()
                    for col in sel.columns:
                        if col not in row_add.columns:
                            row_add[col] = pd.NA
                    row_add = row_add[sel.columns]
                    sel = pd.concat([sel, row_add], axis=0)
                    added.append(str(target_prov))
                    replaced = True
                    progress = True
                    break

            if replaced:
                break

            append_candidates = pool2[
                (pool2["_prov_norm"] == target_prov)
                & (~pool2.index.isin(sel.index))
                & (pool2["_relaxed_stratum_key"].isin(valid_relaxed_strata))
            ].copy()
            if append_candidates.empty:
                continue
            append_candidates = append_candidates.sort_values(
                by=["_componenti_num", "_id_text", "_row_order"],
                ascending=[True, True, True],
                kind="mergesort",
            )

            for cand_idx, cand_row in append_candidates.iterrows():
                row_add = pool2.loc[[cand_idx]].copy()
                for col in sel.columns:
                    if col not in row_add.columns:
                        row_add[col] = pd.NA
                row_add = row_add[sel.columns]
                sel = pd.concat([sel, row_add], axis=0)
                added.append(str(target_prov))
                extra_components_added += float(cand_row["_componenti_num"])
                progress = True
                break

            if progress:
                break

        if not progress:
            break

    current = set(
        sel.loc[sel["_prov_norm"].notna() & (sel["_prov_norm"] != NULL_VALUE), "_prov_norm"].astype("string").tolist()
    )
    missing = [p for p in required_provinces if p not in current]

    sel = sel.drop(columns=["_prov_norm", "_stratum_key", "_relaxed_stratum_key"], errors="ignore")
    return sel, sorted(set(added)), missing, extra_components_added


def build_sample_1(df_base: pd.DataFrame, rng: random.Random):
    subset = df_base[df_base[VACANZA_COL].isin(NO_VALUES)].copy()
    if subset.empty:
        raise ValueError("Campione 1 non costruibile: nessun turista con vacanza_2025 = No")

    subset["gruppo_vacanza"] = "NO_VACANZA"
    return select_equal_rows_by_age(
        subset,
        sample_name="campione_1_no_vacanza",
        group_col="gruppo_vacanza",
        group_values=["NO_VACANZA"],
        rng=rng,
    )


def build_sample_2(df_base: pd.DataFrame, rng: random.Random):
    subset = df_base[
        df_base[VACANZA_COL].isin(SI_VALUES)
        & df_base["_durata_num"].notna()
        & (df_base["_durata_num"] > 0)
    ].copy()
    if subset.empty:
        raise ValueError("Campione 2 non costruibile: nessun turista con vacanza_2025 = Si")

    # Campione 2: bilanciato per fascia eta; la provincia non entra nelle quote,
    # ma proviamo a mantenere almeno un caso per ogni provincia valida.
    subset["_prov_norm"] = normalize_province_series(subset[PROVINCIA_COL])
    subset["_is_prov_valid"] = subset["_prov_norm"].notna() & (subset["_prov_norm"] != NULL_VALUE)
    subset = subset[subset["_is_prov_valid"]].copy()
    if subset.empty:
        raise ValueError("Campione 2 non costruibile: nessuna provincia valida dopo esclusione ND/NULL")

    age_bands = get_valid_values_by_components(subset, ETA_COL)
    target_components = min(float(subset.loc[subset[ETA_COL] == a, "_componenti_num"].sum()) for a in age_bands)
    if target_components <= 0:
        raise ValueError("Campione 2 non costruibile: target componenti per fascia eta nullo")

    selected_indices_all: list[pd.Index] = []
    diag_rows: list[dict[str, object]] = []

    for age in age_bands:
        cell = subset[subset[ETA_COL] == age].copy()
        available_components = float(cell["_componenti_num"].sum())
        if available_components < target_components - 1e-9:
            raise ValueError(f"Campione 2 non costruibile per fascia eta {age}: componenti insufficienti")
        selected_idx, selected_components = select_cell_rows_with_optimal_deviation(
            cell,
            target_components,
            rng=rng,
        )
        selected_indices_all.append(selected_idx)

        diag_rows.append(
            {
                "campione": "campione_2_vacanza",
                "dimensione_gruppo": "gruppo_vacanza",
                "valore_gruppo": "VACANZA",
                "fascia_eta": age,
                "componenti_disponibili": format_number(available_components),
                "target_componenti": format_number(target_components),
                "componenti_selezionati": format_number(selected_components),
                "overshoot_componenti": format_number(selected_components - target_components),
            }
        )

    final_idx = pd.Index([], dtype="int64")
    for idx in selected_indices_all:
        final_idx = final_idx.append(idx)

    selected = subset.loc[final_idx].copy()
    selected, added_provinces, missing_provinces, extra_components_added = ensure_all_provinces_covered(
        selected,
        subset,
        strata_cols=[],
        relaxed_strata_cols=[],
        append_only=True,
    )
    if missing_provinces:
        raise ValueError(f"Campione 2 non costruibile: province non copribili {missing_provinces}")
    selected["campione"] = "campione_2_vacanza"
    selected = selected.sort_values(by=[ETA_COL, "_id_text", "_row_order"], kind="mergesort")
    final_age_bands = get_valid_values_by_components(selected, ETA_COL)

    meta = {
        "campione": "campione_2_vacanza",
        "dimensione_gruppo": "gruppo_vacanza",
        "n_gruppi": 1,
        "n_fasce_eta": len(age_bands),
        "target_componenti_per_cella": format_number(target_components),
        "componenti_totali_selezionati": format_number(float(selected["_componenti_num"].sum())),
        "fasce_eta": ", ".join(age_bands),
        "fasce_eta_nucleo_bilanciato": ", ".join(age_bands),
        "fasce_eta_presenti_finali": ", ".join(final_age_bands),
        "province_aggiunte_per_copertura": ", ".join(added_provinces),
        "province_non_copribili": ", ".join(missing_provinces),
        "extra_componenti_per_copertura_province": format_number(extra_components_added),
        "province_coperte": int(
            normalize_province_series(selected[PROVINCIA_COL]).replace(NULL_VALUE, pd.NA).dropna().nunique()
        ) if PROVINCIA_COL in selected.columns else 0,
    }
    return selected, pd.DataFrame(diag_rows), meta


def build_sample_2a(sample_2_df: pd.DataFrame, rng: random.Random):
    subset = sample_2_df[(sample_2_df[MOTIV_COL] != NULL_VALUE) & (sample_2_df[MOTIV_COL] != BALNEARE_VALUE)].copy()
    if subset.empty:
        raise ValueError("Campione 2a non costruibile: nessun record non balneare")

    subset["gruppo_2a"] = "NO_BALNEARE"
    selected, diag, meta = select_equal_rows_by_age(
        subset,
        sample_name="campione_2a_no_balneare",
        group_col="gruppo_2a",
        group_values=["NO_BALNEARE"],
        rng=rng,
    )
    selected, added_provinces, missing_provinces, extra_components_added = ensure_all_provinces_covered(
        selected,
        subset,
        strata_cols=[ETA_COL],
        relaxed_strata_cols=[],
    )
    if missing_provinces:
        raise ValueError(f"Campione 2a non costruibile: province non copribili {missing_provinces}")
    final_age_bands = get_valid_values_by_components(selected, ETA_COL)
    meta["province_aggiunte_per_copertura"] = ", ".join(added_provinces)
    meta["province_non_copribili"] = ", ".join(missing_provinces)
    meta["extra_componenti_per_copertura_province"] = format_number(extra_components_added)
    meta["province_coperte"] = int(
        normalize_province_series(selected[PROVINCIA_COL]).replace(NULL_VALUE, pd.NA).dropna().nunique()
    ) if PROVINCIA_COL in selected.columns else 0
    meta["componenti_totali_selezionati"] = format_number(float(selected["_componenti_num"].sum()))
    meta["fasce_eta_nucleo_bilanciato"] = meta.get("fasce_eta", "")
    meta["fasce_eta_presenti_finali"] = ", ".join(final_age_bands)
    return selected, diag, meta


def build_sample_2b(sample_2_df: pd.DataFrame, rng: random.Random):
    subset = sample_2_df[sample_2_df["_durata_num"] <= 30].copy()
    subset["classe_durata_30"] = build_duration_class(subset[DURATA_COL])
    subset["destinazione_3grp"] = build_destination_group(subset[DEST_COL])
    subset = subset[subset["classe_durata_30"].notna()].copy()
    subset = subset[subset["destinazione_3grp"].notna()].copy()
    if subset.empty:
        raise ValueError("Campione 2b non costruibile: nessuna durata valida tra 1 e 30")

    duration_classes = ["1-3", "4-7", "8-14", "15-30"]
    destination_groups = ["SARDEGNA", "ITALIA", "ESTERO"]
    present = set(get_valid_values_by_components(subset, "classe_durata_30"))
    missing = [c for c in duration_classes if c not in present]
    if missing:
        raise ValueError(f"Campione 2b non costruibile: classi durata mancanti {missing}")

    missing_pairs: list[str] = []
    reserved_idx: set[int] = set()
    reserved_components_by_duration: dict[str, float] = {cls: 0.0 for cls in duration_classes}
    reserved_combo_labels: list[str] = []

    for duration_class in duration_classes:
        for dest_group in destination_groups:
            cell = subset[
                (subset["classe_durata_30"] == duration_class)
                & (subset["destinazione_3grp"] == dest_group)
            ].copy()
            if cell.empty:
                missing_pairs.append(f"{dest_group}:{duration_class}")
                continue

            # Reserve the lightest row in each destination x duration cell to guarantee coverage
            # while perturbing the component balance as little as possible.
            cell = cell.sort_values(
                by=["_componenti_num", ETA_COL, "_id_text", "_row_order"],
                ascending=[True, True, True, True],
                kind="mergesort",
            )
            chosen_idx = int(cell.index[0])
            if chosen_idx in reserved_idx:
                continue
            reserved_idx.add(chosen_idx)
            reserved_components_by_duration[duration_class] += float(subset.loc[chosen_idx, "_componenti_num"])
            reserved_combo_labels.append(f"{dest_group}:{duration_class}")

    if missing_pairs:
        raise ValueError(
            "Campione 2b non costruibile: combinazioni destinazione x durata mancanti "
            f"{missing_pairs}"
        )

    plans: list[tuple[str, float]] = []
    for duration_class in duration_classes:
        cell = subset[subset["classe_durata_30"] == duration_class].copy()
        available_components = float(cell["_componenti_num"].sum())
        if available_components <= 0:
            raise ValueError(f"Campione campione_2b_durata_1_30 non costruibile: nessun componente per gruppo {duration_class}")
        plans.append((duration_class, available_components))

    target_components = min(available_components for _, available_components in plans)
    if target_components <= 0:
        raise ValueError("Campione campione_2b_durata_1_30 non costruibile: target componenti = 0")

    selected_indices_all: list[pd.Index] = []
    diag_rows: list[dict[str, object]] = []

    for duration_class, available_components in plans:
        cell = subset[subset["classe_durata_30"] == duration_class].copy()
        reserved_cell = cell[cell.index.isin(reserved_idx)].copy()
        reserved_components = float(reserved_cell["_componenti_num"].sum())
        remaining_target = max(target_components - reserved_components, 0.0)
        pool = cell[~cell.index.isin(reserved_idx)].copy()

        extra_idx, extra_components = select_cell_rows_with_optimal_deviation(
            pool,
            remaining_target,
            rng=rng,
        )
        selected_idx = reserved_cell.index.append(extra_idx)
        selected_components = reserved_components + extra_components
        selected_indices_all.append(selected_idx)
        diag_rows.append(
            {
                "campione": "campione_2b_durata_1_30",
                "dimensione_gruppo": "classe_durata_30",
                "valore_gruppo": duration_class,
                "componenti_disponibili": format_number(available_components),
                "target_componenti": format_number(target_components),
                "componenti_riservati_destinazione_durata": format_number(reserved_components),
                "componenti_selezionati": format_number(selected_components),
                "overshoot_componenti": format_number(selected_components - target_components),
            }
        )

    final_idx = pd.Index([], dtype="int64")
    for idx in selected_indices_all:
        final_idx = final_idx.append(idx)

    selected = subset.loc[final_idx].copy()
    selected["campione"] = "campione_2b_durata_1_30"
    diag_df = pd.DataFrame(diag_rows)
    meta = {
        "campione": "campione_2b_durata_1_30",
        "dimensione_gruppo": "classe_durata_30",
        "n_gruppi": len(duration_classes),
        "target_componenti_per_cella": format_number(target_components),
        "componenti_totali_selezionati": format_number(float(selected["_componenti_num"].sum())),
    }
    selected, added_provinces, missing_provinces, extra_components_added = ensure_all_provinces_covered(
        selected,
        subset,
        strata_cols=[],
        relaxed_strata_cols=[],
        append_only=True,
    )
    if missing_provinces:
        raise ValueError(f"Campione 2b non costruibile: province non copribili {missing_provinces}")
    selected["campione"] = "campione_2b_durata_1_30"
    selected = selected.sort_values(by=["classe_durata_30", ETA_COL, "_id_text", "_row_order"], kind="mergesort")
    final_age_bands = get_valid_values_by_components(selected, ETA_COL)

    meta["classi_durata"] = ", ".join(duration_classes)
    meta["componenti_totali_selezionati"] = format_number(float(selected["_componenti_num"].sum()))
    meta["fasce_eta_presenti_finali"] = ", ".join(final_age_bands)
    meta["copertura_minima_destinazione_x_durata"] = ", ".join(reserved_combo_labels)
    meta["n_record_riservati_destinazione_x_durata"] = len(reserved_idx)
    meta["province_aggiunte_per_copertura"] = ", ".join(added_provinces)
    meta["province_non_copribili"] = ", ".join(missing_provinces)
    meta["extra_componenti_per_copertura_province"] = format_number(extra_components_added)
    meta["province_coperte"] = int(
        normalize_province_series(selected[PROVINCIA_COL]).replace(NULL_VALUE, pd.NA).dropna().nunique()
    ) if PROVINCIA_COL in selected.columns else 0
    return selected, diag_df, meta


def build_sample_2c(sample_2_df: pd.DataFrame, rng: random.Random):
    subset = sample_2_df[sample_2_df[MOTIV_COL] != NULL_VALUE].copy()
    if subset.empty:
        raise ValueError("Campione 2c non costruibile: nessuna motivazione principale valorizzata")

    top5_stats = (
        subset.groupby(MOTIV_COL, as_index=False)
        .agg(componenti=("_componenti_num", "sum"))
    )
    top5_stats["_motiv_priority"] = top5_stats[MOTIV_COL].map(motivation_sort_priority_sardinia)
    top5_stats = top5_stats.sort_values(
        by=["componenti", "_motiv_priority", MOTIV_COL],
        ascending=[False, True, True],
        kind="mergesort",
    )
    top5 = top5_stats[MOTIV_COL].astype("string").head(5).tolist()
    if len(top5) < 5:
        raise ValueError(
            "Campione 2c non costruibile: motivazioni con tutte le fasce eta popolate "
            f"{len(top5)} < 5"
        )

    subset = subset[subset[MOTIV_COL].isin(top5)].copy()

    selected, diag, meta = select_equal_rows_by_group_components(
        subset,
        sample_name="campione_2c_top5_motivazioni",
        group_col=MOTIV_COL,
        group_values=top5,
        rng=rng,
        ensure_province_coverage=True,
    )
    selected, added_provinces, missing_provinces, extra_components_added = ensure_all_provinces_covered(
        selected,
        subset,
        strata_cols=[],
        relaxed_strata_cols=[],
        append_only=True,
    )
    if missing_provinces:
        raise ValueError(f"Campione 2c non costruibile: province non copribili {missing_provinces}")
    final_age_bands = get_valid_values_by_components(selected, ETA_COL)
    meta["top5_motivazioni_principali"] = ", ".join(top5)
    meta["province_aggiunte_per_copertura"] = ", ".join(added_provinces)
    meta["province_non_copribili"] = ", ".join(missing_provinces)
    meta["extra_componenti_per_copertura_province"] = format_number(extra_components_added)
    meta["province_coperte"] = int(
        normalize_province_series(selected[PROVINCIA_COL]).replace(NULL_VALUE, pd.NA).dropna().nunique()
    ) if PROVINCIA_COL in selected.columns else 0
    meta["componenti_totali_selezionati"] = format_number(float(selected["_componenti_num"].sum()))
    meta["fasce_eta_presenti_finali"] = ", ".join(final_age_bands)
    return selected, diag, meta


def split_output_columns(df: pd.DataFrame, original_columns: list[str]) -> pd.DataFrame:
    return df[[c for c in original_columns if c in df.columns]].copy()


def run(input_file: Path, output_file: Path, seed: int | None = None) -> Path:
    if not input_file.exists():
        raise FileNotFoundError(f"File input non trovato: {input_file}")

    xls = pd.ExcelFile(input_file)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(input_file, sheet_name=sheet)
    validate_columns(df)
    original_columns = df.columns.tolist()

    base = prepare_base(df)
    if base.empty:
        raise ValueError("Nessun record eleggibile dopo i filtri base")

    rng = random.Random(seed)

    s1_df, s1_diag, s1_meta = build_sample_1(base, rng)
    s2_df, s2_diag, s2_meta = build_sample_2(base, rng)
    s2a_df, s2a_diag, s2a_meta = build_sample_2a(s2_df, rng)
    s2b_df, s2b_diag, s2b_meta = build_sample_2b(s2_df, rng)
    s2c_df, s2c_diag, s2c_meta = build_sample_2c(s2_df, rng)

    diagnostics = pd.concat([s1_diag, s2_diag, s2a_diag, s2b_diag, s2c_diag], ignore_index=True)

    meta_rows = [
        {"chiave": "input_file", "valore": str(input_file)},
        {"chiave": "sheet", "valore": sheet},
        {"chiave": "componenti_eleggibili_senza_pacchetto", "valore": format_number(float(base["_componenti_num"].sum()))},
        {"chiave": "seed", "valore": "AUTO" if seed is None else seed},
    ]

    for m in [s1_meta, s2_meta, s2a_meta, s2b_meta, s2c_meta]:
        campione = str(m["campione"])
        for k, v in m.items():
            if k == "campione":
                continue
            meta_rows.append({"chiave": f"{campione}.{k}", "valore": v})

    meta_df = pd.DataFrame(meta_rows)

    out_path = output_file
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            split_output_columns(s1_df, original_columns).to_excel(writer, sheet_name="campione_1", index=False)
            split_output_columns(s2_df, original_columns).to_excel(writer, sheet_name="campione_2", index=False)
            split_output_columns(s2a_df, original_columns).to_excel(writer, sheet_name="campione_2a", index=False)
            split_output_columns(s2b_df, original_columns).to_excel(writer, sheet_name="campione_2b", index=False)
            split_output_columns(s2c_df, original_columns).to_excel(writer, sheet_name="campione_2c", index=False)
            diagnostics.to_excel(writer, sheet_name="diagnostica_celle", index=False)
            meta_df.to_excel(writer, sheet_name="meta", index=False)
    except PermissionError:
        out_path = output_file.with_name(f"{output_file.stem}_nuovo{output_file.suffix}")
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            split_output_columns(s1_df, original_columns).to_excel(writer, sheet_name="campione_1", index=False)
            split_output_columns(s2_df, original_columns).to_excel(writer, sheet_name="campione_2", index=False)
            split_output_columns(s2a_df, original_columns).to_excel(writer, sheet_name="campione_2a", index=False)
            split_output_columns(s2b_df, original_columns).to_excel(writer, sheet_name="campione_2b", index=False)
            split_output_columns(s2c_df, original_columns).to_excel(writer, sheet_name="campione_2c", index=False)
            diagnostics.to_excel(writer, sheet_name="diagnostica_celle", index=False)
            meta_df.to_excel(writer, sheet_name="meta", index=False)

    print(f"Input letto: {input_file}")
    print(f"Componenti eleggibili senza pacchetto: {format_number(float(base['_componenti_num'].sum()))}")
    print(f"Campione 1 - componenti: {format_number(float(s1_df['_componenti_num'].sum()))}")
    print(f"Campione 2 - componenti: {format_number(float(s2_df['_componenti_num'].sum()))}")
    print(f"Campione 2a - componenti: {format_number(float(s2a_df['_componenti_num'].sum()))}")
    print(f"Campione 2b - componenti: {format_number(float(s2b_df['_componenti_num'].sum()))}")
    print(f"Campione 2c - componenti: {format_number(float(s2c_df['_componenti_num'].sum()))}")
    print(f"Output generato: {out_path}")
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Estrae sottocampioni dal file sardi, escludendo turisti con pacchetto e bilanciando per fascia eta."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--seed", type=int, default=None)
    args = parser.parse_args()

    run(args.input, args.output, seed=args.seed)


if __name__ == "__main__":
    main()
