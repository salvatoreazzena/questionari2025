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
ETA_COL = "fascia_età"
DURATA_COL = "durata_soggiorno"
MOTIV_COL = "motivazione_principale"

NULL_VALUE = "NULL"
NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}
SI_VALUES = {"SI", "SÌ"}
NO_VALUES = {"NO"}
BALNEARE_VALUE = "BALNEARE"

REQUIRED_COLUMNS = [VACANZA_COL, PACCHETTO_COL, ETA_COL, DURATA_COL, MOTIV_COL]
PROVINCIA_COL = "provincia_provenienza"


@dataclass(frozen=True)
class CellPlan:
    group_value: str
    fascia_eta: str
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


def normalize_province_series(series: pd.Series) -> pd.Series:
    return normalize_upper(series)


def prepare_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out[VACANZA_COL] = normalize_upper(out[VACANZA_COL])
    out[ETA_COL] = build_age_sampling_band(out[ETA_COL])
    out[MOTIV_COL] = normalize_upper(out[MOTIV_COL])
    out["_durata_num"] = pd.to_numeric(out[DURATA_COL], errors="coerce")

    out["_row_order"] = out.index.astype("int64")
    if ID_COL in out.columns:
        out["_id_text"] = out[ID_COL].astype("string").fillna("")
    else:
        out["_id_text"] = (out.index + 2).astype("string")

    mask = (
        out[PACCHETTO_COL].map(is_empty_or_na)
        & (out[ETA_COL] != NULL_VALUE)
        & out[VACANZA_COL].isin(SI_VALUES | NO_VALUES)
    )
    return out.loc[mask].copy()


def compute_common_age_bands(df: pd.DataFrame, group_col: str, group_values: list[str]) -> list[str]:
    sets: list[set[str]] = []
    for gv in group_values:
        age_vals = set(df.loc[df[group_col] == gv, ETA_COL].dropna().astype("string").tolist())
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
            n = int(len(df[(df[group_col] == gv) & (df[ETA_COL] == age)]))
            plans.append(CellPlan(gv, age, n))

    target_n = min(p.available_rows for p in plans)
    if target_n <= 0:
        raise ValueError(f"Campione {sample_name} non costruibile: target per cella = {target_n}")

    selected_chunks: list[pd.DataFrame] = []
    diag_rows: list[dict[str, object]] = []

    for p in plans:
        cell = df[(df[group_col] == p.group_value) & (df[ETA_COL] == p.fascia_eta)].copy()
        sampled = cell.sample(n=target_n, random_state=rng.randint(0, 10**9), replace=False)
        selected_chunks.append(sampled)
        diag_rows.append(
            {
                "campione": sample_name,
                "dimensione_gruppo": group_col,
                "valore_gruppo": p.group_value,
                "fascia_eta": p.fascia_eta,
                "questionari_disponibili": p.available_rows,
                "target_questionari": target_n,
                "questionari_selezionati": int(len(sampled)),
            }
        )

    selected = pd.concat(selected_chunks, ignore_index=False)
    selected["campione"] = sample_name
    selected = selected.sort_values(by=[group_col, ETA_COL, "_id_text", "_row_order"], kind="mergesort")

    meta = {
        "campione": sample_name,
        "dimensione_gruppo": group_col,
        "n_gruppi": len(group_values),
        "n_fasce_eta": len(age_bands),
        "target_questionari_per_cella": int(target_n),
        "questionari_totali": int(len(selected)),
        "fasce_eta": ", ".join(age_bands),
    }
    return selected, pd.DataFrame(diag_rows), meta


def ensure_all_provinces_covered(
    selected: pd.DataFrame,
    pool: pd.DataFrame,
) -> tuple[pd.DataFrame, list[str]]:
    if PROVINCIA_COL not in selected.columns or PROVINCIA_COL not in pool.columns:
        return selected, []

    sel = selected.copy()
    pool2 = pool.copy()
    sel["_prov_norm"] = normalize_province_series(sel[PROVINCIA_COL])
    pool2["_prov_norm"] = normalize_province_series(pool2[PROVINCIA_COL])

    valid_mask_pool = pool2["_prov_norm"].notna() & (pool2["_prov_norm"] != NULL_VALUE)
    required_provinces = sorted(set(pool2.loc[valid_mask_pool, "_prov_norm"].astype("string").tolist()))
    if not required_provinces:
        return selected, []

    added: list[str] = []

    while True:
        current = set(sel.loc[sel["_prov_norm"].notna() & (sel["_prov_norm"] != NULL_VALUE), "_prov_norm"].astype("string").tolist())
        missing = [p for p in required_provinces if p not in current]
        if not missing:
            break

        target_prov = missing[0]
        candidates = pool2[(pool2["_prov_norm"] == target_prov) & (~pool2.index.isin(sel.index))].copy()
        if candidates.empty:
            break
        candidates = candidates.sort_values(by=["_id_text", "_row_order"], kind="mergesort")

        replaced = False
        for cand_idx, cand_row in candidates.iterrows():
            age = str(cand_row[ETA_COL])
            same_age = sel[sel[ETA_COL] == age].copy()
            if same_age.empty:
                continue

            prov_counts = same_age["_prov_norm"].value_counts(dropna=False)
            removable = same_age[
                same_age["_prov_norm"].map(lambda p: prov_counts.get(p, 0) > 1 or pd.isna(p) or p == NULL_VALUE)
            ].copy()
            if removable.empty:
                continue

            removable = removable.sort_values(by=["_id_text", "_row_order"], ascending=[False, False], kind="mergesort")
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
            break

        if not replaced:
            break

    sel = sel.drop(columns=["_prov_norm"], errors="ignore")
    return sel, sorted(set(added))


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

    # Campione 2: bilanciato per fascia eta + copertura di tutte le province valide.
    subset["_prov_norm"] = normalize_province_series(subset[PROVINCIA_COL])
    subset["_is_prov_valid"] = subset["_prov_norm"].notna() & (subset["_prov_norm"] != NULL_VALUE)
    subset = subset[subset["_is_prov_valid"]].copy()
    if subset.empty:
        raise ValueError("Campione 2 non costruibile: nessuna provincia valida dopo esclusione ND/NULL")

    age_bands = sorted(subset[ETA_COL].astype("string").unique().tolist())
    target_n = min(int(len(subset[subset[ETA_COL] == a])) for a in age_bands)
    if target_n <= 0:
        raise ValueError("Campione 2 non costruibile: target per fascia eta nullo")

    required_provinces = sorted(
        set(subset.loc[subset["_is_prov_valid"], "_prov_norm"].astype("string").tolist())
    )

    reserved_idx: set[int] = set()
    reserved_per_age: dict[str, int] = {str(a): 0 for a in age_bands}
    prov_freq = subset.loc[subset["_is_prov_valid"], "_prov_norm"].value_counts().to_dict()

    for prov in sorted(required_provinces, key=lambda p: (prov_freq.get(p, 0), str(p))):
        candidates = subset[(subset["_prov_norm"] == prov) & (~subset.index.isin(reserved_idx))].copy()
        if candidates.empty:
            continue
        candidates = candidates.sort_values(by=["_id_text", "_row_order"], kind="mergesort")

        chosen_idx = None
        chosen_key = None
        for idx, row in candidates.iterrows():
            age = str(row[ETA_COL])
            if reserved_per_age.get(age, 0) >= target_n:
                continue
            key = (reserved_per_age.get(age, 0), str(age), str(row["_id_text"]))
            if chosen_key is None or key < chosen_key:
                chosen_key = key
                chosen_idx = int(idx)
        if chosen_idx is None:
            continue

        reserved_idx.add(chosen_idx)
        chosen_age = str(subset.loc[chosen_idx, ETA_COL])
        reserved_per_age[chosen_age] += 1

    selected_chunks: list[pd.DataFrame] = []
    diag_rows: list[dict[str, object]] = []

    for age in age_bands:
        cell = subset[subset[ETA_COL] == age].copy()
        base_reserved = cell.loc[cell.index.isin(reserved_idx)].copy()
        need = target_n - len(base_reserved)
        if need < 0:
            base_reserved = base_reserved.sample(n=target_n, random_state=rng.randint(0, 10**9), replace=False)
            need = 0
        pool = cell.loc[~cell.index.isin(base_reserved.index)].copy()
        if len(pool) < need:
            raise ValueError(f"Campione 2 non costruibile per fascia eta {age}: disponibilita insufficiente")
        extra = pool.sample(n=need, random_state=rng.randint(0, 10**9), replace=False)
        sampled = pd.concat([base_reserved, extra], ignore_index=False)
        selected_chunks.append(sampled)

        diag_rows.append(
            {
                "campione": "campione_2_vacanza",
                "dimensione_gruppo": "gruppo_vacanza",
                "valore_gruppo": "VACANZA",
                "fascia_eta": age,
                "questionari_disponibili": int(len(cell)),
                "target_questionari": int(target_n),
                "questionari_selezionati": int(len(sampled)),
            }
        )

    selected = pd.concat(selected_chunks, ignore_index=False)
    selected["campione"] = "campione_2_vacanza"
    selected = selected.sort_values(by=[ETA_COL, "_id_text", "_row_order"], kind="mergesort")

    meta = {
        "campione": "campione_2_vacanza",
        "dimensione_gruppo": "gruppo_vacanza",
        "n_gruppi": 1,
        "n_fasce_eta": len(age_bands),
        "target_questionari_per_cella": int(target_n),
        "questionari_totali": int(len(selected)),
        "fasce_eta": ", ".join(age_bands),
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
    return select_equal_rows_by_age(
        subset,
        sample_name="campione_2a_no_balneare",
        group_col="gruppo_2a",
        group_values=["NO_BALNEARE"],
        rng=rng,
    )


def build_sample_2b(sample_2_df: pd.DataFrame, rng: random.Random):
    subset = sample_2_df[sample_2_df["_durata_num"] <= 30].copy()
    subset["classe_durata_30"] = build_duration_class(subset[DURATA_COL])
    subset = subset[subset["classe_durata_30"].notna()].copy()
    if subset.empty:
        raise ValueError("Campione 2b non costruibile: nessuna durata valida tra 1 e 30")

    duration_classes = ["1-3", "4-7", "8-14", "15-30"]
    present = set(subset["classe_durata_30"].astype("string").tolist())
    missing = [c for c in duration_classes if c not in present]
    if missing:
        raise ValueError(f"Campione 2b non costruibile: classi durata mancanti {missing}")

    age_bands = compute_common_age_bands(subset, "classe_durata_30", duration_classes)

    target_n = min(
        int(len(subset[(subset["classe_durata_30"] == d) & (subset[ETA_COL] == a)]))
        for d in duration_classes
        for a in age_bands
    )
    if target_n <= 0:
        raise ValueError("Campione 2b non costruibile: target per cella nullo")

    selected_chunks = []
    diag_rows = []
    for d in duration_classes:
        for a in age_bands:
            cell = subset[(subset["classe_durata_30"] == d) & (subset[ETA_COL] == a)].copy()
            sampled = cell.sample(n=target_n, random_state=rng.randint(0, 10**9), replace=False)
            selected_chunks.append(sampled)
            diag_rows.append(
                {
                    "campione": "campione_2b_durata_1_30",
                    "dimensione_gruppo": "classe_durata_30",
                    "valore_gruppo": d,
                    "fascia_eta": a,
                    "questionari_disponibili": int(len(cell)),
                    "target_questionari": int(target_n),
                    "questionari_selezionati": int(len(sampled)),
                }
            )

    selected = pd.concat(selected_chunks, ignore_index=False)
    selected["campione"] = "campione_2b_durata_1_30"
    selected = selected.sort_values(by=["classe_durata_30", ETA_COL, "_id_text", "_row_order"], kind="mergesort")

    meta = {
        "campione": "campione_2b_durata_1_30",
        "dimensione_gruppo": "classe_durata_30",
        "n_gruppi": len(duration_classes),
        "n_fasce_eta": len(age_bands),
        "target_questionari_per_cella": int(target_n),
        "questionari_totali": int(len(selected)),
        "classi_durata": ", ".join(duration_classes),
        "fasce_eta": ", ".join(age_bands),
    }
    return selected, pd.DataFrame(diag_rows), meta


def build_sample_2c(sample_2_df: pd.DataFrame, rng: random.Random):
    subset = sample_2_df[sample_2_df[MOTIV_COL] != NULL_VALUE].copy()
    if subset.empty:
        raise ValueError("Campione 2c non costruibile: nessuna motivazione principale valorizzata")

    top5 = (
        subset[MOTIV_COL]
        .value_counts()
        .sort_values(ascending=False)
        .head(5)
        .index.astype("string")
        .tolist()
    )
    if len(top5) < 5:
        raise ValueError(f"Campione 2c non costruibile: motivazioni disponibili {len(top5)} < 5")

    subset = subset[subset[MOTIV_COL].isin(top5)].copy()

    selected, diag, meta = select_equal_rows_by_age(
        subset,
        sample_name="campione_2c_top5_motivazioni",
        group_col=MOTIV_COL,
        group_values=top5,
        rng=rng,
    )
    meta["top5_motivazioni_principali"] = ", ".join(top5)
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
        {"chiave": "record_eleggibili_senza_pacchetto", "valore": int(len(base))},
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
    print(f"Record eleggibili senza pacchetto: {len(base)}")
    print(f"Campione 1: {len(s1_df)}")
    print(f"Campione 2: {len(s2_df)}")
    print(f"Campione 2a: {len(s2a_df)}")
    print(f"Campione 2b: {len(s2b_df)}")
    print(f"Campione 2c: {len(s2c_df)}")
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
