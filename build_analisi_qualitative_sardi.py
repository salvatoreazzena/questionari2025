from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd
from build_analisi_qualitative_campione1 import write_structured_excel

DEFAULT_INPUT = Path("questionari_sardi_sottocampioni.xlsx")
DEFAULT_OUTPUT = Path("analisi_qualitative_sardi.xlsx")
SHEET_CAMPIONE_1 = "campione_1"
SHEET_CAMPIONE_2 = "campione_2"

NULL_VALUE = "NULL"
UNDEFINED_LABEL = "NON DEFINITO"
NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}

ID_COL = "ID"
PROVINCIA_COL = "provincia_provenienza"
VACANZA_COL = "vacanza_2025"
DOVE_COL = "si_dove"
NO_MOTIVO_COL = "no_motivo"
LOCALITA_SARDEGNA_COL = "località_sardegna"
COMPONENTI_COL = "numero_componenti"
WEB_COL = "web"
WEB_DETAIL_COL = "web_specifica"
LODGING_COL = "tipologia_alloggio"
PRIMARY_REASON_COL = "motivazione_principale"
SECONDARY_REASON_COL = "motivazione_secondaria"
JUDGMENT_COL = "giudizio"
FUTURE_SPEND_COL = "previsione_spesa"

REQUIRED_COLUMNS = [
    ID_COL,
    PROVINCIA_COL,
    VACANZA_COL,
    DOVE_COL,
    NO_MOTIVO_COL,
    LOCALITA_SARDEGNA_COL,
    COMPONENTI_COL,
    WEB_COL,
    WEB_DETAIL_COL,
    LODGING_COL,
    PRIMARY_REASON_COL,
    SECONDARY_REASON_COL,
    JUDGMENT_COL,
    FUTURE_SPEND_COL,
]

DEST_SARDEGNA = "SARDI IN VACANZA IN SARDEGNA"
DEST_ITALIA = "SARDI IN VACANZA IN ALTRE REGIONI ITALIANE"
DEST_ESTERO = "SARDI IN VACANZA ALL'ESTERO"
DEST_UNDEFINED = "VACANZA (DESTINAZIONE NON DEFINITA)"

JUDGMENT_SCORE_MAP = {
    "PESSIMO": 1,
    "SUFFICIENTE": 2,
    "BUONO": 3,
    "OTTIMO": 4,
}

PROVINCIA_WEIGHT_MAP = {
    "CA": 0.83,
    "SS": 0.5,
    "OT": 1.53,
    "OR": 8.53,
    "NU": 1.02,
    "CI": 12.08,
    "SU": 12.08,
    "VS": 30.32,
    "OG": 5.48,
}

WEB_NEGATION_PATTERNS = [
    r"\b(non ho prenotato|nessuna prenotazione|non prenotato|non usato il web)\b",
]

WEB_THEME_PATTERNS: list[tuple[str, str]] = [
    ("VOLO/AEREO", r"\b(vol\w*|aere\w*|flight\w*)\b"),
    ("NAVE/TRAGHETTO", r"\b(nav\w*|traghett\w*|ferr\w*)\b"),
    (
        "ALLOGGIO",
        r"\b(allogg\w*|hotel\w*|b&b|appartament\w*|casa vacanz\w*|campegg\w*|residence\w*|agriturism\w*|ostell\w*|camera)\b",
    ),
    ("AUTONOLEGGIO/TRASPORTI", r"\b(auto\b|nolegg\w*|scooter\w*|moto\w*|bus\w*|autobus\w*|tren\w*|transfer\w*|taxi\w*)\b"),
    ("ESCURSIONI/ATTIVITA", r"\b(escursion\w*|git\w*|tour\w*|barc\w*|boat\w*|attivit\w*|esperienz\w*|bigliett\w*|ticket\w*|muse\w*|parc\w*|visit\w*|ingress\w*)\b"),
    ("RISTORANTI", r"\b(ristor\w*|pizzeri\w*|cen\w*|pranz\w*|food\w*|colazion\w*)\b"),
]

TEXT_NORMALIZATION_REPLACEMENTS = [
    (r"\bb b\b", "b&b"),
    (r"\bbnb\b", "b&b"),
    (r"\bbed and breakfast\b", "b&b"),
    (r"\bcasa vacanza\b", "casa vacanze"),
    (r"\bcase vacanza\b", "casa vacanze"),
    (r"\bautonoleggio\b", "noleggio auto"),
    (r"\bauto noleggio\b", "noleggio auto"),
]

TEXT_SPLIT_RE = re.compile(r"[;,/|\n]+|\be\b|\bed\b|\by\b", flags=re.IGNORECASE)

LOCALITA_SARDEGNA_VARIANT_MAP = {
    "S. TEODORO": "SAN TEODORO",
    "S TEODORO": "SAN TEODORO",
    "S. TERESA": "SANTA TERESA DI GALLURA",
    "S TERESA": "SANTA TERESA DI GALLURA",
}


def normalize_text_series(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("").str.strip()
    upper = s.str.upper()
    mask_null = upper.isin({token.upper() for token in NULL_TOKENS})
    return s.mask(mask_null, NULL_VALUE)


def normalize_upper_series(series: pd.Series) -> pd.Series:
    return normalize_text_series(series).astype("string").str.upper()


def prettify_value(value: object) -> str:
    if pd.isna(value):
        return UNDEFINED_LABEL
    text = str(value).strip()
    if not text or text.upper() == NULL_VALUE:
        return UNDEFINED_LABEL
    return text


def prettify_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = out[col].map(prettify_value)
    return out


def strip_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize_free_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text or text.upper() in {token.upper() for token in NULL_TOKENS}:
        return ""
    text = strip_accents(text.lower())
    for pattern, repl in TEXT_NORMALIZATION_REPLACEMENTS:
        text = re.sub(pattern, repl, text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def classify_text_fragments(
    text: object,
    *,
    patterns: list[tuple[str, str]],
    extra_negation_patterns: list[str] | None = None,
) -> list[str]:
    normalized = normalize_free_text(text)
    if not normalized:
        return []

    negation_patterns = list(extra_negation_patterns or [])
    if any(re.search(pattern, normalized, flags=re.IGNORECASE) for pattern in negation_patterns):
        return []

    themes: list[str] = []
    parts = [p.strip(" .,:;-") for p in TEXT_SPLIT_RE.split(normalized)]
    for part in parts:
        if not part:
            continue
        matched = False
        for label, pattern in patterns:
            if re.search(pattern, part, flags=re.IGNORECASE):
                themes.append(label)
                matched = True
        if not matched and len(parts) == 1:
            themes.append("ALTRO/NON CLASSIFICATO")
    return list(dict.fromkeys(themes))


def validate_columns(df: pd.DataFrame) -> None:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel dataset: {missing}")


def normalize_yes_no(series: pd.Series) -> pd.Series:
    upper = normalize_upper_series(series)
    out = pd.Series(UNDEFINED_LABEL, index=upper.index, dtype="string")
    out = out.mask(upper.isin({"SI", "SÌ"}), "SI")
    out = out.mask(upper.eq("NO"), "NO")
    return out


def build_destination_category(vacanza: pd.Series, dove: pd.Series) -> pd.Series:
    vac = normalize_upper_series(vacanza)
    where = normalize_upper_series(dove)
    out = pd.Series(DEST_UNDEFINED, index=vac.index, dtype="string")
    out = out.mask(vac.eq("NO"), "NO VACANZA")
    out = out.mask(vac.isin({"SI", "SÌ"}) & where.eq("IN SARDEGNA"), DEST_SARDEGNA)
    out = out.mask(vac.isin({"SI", "SÌ"}) & where.eq("IN ITALIA"), DEST_ITALIA)
    out = out.mask(vac.isin({"SI", "SÌ"}) & where.eq("ALL'ESTERO"), DEST_ESTERO)
    return out


def normalize_localita_sardegna_prevalente(series: pd.Series) -> pd.Series:
    s = normalize_text_series(series)
    s = s.str.replace(r"\s*\([^)]*\)\s*", " ", regex=True)
    s = s.str.split(r"[;,/]", n=1, regex=True).str[0].str.strip()
    s = s.str.split(",", n=1).str[0].str.strip()
    s = s.str.replace(r"\s+", " ", regex=True)
    s = s.mask(s.eq(""), NULL_VALUE)
    out = s.astype("string").str.upper().fillna(NULL_VALUE)
    out = out.replace(LOCALITA_SARDEGNA_VARIANT_MAP)
    out_ascii = out.map(lambda value: strip_accents(str(value)) if pd.notna(value) else value)
    out = out.mask(out.isin(LOCALITA_SARDEGNA_VARIANT_MAP.keys()), out.replace(LOCALITA_SARDEGNA_VARIANT_MAP))
    out = out.mask(out_ascii.isin(LOCALITA_SARDEGNA_VARIANT_MAP.keys()), out_ascii.replace(LOCALITA_SARDEGNA_VARIANT_MAP))
    return out.astype("string")


def add_share_within_group(df: pd.DataFrame, group_cols: list[str], value_col: str = "questionari") -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        out["pct_su_gruppo"] = pd.Series(dtype="float64")
        return out
    totals = out.groupby(group_cols, dropna=False)[value_col].transform("sum")
    out["pct_su_gruppo"] = (out[value_col] / totals * 100.0).round(2)
    return out


def build_distribution(
    df: pd.DataFrame,
    *,
    group_cols: list[str],
    value_col: str,
    sort_cols: list[str] | None = None,
) -> pd.DataFrame:
    out = (
        df.groupby([*group_cols, value_col], dropna=False, as_index=False)
        .agg(
            questionari=(ID_COL, "nunique"),
            componenti=(COMPONENTI_COL, "sum"),
            componenti_pesati=("componenti_pesati", "sum"),
        )
    )
    sort_by = sort_cols or [*group_cols, "questionari", value_col]
    ascending = [True] * len(group_cols) + [False, True]
    out = out.sort_values(by=sort_by, ascending=ascending[: len(sort_by)], kind="mergesort")
    out = add_share_within_group(out, group_cols)
    return prettify_columns(out, [*group_cols, value_col])


def build_multi_assoc(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = (
        df.groupby(cols, dropna=False, as_index=False)
        .agg(
            questionari=(ID_COL, "nunique"),
            componenti=(COMPONENTI_COL, "sum"),
            componenti_pesati=("componenti_pesati", "sum"),
        )
        .sort_values(by=cols + ["questionari"], ascending=[True] * len(cols) + [False], kind="mergesort")
    )
    return prettify_columns(out, cols)


def build_web_detail_outputs(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    web_yes = df[df["web_flag"] == "SI"].copy()
    audit_rows: list[dict[str, object]] = []
    theme_rows: list[dict[str, object]] = []

    for _, row in web_yes.iterrows():
        themes = classify_text_fragments(
            row[WEB_DETAIL_COL],
            patterns=WEB_THEME_PATTERNS,
            extra_negation_patterns=WEB_NEGATION_PATTERNS,
        )
        normalized = normalize_free_text(row[WEB_DETAIL_COL])
        if not themes:
            themes = ["ALTRO/NON CLASSIFICATO"] if normalized else [UNDEFINED_LABEL]
        for theme in themes:
            theme_rows.append(
                {
                    "categoria_destinazione": row["categoria_destinazione"],
                    PROVINCIA_COL: row[PROVINCIA_COL],
                    "cosa_prenotata": theme,
                    ID_COL: row[ID_COL],
                    COMPONENTI_COL: row[COMPONENTI_COL],
                    "peso_provincia": row["peso_provincia"],
                    "componenti_pesati": row["componenti_pesati"],
                }
            )
        audit_rows.append(
            {
                ID_COL: row[ID_COL],
                PROVINCIA_COL: row[PROVINCIA_COL],
                "categoria_destinazione": row["categoria_destinazione"],
                "web_specifica_originale": prettify_value(row[WEB_DETAIL_COL]),
                "web_specifica_normalizzata": normalized or UNDEFINED_LABEL,
                "cose_prenotate_classificate": " | ".join(themes),
            }
        )

    audit_df = pd.DataFrame(audit_rows)
    summary_df = pd.DataFrame(theme_rows)
    if summary_df.empty:
        summary_out = pd.DataFrame(
            columns=["categoria_destinazione", PROVINCIA_COL, "cosa_prenotata", "questionari", "componenti", "componenti_pesati", "pct_su_gruppo"]
        )
    else:
        summary_out = (
            summary_df.groupby(["categoria_destinazione", PROVINCIA_COL, "cosa_prenotata"], dropna=False, as_index=False)
            .agg(
                questionari=(ID_COL, "nunique"),
                componenti=(COMPONENTI_COL, "sum"),
                componenti_pesati=("componenti_pesati", "sum"),
            )
            .sort_values(
                by=["categoria_destinazione", PROVINCIA_COL, "questionari", "cosa_prenotata"],
                ascending=[True, True, False, True],
                kind="mergesort",
            )
        )
        summary_out = add_share_within_group(summary_out, ["categoria_destinazione", PROVINCIA_COL])
    audit_df = prettify_columns(audit_df, [PROVINCIA_COL, "categoria_destinazione"])
    summary_out = prettify_columns(summary_out, ["categoria_destinazione", PROVINCIA_COL, "cosa_prenotata"])
    return audit_df, summary_out


def prepare_dataframe(input_file: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    validate_columns(df)

    for col in REQUIRED_COLUMNS:
        if col == COMPONENTI_COL:
            continue
        df[col] = normalize_text_series(df[col])

    df[COMPONENTI_COL] = pd.to_numeric(df[COMPONENTI_COL], errors="coerce").fillna(0)
    df[PROVINCIA_COL] = normalize_upper_series(df[PROVINCIA_COL])
    df["peso_provincia"] = df[PROVINCIA_COL].map(PROVINCIA_WEIGHT_MAP).fillna(1.0)
    df["categoria_destinazione"] = build_destination_category(df[VACANZA_COL], df[DOVE_COL])
    df["web_flag"] = normalize_yes_no(df[WEB_COL])
    df["giudizio_norm"] = normalize_upper_series(df[JUDGMENT_COL])
    df["giudizio_score"] = df["giudizio_norm"].map(JUDGMENT_SCORE_MAP)
    df["localita_sardegna_prevalente"] = normalize_localita_sardegna_prevalente(df[LOCALITA_SARDEGNA_COL])
    df["componenti_pesati"] = df[COMPONENTI_COL] * df["peso_provincia"]
    df["giudizio_score_pesato"] = df["giudizio_score"] * df["peso_provincia"]

    for col in [NO_MOTIVO_COL, LODGING_COL, PRIMARY_REASON_COL, SECONDARY_REASON_COL, FUTURE_SPEND_COL]:
        df[col] = normalize_upper_series(df[col])

    return df


def build_outputs(campione_1: pd.DataFrame, campione_2: pd.DataFrame) -> dict[str, pd.DataFrame]:
    outputs: dict[str, pd.DataFrame] = {}

    meta_rows = [
        {"chiave": "input_file", "valore": str(DEFAULT_INPUT)},
        {"chiave": "campione_1_sheet", "valore": SHEET_CAMPIONE_1},
        {"chiave": "campione_2_sheet", "valore": SHEET_CAMPIONE_2},
        {"chiave": "campione_1_questionari", "valore": int(campione_1[ID_COL].nunique())},
        {"chiave": "campione_2_questionari", "valore": int(campione_2[ID_COL].nunique())},
        {"chiave": "proxy_localita_sardegna_prevalente", "valore": "prima localita indicata nel campo località_sardegna"},
        {"chiave": "province_con_peso_default_1", "valore": ", ".join(sorted(set(campione_1.loc[~campione_1[PROVINCIA_COL].isin(PROVINCIA_WEIGHT_MAP), PROVINCIA_COL].tolist() + campione_2.loc[~campione_2[PROVINCIA_COL].isin(PROVINCIA_WEIGHT_MAP), PROVINCIA_COL].tolist())))},
    ]
    outputs["meta"] = pd.DataFrame(meta_rows)

    campione_1_valid = campione_1[campione_1["categoria_destinazione"] == "NO VACANZA"].copy()
    outputs["c1_no_vacanza_x_prov"] = build_distribution(
        campione_1_valid,
        group_cols=[PROVINCIA_COL],
        value_col=NO_MOTIVO_COL,
    )

    campione_2_valid = campione_2[
        campione_2["categoria_destinazione"].isin([DEST_SARDEGNA, DEST_ITALIA, DEST_ESTERO])
    ].copy()

    outputs["c2_destinazioni_x_prov"] = build_distribution(
        campione_2_valid,
        group_cols=[PROVINCIA_COL],
        value_col="categoria_destinazione",
    )

    solo_sardegna = campione_2_valid[campione_2_valid["categoria_destinazione"] == DEST_SARDEGNA].copy()
    outputs["c2_localita_sardegna"] = build_distribution(
        solo_sardegna,
        group_cols=[PROVINCIA_COL],
        value_col="localita_sardegna_prevalente",
    )

    outputs["c2_web_x_prov"] = build_distribution(
        campione_2_valid,
        group_cols=["categoria_destinazione", PROVINCIA_COL],
        value_col="web_flag",
    )
    audit_web, summary_web = build_web_detail_outputs(campione_2_valid)
    outputs["audit_web"] = audit_web
    outputs["c2_web_cosa"] = summary_web

    outputs["c2_alloggio_x_prov"] = build_distribution(
        campione_2_valid,
        group_cols=["categoria_destinazione", PROVINCIA_COL],
        value_col=LODGING_COL,
    )

    outputs["c2_motiv_princ_x_prov"] = build_distribution(
        campione_2_valid,
        group_cols=["categoria_destinazione", PROVINCIA_COL],
        value_col=PRIMARY_REASON_COL,
    )

    outputs["c2_spesa_mot_giud"] = build_multi_assoc(
        campione_2_valid,
        ["categoria_destinazione", PROVINCIA_COL, FUTURE_SPEND_COL, PRIMARY_REASON_COL, "giudizio_norm"],
    )

    outputs["c2_motiv1_motiv2"] = build_multi_assoc(
        campione_2_valid,
        ["categoria_destinazione", PROVINCIA_COL, PRIMARY_REASON_COL, SECONDARY_REASON_COL],
    )

    outputs["c2_giudizio_dist"] = build_distribution(
        campione_2_valid,
        group_cols=["categoria_destinazione", PROVINCIA_COL],
        value_col="giudizio_norm",
    )

    solo_sardegna_definite = solo_sardegna[
        ~solo_sardegna["localita_sardegna_prevalente"].isin([UNDEFINED_LABEL, NULL_VALUE, ""])
    ].copy()

    top15_localita = (
        solo_sardegna_definite.groupby("localita_sardegna_prevalente", as_index=False)
        .agg(questionari=(ID_COL, "nunique"))
        .sort_values(by=["questionari", "localita_sardegna_prevalente"], ascending=[False, True], kind="mergesort")
        .head(15)["localita_sardegna_prevalente"]
        .tolist()
    )
    top15_df = solo_sardegna_definite[
        solo_sardegna_definite["localita_sardegna_prevalente"].isin(top15_localita)
        & solo_sardegna_definite["giudizio_score"].notna()
    ].copy()
    outputs["c2_top15_giud_medio"] = (
        top15_df.groupby(["localita_sardegna_prevalente"], as_index=False)
        .agg(
            questionari=(ID_COL, "nunique"),
            componenti=(COMPONENTI_COL, "sum"),
            componenti_pesati=("componenti_pesati", "sum"),
            media_giudizio=("giudizio_score", "mean"),
            somma_pesi=("peso_provincia", "sum"),
            somma_giudizio_pesato=("giudizio_score_pesato", "sum"),
        )
        .assign(media_giudizio_pesata=lambda x: x["somma_giudizio_pesato"] / x["somma_pesi"])
        .drop(columns=["somma_pesi", "somma_giudizio_pesato"])
        .sort_values(
            by=["media_giudizio_pesata", "media_giudizio", "questionari", "localita_sardegna_prevalente"],
            ascending=[False, False, False, True],
            kind="mergesort",
        )
    )
    outputs["c2_top15_giud_medio"] = prettify_columns(outputs["c2_top15_giud_medio"], ["localita_sardegna_prevalente"])

    outputs["c2_top15_giud_x_prov"] = (
        top15_df.groupby(["localita_sardegna_prevalente", PROVINCIA_COL], as_index=False)
        .agg(
            questionari=(ID_COL, "nunique"),
            componenti=(COMPONENTI_COL, "sum"),
            componenti_pesati=("componenti_pesati", "sum"),
            media_giudizio=("giudizio_score", "mean"),
            somma_pesi=("peso_provincia", "sum"),
            somma_giudizio_pesato=("giudizio_score_pesato", "sum"),
        )
        .assign(media_giudizio_pesata=lambda x: x["somma_giudizio_pesato"] / x["somma_pesi"])
        .drop(columns=["somma_pesi", "somma_giudizio_pesato"])
        .sort_values(
            by=["localita_sardegna_prevalente", "media_giudizio_pesata", "media_giudizio", "questionari", PROVINCIA_COL],
            ascending=[True, False, False, False, True],
            kind="mergesort",
        )
    )
    outputs["c2_top15_giud_x_prov"] = prettify_columns(
        outputs["c2_top15_giud_x_prov"],
        ["localita_sardegna_prevalente", PROVINCIA_COL],
    )

    return outputs


REPORT_LAYOUT_SARDI = [
    (
        "sintesi",
        "Analisi qualitative sardi - sintesi",
        [
            ("meta", "Metadati del campione"),
            ("c1_no_vacanza_x_prov", "Motivazioni del non fare vacanza per provincia"),
            ("c2_destinazioni_x_prov", "Destinazioni di vacanza per provincia"),
            ("c2_localita_sardegna", "Localita prevalente in Sardegna per provincia"),
            ("c2_web_x_prov", "Uso del web per destinazione e provincia"),
            ("c2_web_cosa", "Cosa viene prenotato via web"),
        ],
    ),
    (
        "motivazioni_alloggio",
        "Analisi qualitative sardi - motivazioni e alloggio",
        [
            ("c2_alloggio_x_prov", "Tipologia alloggio per destinazione e provincia"),
            ("c2_motiv_princ_x_prov", "Motivazione principale per destinazione e provincia"),
            ("c2_spesa_mot_giud", "Spesa futura associata a motivazione e giudizio"),
            ("c2_motiv1_motiv2", "Motivazione principale e secondaria"),
            ("c2_giudizio_dist", "Distribuzione giudizi per destinazione e provincia"),
        ],
    ),
    (
        "top_localita",
        "Analisi qualitative sardi - top localita",
        [
            ("c2_top15_giud_medio", "Top 15 localita definite per giudizio medio (escluso NON DEFINITO)"),
            ("c2_top15_giud_x_prov", "Top 15 localita definite per giudizio medio e provincia (escluso NON DEFINITO)"),
        ],
    ),
    (
        "audit",
        "Analisi qualitative sardi - audit",
        [
            ("audit_web", "Audit classificazione web"),
        ],
    ),
]


def write_excel(outputs: dict[str, pd.DataFrame], output_file: Path) -> Path:
    saved_path = output_file
    try:
        write_structured_excel(outputs, saved_path, REPORT_LAYOUT_SARDI)
    except PermissionError:
        saved_path = output_file.with_name(f"{output_file.stem}_nuovo{output_file.suffix}")
        write_structured_excel(outputs, saved_path, REPORT_LAYOUT_SARDI)
    return saved_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Costruisce le analisi qualitative del campione sardi.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise FileNotFoundError(f"File non trovato: {args.input}")

    campione_1 = prepare_dataframe(args.input, SHEET_CAMPIONE_1)
    campione_2 = prepare_dataframe(args.input, SHEET_CAMPIONE_2)
    outputs = build_outputs(campione_1, campione_2)
    saved_path = write_excel(outputs, args.output)
    print(f"Analisi qualitative sardi salvate in: {saved_path}")


if __name__ == "__main__":
    main()
