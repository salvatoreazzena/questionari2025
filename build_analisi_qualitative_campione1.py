from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd
from build_analisi_preliminare import extract_prevalent_destination

DEFAULT_INPUT = Path("questionari_sottocampioni.xlsx")
DEFAULT_OUTPUT = Path("analisi_qualitative_campione_1.xlsx")
DEFAULT_SHEET = "campione_1"

NULL_VALUE = "NULL"
UNDEFINED_LABEL = "NON DEFINITO"
NULL_TOKENS = {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}

ID_COL = "ID"
AGE_COL = "fascia_età"
STATE_COL = "stato_provenienza"
REGION_COL = "regione_provenienza"
ARRIVAL_COL = "arrivo_sardegna"
LOCATIONS_COL = "località_visitate"
DURATION_COL = "durata_soggiorno"
COMPONENTS_COL = "numero_componenti"
PRIMARY_REASON_COL = "motivazione_principale"
SECONDARY_REASON_COL = "motivazione_secondaria"
LODGING_COL = "tipologia_sistemazione"
WEB_COL = "web"
WEB_DETAIL_COL = "web_specifica"
STRENGTHS_COL = "apprezzamenti"
WEAKNESSES_COL = "miglioramenti"
JUDGMENT_COL = "giudizio"
FUTURE_SPEND_COL = "previsione_spesa"

REQUIRED_COLUMNS = [
    ID_COL,
    AGE_COL,
    STATE_COL,
    REGION_COL,
    ARRIVAL_COL,
    LOCATIONS_COL,
    DURATION_COL,
    COMPONENTS_COL,
    PRIMARY_REASON_COL,
    SECONDARY_REASON_COL,
    LODGING_COL,
    WEB_COL,
    WEB_DETAIL_COL,
    STRENGTHS_COL,
    WEAKNESSES_COL,
    JUDGMENT_COL,
    FUTURE_SPEND_COL,
]

JUDGMENT_SCORE_MAP = {
    "PESSIMO": 1,
    "SUFFICIENTE": 2,
    "BUONO": 3,
    "OTTIMO": 4,
}

TEXT_SPLIT_RE = re.compile(r"[;,/|\n]+|\be\b|\bed\b|\by\b", flags=re.IGNORECASE)

WEB_THEME_PATTERNS: list[tuple[str, str]] = [
    ("VOLO/AEREO", r"\b(volo|voli|aereo|flight|flights)\b"),
    ("NAVE/TRAGHETTO", r"\b(nave|traghetto|traghetti|ferry|ferries)\b"),
    ("ALLOGGIO", r"\b(alloggio|alloggi|hotel|b&b|bnb|appartamento|casa vacanz|vacanze|campeggio|residence|agriturismo|ostello)\b"),
    ("AUTONOLEGGIO/TRASPORTI", r"\b(auto|autonoleggio|noleggio|scooter|moto|bus|autobus|treno|transfer|taxi)\b"),
    ("ESCURSIONI/ATTIVITA", r"\b(escursion|gita|tour|barca|boat|attivit|esperienz|bigliett|ticket|muse|parc|visita)\b"),
    ("RISTORANTI", r"\b(ristorant|pizzeria|cena|pranzo|food|colazione)\b"),
    ("ATTREZZATURA MARE", r"\b(ombrellon|lettin|sdrai|attrezzatura per il mare)\b"),
]

THEME_PATTERNS: list[tuple[str, str]] = [
    ("MARE/SPIAGGE", r"\b(mare|spiagg|acqua|costa|litorale)\b"),
    ("NATURA/PAESAGGIO", r"\b(natura|paesagg|panorama|montagna|parco|verde)\b"),
    ("CLIMA", r"\b(clima|tempo|meteo)\b"),
    ("OSPITALITA/ACCOGLIENZA", r"\b(ospital|accoglien|gentilezza|cordial|persone)\b"),
    ("CIBO/ENOGASTRONOMIA", r"\b(cibo|cucina|ristor|mangiare|enogastr|gastron|food)\b"),
    ("TRANQUILLITA/RELAX", r"\b(tranquill|relax|silenz|pace|calma)\b"),
    ("DIVERTIMENTO/VITA NOTTURNA", r"\b(divert|movida|locali|notturn|serate|mojito)\b"),
    ("CULTURA/BORGHI", r"\b(cultura|muse|storia|borg|artist|architett)\b"),
]

WEAKNESS_PATTERNS: list[tuple[str, str]] = [
    ("VIABILITA/STRADE", r"\b(strad|viabil|cantier|asfalt)\b"),
    ("TRASPORTI/COLLEGAMENTI", r"\b(collegament|trasport|autobus|bus|treno|voli dirett|traghett|aere[io]|taxi|mezzi)\b"),
    ("PARCHEGGI", r"\b(parchegg)\b"),
    ("PREZZI/COSTI", r"\b(prezz|car[oi]|costi|troppo cost)\b"),
    ("SERVIZI DIGITALI/INFORMAZIONI", r"\b(online|sito|app|informazioni|servizi)\b"),
    ("PULIZIA/RIFIUTI", r"\b(pulizi|rifiut|cassonett|sporc)\b"),
    ("SOVRAFFOLLAMENTO", r"\b(affoll|confusion|caos)\b"),
    ("SPIAGGE/ACCESSIBILITA", r"\b(spiagg|access|accessibil|prenotazion)\b"),
]

STOPWORDS = {
    "il", "lo", "la", "i", "gli", "le", "un", "una", "uno", "di", "a", "da", "in", "su", "per", "con",
    "del", "della", "dello", "dei", "degli", "delle", "al", "alla", "allo", "ai", "agli", "alle", "ed",
    "e", "o", "the", "and", "to", "of", "is", "are", "non", "piu", "più", "molto", "molti", "molte",
    "nulla", "niente", "tutto", "tutti", "tutte", "same", "good", "nice", "very",
}


def normalize_text_series(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("").str.strip()
    upper = s.str.upper()
    null_upper = {token.upper() for token in NULL_TOKENS}
    return s.mask(upper.isin(null_upper), NULL_VALUE)


def normalize_upper_series(series: pd.Series) -> pd.Series:
    return normalize_text_series(series).str.upper()


def validate_columns(df: pd.DataFrame) -> None:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Colonne mancanti nel dataset: {missing}")


def strip_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize_free_text(value: object) -> str:
    if pd.isna(value):
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    upper = raw.upper()
    if upper in {token.upper() for token in NULL_TOKENS}:
        return ""
    text = strip_accents(raw.lower())
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def prettify_value(value: object) -> str:
    if pd.isna(value):
        return UNDEFINED_LABEL
    text = str(value).strip()
    if not text:
        return UNDEFINED_LABEL
    if text.upper() == NULL_VALUE:
        return UNDEFINED_LABEL
    return text


def build_macro_provenienza(state_series: pd.Series) -> pd.Series:
    upper = normalize_upper_series(state_series)
    out = pd.Series(UNDEFINED_LABEL, index=upper.index, dtype="string")
    out = out.mask(upper.eq("ITALIA"), "ITALIANI")
    out = out.mask((upper != "") & (upper != NULL_VALUE) & (~upper.eq("ITALIA")), "STRANIERI")
    return out


def build_provenienza_label(df: pd.DataFrame) -> pd.Series:
    macro = df["macro_provenienza"]
    state = normalize_upper_series(df[STATE_COL])
    region = normalize_upper_series(df[REGION_COL])

    out = pd.Series(UNDEFINED_LABEL, index=df.index, dtype="string")
    out = out.mask(macro.eq("ITALIANI") & region.ne(NULL_VALUE), region)
    out = out.mask(macro.eq("STRANIERI") & state.ne(NULL_VALUE), state)
    return out


def extract_destination_days(value: object) -> list[tuple[str, int]]:
    if pd.isna(value):
        return []

    raw_text = str(value).strip()
    if not raw_text or raw_text.upper() == NULL_VALUE:
        return []

    if raw_text.startswith("(") and raw_text.endswith(")"):
        raw_text = raw_text[1:-1].strip()

    pair_pattern = re.compile(
        r"\s*([^\d,;:][^,;:]*?)\s*(?:,+\s*|:\s*|\s+)([0-9]+(?:[.,][0-9]+)?)(?![A-Za-z])\s*(?=;|,|:|$)"
    )
    matches = list(pair_pattern.finditer(raw_text))

    results: list[tuple[str, int]] = []
    for match in matches:
        destination_raw = match.group(1).strip()
        days_raw = match.group(2).strip()
        destination_clean = re.sub(r"\s*\([^)]*\)\s*", " ", destination_raw)
        destination_clean = re.sub(r"^(?:\s*\d+\s*:\s*)+", "", destination_clean)
        destination_clean = re.sub(r"\s+", " ", destination_clean).strip(" ,;")
        if not destination_clean:
            continue
        try:
            days_float = float(days_raw.replace(",", "."))
        except ValueError:
            continue
        if not days_float.is_integer():
            continue
        days_int = int(days_float)
        if days_int < 0:
            continue
        results.append((destination_clean.upper(), days_int))

    if results:
        return results

    if ";" not in raw_text and not re.search(r"\d", raw_text):
        single_destination = re.sub(r"\s*\([^)]*\)\s*", " ", raw_text)
        single_destination = re.sub(r"^(?:\s*\d+\s*:\s*)+", "", single_destination)
        single_destination = re.sub(r"\s+", " ", single_destination).strip(" ,;")
        if single_destination:
            return [(single_destination.upper(), 1)]

    return []


def build_destination_metrics(locations_series: pd.Series) -> pd.DataFrame:
    prevalent: list[str] = []
    unique_positive_counts: list[int] = []
    travel_type: list[str] = []

    for value in locations_series.tolist():
        pairs = extract_destination_days(value)
        if not pairs:
            prevalent.append(extract_prevalent_destination(value))
            unique_positive_counts.append(0)
            travel_type.append(UNDEFINED_LABEL)
            continue

        totals: dict[str, int] = {}
        first_seen: dict[str, int] = {}
        positive_destinations: list[str] = []
        for idx, (dest, days) in enumerate(pairs):
            if dest not in first_seen:
                first_seen[dest] = idx
            totals[dest] = totals.get(dest, 0) + days
            if days > 0:
                positive_destinations.append(dest)

        if not totals:
            prevalent.append(extract_prevalent_destination(value))
            unique_positive_counts.append(0)
            travel_type.append(UNDEFINED_LABEL)
            continue

        prev = extract_prevalent_destination(value)
        if prev == NULL_VALUE:
            prev = UNDEFINED_LABEL
        if prev not in totals:
            prev = min(totals.keys(), key=lambda k: (-totals[k], first_seen[k], k))
        positive_unique = len(dict.fromkeys(positive_destinations))

        prevalent.append(prev)
        unique_positive_counts.append(positive_unique)
        if positive_unique <= 0:
            travel_type.append(UNDEFINED_LABEL)
        elif positive_unique == 1:
            travel_type.append("STANZIALE")
        else:
            travel_type.append("ITINERANTE")

    return pd.DataFrame(
        {
            "destinazione_prevalente": pd.Series(prevalent, dtype="string"),
            "numero_destinazioni_positive": pd.Series(unique_positive_counts, dtype="int64"),
            "tipo_turismo": pd.Series(travel_type, dtype="string"),
        }
    )


def normalize_judgment(series: pd.Series) -> pd.Series:
    return normalize_upper_series(series)


def normalize_yes_no(series: pd.Series) -> pd.Series:
    upper = normalize_upper_series(series)
    out = pd.Series(UNDEFINED_LABEL, index=upper.index, dtype="string")
    out = out.mask(upper.isin({"SI", "SÌ"}), "SI")
    out = out.mask(upper.eq("NO"), "NO")
    return out


def classify_text_themes(text: str, patterns: list[tuple[str, str]], default_label: str) -> list[str]:
    if not text:
        return []
    themes = [label for label, pattern in patterns if re.search(pattern, text, flags=re.IGNORECASE)]
    if themes:
        return themes
    return [default_label]


def split_text_fragments(text: str) -> list[str]:
    if not text:
        return []
    fragments = []
    for part in TEXT_SPLIT_RE.split(text):
        cleaned = re.sub(r"\s+", " ", part).strip(" .,:;-")
        if not cleaned or cleaned in STOPWORDS:
            continue
        if len(cleaned) <= 2:
            continue
        fragments.append(cleaned)
    return fragments


def build_text_theme_table(
    df: pd.DataFrame,
    *,
    text_col: str,
    group_cols: list[str],
    patterns: list[tuple[str, str]],
    default_label: str,
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for _, row in df.iterrows():
        text = normalize_free_text(row[text_col])
        if not text:
            continue
        themes = classify_text_themes(text, patterns, default_label)
        for theme in themes:
            base = {col: row[col] for col in group_cols}
            base.update({"tema": theme, "ID": row[ID_COL]})
            rows.append(base)

    if not rows:
        return pd.DataFrame(columns=[*group_cols, "tema", "questionari"])

    out = pd.DataFrame(rows)
    return (
        out.groupby([*group_cols, "tema"], dropna=False, as_index=False)
        .agg(questionari=("ID", "nunique"))
        .sort_values(by=[*group_cols, "questionari", "tema"], ascending=[True] * len(group_cols) + [False, True], kind="mergesort")
    )


def build_text_fragment_table(
    df: pd.DataFrame,
    *,
    text_col: str,
    group_cols: list[str],
    top_n: int,
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for _, row in df.iterrows():
        text = normalize_free_text(row[text_col])
        if not text:
            continue
        fragments = split_text_fragments(text)
        for fragment in fragments:
            base = {col: row[col] for col in group_cols}
            base.update({"frammento": fragment, "ID": row[ID_COL]})
            rows.append(base)

    if not rows:
        return pd.DataFrame(columns=[*group_cols, "frammento", "questionari"])

    out = (
        pd.DataFrame(rows)
        .groupby([*group_cols, "frammento"], dropna=False, as_index=False)
        .agg(questionari=("ID", "nunique"))
    )
    out = out.sort_values(
        by=[*group_cols, "questionari", "frammento"],
        ascending=[True] * len(group_cols) + [False, True],
        kind="mergesort",
    )
    return out.groupby(group_cols, dropna=False, group_keys=False).head(top_n).reset_index(drop=True)


def add_share_within_group(df: pd.DataFrame, group_cols: list[str], value_col: str = "questionari") -> pd.DataFrame:
    if df.empty:
        return df.copy()
    out = df.copy()
    totals = out.groupby(group_cols, dropna=False)[value_col].transform("sum")
    out["pct_su_gruppo"] = (out[value_col] / totals * 100.0).round(2)
    return out


def build_distribution(
    df: pd.DataFrame,
    *,
    group_cols: list[str],
    value_col: str,
    include_components: bool = True,
) -> pd.DataFrame:
    grouped = (
        df.groupby([*group_cols, value_col], dropna=False, as_index=False)
        .agg(
            questionari=(ID_COL, "nunique"),
            componenti=(COMPONENTS_COL, "sum"),
        )
    )
    if not include_components:
        grouped = grouped.drop(columns=["componenti"])
    return grouped.sort_values(
        by=[*group_cols, "questionari", value_col],
        ascending=[True] * len(group_cols) + [False, True],
        kind="mergesort",
    )


def build_top_n_by_group(
    df: pd.DataFrame,
    *,
    group_cols: list[str],
    value_col: str,
    top_n: int,
) -> pd.DataFrame:
    out = build_distribution(df, group_cols=group_cols, value_col=value_col, include_components=True)
    out = out.groupby(group_cols, dropna=False, group_keys=False).head(top_n).reset_index(drop=True)
    return add_share_within_group(out, group_cols)


def prepare_dataframe(input_file: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(input_file, sheet_name=sheet_name)
    validate_columns(raw)

    df = raw.copy()
    for col in REQUIRED_COLUMNS:
        if col == COMPONENTS_COL:
            continue
        df[col] = normalize_text_series(df[col])

    df[COMPONENTS_COL] = pd.to_numeric(df[COMPONENTS_COL], errors="coerce").fillna(0)
    df[DURATION_COL] = pd.to_numeric(df[DURATION_COL], errors="coerce")
    df["macro_provenienza"] = build_macro_provenienza(df[STATE_COL])
    df["provenienza_associata"] = build_provenienza_label(df)
    df["web_flag"] = normalize_yes_no(df[WEB_COL])
    df["giudizio_norm"] = normalize_judgment(df[JUDGMENT_COL])
    df["giudizio_score"] = df["giudizio_norm"].map(JUDGMENT_SCORE_MAP)

    destination_metrics = build_destination_metrics(df[LOCATIONS_COL])
    df = pd.concat([df, destination_metrics], axis=1)
    df["destinazione_prevalente"] = (
        df["destinazione_prevalente"]
        .astype("string")
        .fillna(UNDEFINED_LABEL)
        .replace({NULL_VALUE: UNDEFINED_LABEL, "": UNDEFINED_LABEL})
    )

    return df


def build_outputs(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    outputs: dict[str, pd.DataFrame] = {}

    meta_rows = [
        {"chiave": "sheet_campione", "valore": DEFAULT_SHEET},
        {"chiave": "questionari_totali", "valore": int(df[ID_COL].nunique())},
        {"chiave": "componenti_totali", "valore": int(df[COMPONENTS_COL].sum())},
        {"chiave": "italiani_questionari", "valore": int(df["macro_provenienza"].eq("ITALIANI").sum())},
        {"chiave": "stranieri_questionari", "valore": int(df["macro_provenienza"].eq("STRANIERI").sum())},
    ]
    outputs["meta"] = pd.DataFrame(meta_rows)

    web_usage = build_distribution(df, group_cols=["macro_provenienza"], value_col="web_flag", include_components=True)
    outputs["web_prenotazione"] = add_share_within_group(web_usage, ["macro_provenienza"])

    web_yes = df[df["web_flag"] == "SI"].copy()
    web_theme_rows: list[dict[str, object]] = []
    for _, row in web_yes.iterrows():
        text = normalize_free_text(row[WEB_DETAIL_COL])
        themes = classify_text_themes(text, WEB_THEME_PATTERNS, "ALTRO/NON CLASSIFICATO")
        for theme in themes:
            web_theme_rows.append(
                {
                    "macro_provenienza": row["macro_provenienza"],
                    "cosa_prenotata": theme,
                    "ID": row[ID_COL],
                }
            )
    web_theme_df = pd.DataFrame(web_theme_rows)
    if web_theme_df.empty:
        outputs["web_cosa_prenotata"] = pd.DataFrame(columns=["macro_provenienza", "cosa_prenotata", "questionari", "pct_su_gruppo"])
    else:
        web_theme_df = (
            web_theme_df.groupby(["macro_provenienza", "cosa_prenotata"], as_index=False)
            .agg(questionari=("ID", "nunique"))
            .sort_values(by=["macro_provenienza", "questionari", "cosa_prenotata"], ascending=[True, False, True], kind="mergesort")
        )
        outputs["web_cosa_prenotata"] = add_share_within_group(web_theme_df, ["macro_provenienza"])

    outputs["web_dettagli_top20"] = build_text_fragment_table(
        web_yes,
        text_col=WEB_DETAIL_COL,
        group_cols=["macro_provenienza"],
        top_n=20,
    )

    travel_type = build_distribution(df, group_cols=["macro_provenienza"], value_col="tipo_turismo", include_components=True)
    outputs["stanziale_itinerante"] = add_share_within_group(travel_type, ["macro_provenienza"])

    arrivals = build_distribution(
        df,
        group_cols=["macro_provenienza", "destinazione_prevalente"],
        value_col=ARRIVAL_COL,
        include_components=True,
    )
    outputs["arrivi_x_destinazione"] = add_share_within_group(arrivals, ["macro_provenienza", "destinazione_prevalente"])

    lodging = build_distribution(df, group_cols=["macro_provenienza"], value_col=LODGING_COL, include_components=True)
    outputs["alloggio_prevalente"] = add_share_within_group(lodging, ["macro_provenienza"])

    reason_lodging = build_distribution(
        df,
        group_cols=["macro_provenienza", PRIMARY_REASON_COL],
        value_col=LODGING_COL,
        include_components=True,
    )
    outputs["motivazione_x_alloggio"] = add_share_within_group(reason_lodging, ["macro_provenienza", PRIMARY_REASON_COL])

    outputs["dest_top10_provenienze"] = build_top_n_by_group(
        df,
        group_cols=["macro_provenienza", "destinazione_prevalente"],
        value_col="provenienza_associata",
        top_n=10,
    )
    outputs["dest_top_motivazioni"] = build_top_n_by_group(
        df,
        group_cols=["macro_provenienza", "destinazione_prevalente"],
        value_col=PRIMARY_REASON_COL,
        top_n=10,
    )

    spend_assoc = (
        df.groupby(
            [FUTURE_SPEND_COL, "macro_provenienza", "provenienza_associata", PRIMARY_REASON_COL, "giudizio_norm"],
            dropna=False,
            as_index=False,
        )
        .agg(questionari=(ID_COL, "nunique"), componenti=(COMPONENTS_COL, "sum"))
        .sort_values(
            by=[FUTURE_SPEND_COL, "macro_provenienza", "questionari", "provenienza_associata"],
            ascending=[True, True, False, True],
            kind="mergesort",
        )
    )
    outputs["spesa_futura_assoc"] = spend_assoc

    top20_destinations = (
        df.groupby("destinazione_prevalente", as_index=False)
        .agg(questionari=(ID_COL, "nunique"))
        .sort_values(by=["questionari", "destinazione_prevalente"], ascending=[False, True], kind="mergesort")
        .head(20)["destinazione_prevalente"]
        .tolist()
    )
    top20_df = df[df["destinazione_prevalente"].isin(top20_destinations)].copy()
    top20_reason = build_distribution(
        top20_df,
        group_cols=["destinazione_prevalente"],
        value_col=PRIMARY_REASON_COL,
        include_components=True,
    )
    outputs["top20_dest_x_motivaz"] = add_share_within_group(top20_reason, ["destinazione_prevalente"])

    reason_secondary = build_distribution(
        df[df[SECONDARY_REASON_COL] != NULL_VALUE],
        group_cols=[PRIMARY_REASON_COL],
        value_col=SECONDARY_REASON_COL,
        include_components=True,
    )
    outputs["motivaz_prim_second"] = add_share_within_group(reason_secondary, [PRIMARY_REASON_COL])

    judgment_dist = build_distribution(
        df,
        group_cols=["macro_provenienza", AGE_COL],
        value_col="giudizio_norm",
        include_components=True,
    )
    outputs["giudizio_distrib"] = add_share_within_group(judgment_dist, ["macro_provenienza", AGE_COL])

    top15_destinations = (
        df.groupby("destinazione_prevalente", as_index=False)
        .agg(questionari=(ID_COL, "nunique"))
        .sort_values(by=["questionari", "destinazione_prevalente"], ascending=[False, True], kind="mergesort")
        .head(15)["destinazione_prevalente"]
        .tolist()
    )
    top15_df = df[df["destinazione_prevalente"].isin(top15_destinations) & df["giudizio_score"].notna()].copy()
    judgment_mean = (
        top15_df.groupby(["destinazione_prevalente", "macro_provenienza"], as_index=False)
        .agg(
            questionari=(ID_COL, "nunique"),
            componenti=(COMPONENTS_COL, "sum"),
            media_giudizio=("giudizio_score", "mean"),
        )
    )
    overall_mean = (
        top15_df.groupby("destinazione_prevalente", as_index=False)
        .agg(
            questionari=(ID_COL, "nunique"),
            componenti=(COMPONENTS_COL, "sum"),
            media_giudizio=("giudizio_score", "mean"),
        )
    )
    overall_mean.insert(1, "macro_provenienza", "TOTALE")
    outputs["giudizio_top15_dest"] = (
        pd.concat([overall_mean, judgment_mean], ignore_index=True)
        .sort_values(by=["media_giudizio", "questionari", "destinazione_prevalente"], ascending=[False, False, True], kind="mergesort")
    )

    outputs["punti_forza_temi"] = add_share_within_group(
        build_text_theme_table(
            df,
            text_col=STRENGTHS_COL,
            group_cols=["macro_provenienza"],
            patterns=THEME_PATTERNS,
            default_label="ALTRO/NON CLASSIFICATO",
        ),
        ["macro_provenienza"],
    )
    outputs["punti_forza_top20"] = build_text_fragment_table(
        df,
        text_col=STRENGTHS_COL,
        group_cols=["macro_provenienza"],
        top_n=20,
    )

    dissent_df = df[df["giudizio_norm"].isin(["SUFFICIENTE", "PESSIMO"])].copy()
    outputs["dissenso_temi"] = add_share_within_group(
        build_text_theme_table(
            dissent_df,
            text_col=WEAKNESSES_COL,
            group_cols=["macro_provenienza", "giudizio_norm"],
            patterns=WEAKNESS_PATTERNS,
            default_label="ALTRO/NON CLASSIFICATO",
        ),
        ["macro_provenienza", "giudizio_norm"],
    )
    outputs["dissenso_top20"] = build_text_fragment_table(
        dissent_df,
        text_col=WEAKNESSES_COL,
        group_cols=["macro_provenienza", "giudizio_norm"],
        top_n=20,
    )
    outputs["dissenso_dettaglio"] = dissent_df[
        [
            ID_COL,
            "macro_provenienza",
            AGE_COL,
            "provenienza_associata",
            "destinazione_prevalente",
            PRIMARY_REASON_COL,
            "giudizio_norm",
            WEAKNESSES_COL,
        ]
    ].sort_values(by=["macro_provenienza", "giudizio_norm", ID_COL], kind="mergesort")

    return outputs


def write_excel(outputs: dict[str, pd.DataFrame], output_file: Path) -> None:
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, df in outputs.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Costruisce le analisi qualitative del Campione 1.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help=f"File input Excel (default: {DEFAULT_INPUT})")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"Foglio da analizzare (default: {DEFAULT_SHEET})")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help=f"File output Excel (default: {DEFAULT_OUTPUT})")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise FileNotFoundError(f"File non trovato: {args.input}")

    df = prepare_dataframe(args.input, args.sheet)
    outputs = build_outputs(df)
    write_excel(outputs, args.output)
    print(f"Analisi qualitative Campione 1 salvate in: {args.output}")


if __name__ == "__main__":
    main()
