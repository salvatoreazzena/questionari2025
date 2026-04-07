from __future__ import annotations

import argparse
import json
import shutil
import unicodedata
from pathlib import Path
from urllib.request import urlopen

import pandas as pd

COMUNI_JSON_URL = "https://raw.githubusercontent.com/matteocontrini/comuni-json/master/comuni.json"
DEFAULT_INPUT = Path("questionari_fonte.xlsx")
DEFAULT_OUTPUT = Path("questionari_fonte_modificato.xlsx")

ITALIAN_REGIONS_CANONICAL = {
    "ABRUZZO": "ABRUZZO",
    "BASILICATA": "BASILICATA",
    "CALABRIA": "CALABRIA",
    "CAMPANIA": "CAMPANIA",
    "EMILIA ROMAGNA": "EMILIA-ROMAGNA",
    "EMILIA-ROMAGNA": "EMILIA-ROMAGNA",
    "EMILIAROMAGNA": "EMILIA-ROMAGNA",
    "FRIULI VENEZIA GIULIA": "FRIULI-VENEZIA GIULIA",
    "FRIULI-VENEZIA GIULIA": "FRIULI-VENEZIA GIULIA",
    "FRIULIVENEZIA GIULIA": "FRIULI-VENEZIA GIULIA",
    "LAZIO": "LAZIO",
    "LIGURIA": "LIGURIA",
    "LOMBARDIA": "LOMBARDIA",
    "MARCHE": "MARCHE",
    "MOLISE": "MOLISE",
    "PIEMONTE": "PIEMONTE",
    "PUGLIA": "PUGLIA",
    "SARDEGNA": "SARDEGNA",
    "SICILIA": "SICILIA",
    "TOSCANA": "TOSCANA",
    "TRENTINO ALTO ADIGE": "TRENTINO-ALTO ADIGE",
    "TRENTINO-ALTO ADIGE": "TRENTINO-ALTO ADIGE",
    "TRENTINO ALTO ADIGE SUDTIROL": "TRENTINO-ALTO ADIGE",
    "UMBRIA": "UMBRIA",
    "VALLE D AOSTA": "VALLE D'AOSTA",
    "VALLE D AOSTA VALLEE D AOSTE": "VALLE D'AOSTA",
    "VENETO": "VENETO",
}


def normalize_key(value: str) -> str:
    txt = value.strip().upper()
    txt = txt.replace("'", " ").replace("`", " ").replace("’", " ")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = " ".join(txt.split())
    return txt


def is_empty(value: object) -> bool:
    if pd.isna(value):
        return True
    txt = str(value).strip().upper()
    return txt in {"", "ND", "NR", "NA", "N/D", "N.R.", "N.A."}


def load_comuni_mapping() -> tuple[dict[str, str], set[str]]:
    data = json.load(urlopen(COMUNI_JSON_URL, timeout=30))
    by_name: dict[str, set[str]] = {}
    for rec in data:
        comune = normalize_key(rec["nome"])
        regione = rec["regione"]["nome"].strip().upper()
        by_name.setdefault(comune, set()).add(regione)

    unique_map: dict[str, str] = {}
    ambiguous: set[str] = set()
    for comune, regioni in by_name.items():
        if len(regioni) == 1:
            unique_map[comune] = next(iter(regioni))
        else:
            ambiguous.add(comune)

    return unique_map, ambiguous


def transform(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    if "stato_provenienza" not in df.columns:
        raise ValueError("Colonna mancante: stato_provenienza")

    # Supporta sia schema originale (regione_provenienza) sia schema gia trasformato.
    if "comune_provenienza" not in df.columns:
        if "regione_provenienza" not in df.columns:
            raise ValueError("Colonna mancante: regione_provenienza")
        df = df.rename(columns={"regione_provenienza": "comune_provenienza"})

    comuni_to_regione, comuni_ambigui = load_comuni_mapping()

    # Inserire/riordinare la nuova colonna regione_provenienza tra stato e comune.
    if "regione_provenienza" not in df.columns:
        df["regione_provenienza"] = pd.NA
    cols_no_regione = [c for c in df.columns if c != "regione_provenienza"]
    idx_stato = cols_no_regione.index("stato_provenienza")
    cols_no_regione.insert(idx_stato + 1, "regione_provenienza")
    df = df[cols_no_regione]

    cnt_region_direct = 0
    cnt_comune_mapped = 0
    cnt_unresolved = 0
    cnt_non_italia = 0

    for i, row in df.iterrows():
        stato = row["stato_provenienza"]
        comune_raw = row["comune_provenienza"]

        if is_empty(comune_raw):
            continue

        stato_key = normalize_key(str(stato)) if not is_empty(stato) else ""
        if stato_key != "ITALIA":
            cnt_non_italia += 1
            continue

        comune_txt = str(comune_raw).strip()
        key = normalize_key(comune_txt)

        # Se nella vecchia colonna c'e gia una regione, trasferiscila e svuota comune.
        if key in ITALIAN_REGIONS_CANONICAL:
            df.at[i, "regione_provenienza"] = ITALIAN_REGIONS_CANONICAL[key]
            df.at[i, "comune_provenienza"] = pd.NA
            cnt_region_direct += 1
            continue

        # Altrimenti prova mappatura comune -> regione.
        if key in comuni_to_regione and key not in comuni_ambigui:
            df.at[i, "regione_provenienza"] = comuni_to_regione[key]
            cnt_comune_mapped += 1
        else:
            cnt_unresolved += 1

    stats = {
        "regioni_trasferite_direttamente": cnt_region_direct,
        "comuni_mappati_in_regione": cnt_comune_mapped,
        "non_risolti": cnt_unresolved,
        "righe_non_italia_con_comune": cnt_non_italia,
    }
    return df, stats


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Corregge provenienza: rinomina regione_provenienza in comune_provenienza "
            "e crea nuova regione_provenienza valorizzata da regione gia presente o da mapping comune->regione."
        )
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument(
        "--inplace",
        action="store_true",
        help="Sovrascrive il file input creando prima un backup .backup.xlsx",
    )
    args = parser.parse_args()

    if not args.input.exists():
        raise FileNotFoundError(f"File input non trovato: {args.input}")

    xls = pd.ExcelFile(args.input)
    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(args.input, sheet_name=first_sheet)

    out_df, stats = transform(df)

    if args.inplace:
        backup = args.input.with_name(f"{args.input.stem}.backup_pre_provenienza.xlsx")
        shutil.copy2(args.input, backup)
        output_path = args.input
    else:
        output_path = args.output

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name=first_sheet, index=False)

    print(f"File scritto: {output_path}")
    if args.inplace:
        print(f"Backup creato: {backup}")
    print("Statistiche:")
    for k, v in stats.items():
        print(f"- {k}: {v}")


if __name__ == "__main__":
    main()
