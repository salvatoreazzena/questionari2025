from __future__ import annotations

import argparse
from pathlib import Path

from build_analisi_qualitative_campione1 import (
    DEFAULT_INPUT,
    DEFAULT_SHEET_1D,
    prepare_dataframe_1d,
    build_outputs_1d,
    write_structured_excel,
)

DEFAULT_OUTPUT = Path("analisi_campione_1d_motivazione_apprezzamento.xlsx")

REPORT_LAYOUT_1D = [
    (
        "campione_1d",
        "Analisi qualitative - campione 1d",
        [
            ("c1d_meta", "Metadati campione 1d"),
            ("c1d_top_apprezzamenti", "Top 10 apprezzamenti per provenienza e motivazione"),
            ("c1d_dettaglio_completo", "Dettaglio completo apprezzamenti campione 1d"),
        ],
    ),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Wrapper compatibile per l'analisi del campione 1d.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help=f"File input Excel (default: {DEFAULT_INPUT})")
    parser.add_argument("--sheet", default=DEFAULT_SHEET_1D, help=f"Foglio da analizzare (default: {DEFAULT_SHEET_1D})")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help=f"File output Excel (default: {DEFAULT_OUTPUT})")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise FileNotFoundError(f"File non trovato: {args.input}")

    df = prepare_dataframe_1d(args.input, args.sheet)
    outputs = build_outputs_1d(df, args.sheet)
    write_structured_excel(outputs, args.output, REPORT_LAYOUT_1D)

    print(f"Input letto: {args.input} [{args.sheet}]")
    print(f"Output generato: {args.output}")


if __name__ == "__main__":
    main()
