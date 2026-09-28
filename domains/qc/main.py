import argparse
import sys
from pathlib import Path

from src.workbook_processor import QCWorkbookProcessor


def build_parser() -> argparse.ArgumentParser:

    parser = argparse.ArgumentParser(
        description=(
            "QC Shift Report Data Mining - "
            "normalize Sampling Point data"
        )
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Path file XLS/XLSX input",
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Path file XLSX output",
    )

    parser.add_argument(
        "--master",
        help="Path NEXARIONT Sampling Point Master XLSX",
    )

    parser.add_argument(
        "--domain",
        help="Domain master SSP, mis. NPG, Octanol, Syngas, Utility, Utility NPG, wwt",
    )

    return parser


def main() -> int:

    args = build_parser().parse_args()

    input_file = Path(
        args.input
    ).resolve()

    output_file = Path(
        args.output
    ).resolve()

    if bool(args.master) != bool(args.domain):
        raise SystemExit(
            "--master dan --domain harus diberikan bersama-sama"
        )

    master_file = (
        Path(args.master).resolve()
        if args.master
        else None
    )

    processor = QCWorkbookProcessor()

    try:

        result = processor.process(
            input_file=input_file,
            output_file=output_file,
            master_file=master_file,
            domain=args.domain,
        )

        print()
        print("QC DATA MINING COMPLETED")
        print("-" * 50)
        print(
            f"Input     : {input_file}"
        )
        print(
            f"Output    : {output_file}"
        )
        print(
            f"Records   : {result.total_records}"
        )
        print(
            f"Worksheets: {len(result.worksheet_results)}"
        )
        print(
            f"Success   : {result.success_count}"
        )
        print(
            f"Excluded  : {result.excluded_count}"
        )
        print(
            f"Error     : {result.error_count}"
        )
        print("-" * 50)

        return 0

    except Exception as exc:

        print(
            f"ERROR: {exc}",
            file=sys.stderr,
        )

        return 1


if __name__ == "__main__":
    raise SystemExit(main())