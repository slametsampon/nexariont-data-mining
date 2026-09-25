import argparse
import sys
from pathlib import Path

from src.config import MiningConfig
from src.service import QCDataMiningService


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
        help="Path file XLSX input",
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Path file XLSX output",
    )

    parser.add_argument(
        "--sheet",
        default="shift-pagi",
        help=(
            "Nama worksheet source "
            "(default: shift-pagi)"
        ),
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

    config = MiningConfig(
        worksheet_name=args.sheet
    )

    service = QCDataMiningService(
        config
    )

    try:

        total_records = service.process_file(
            input_file=input_file,
            output_file=output_file,
        )

        print()
        print("QC DATA MINING COMPLETED")
        print("-" * 50)
        print(
            f"Input     : {input_file}"
        )
        print(
            f"Worksheet : {config.worksheet_name}"
        )
        print(
            f"Output    : {output_file}"
        )
        print(
            f"Records   : {total_records}"
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