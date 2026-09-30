"""CLI entry point untuk pemrosesan satu workbook QC NEXARIONT.

Module ini menangani parsing command-line argument, validasi pasangan master/domain,
pemanggilan QCWorkbookProcessor, serta presentation hasil dan exit status.

Actual workbook processing tidak diimplementasikan di module ini. Pemrosesan
didelegasikan ke QCWorkbookProcessor agar CLI tetap thin dan terpisah dari
application/core logic.

CLI:
    --input     Workbook .xls/.xlsx sumber.
    --output    Workbook .xlsx hasil.
    --master    Optional Sampling Point Master.
    --domain    Optional domain master; wajib diberikan bersama --master.
"""

import argparse
import sys
from pathlib import Path

from src.workbook_processor import QCWorkbookProcessor

def build_parser() -> argparse.ArgumentParser:
    """Membuat command-line parser untuk single-workbook QC mining.

    Parser menyediakan input/output wajib serta pasangan master/domain opsional.
    Method ini hanya mendefinisikan CLI contract dan tidak melakukan file access
    maupun workbook processing.

    Returns:
        argparse.ArgumentParser: Parser dengan argument --input, --output,
        --master, dan --domain.
    """

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
    """Menjalankan single-workbook QC mining dari command line.

    Argument CLI dikonversi menjadi absolute Path. --master dan --domain harus
    diberikan bersama-sama. Pemrosesan workbook kemudian didelegasikan ke
    QCWorkbookProcessor dan hasil execution ditampilkan ke terminal.

    Returns:
        int: 0 jika processing selesai tanpa unhandled exception; 1 jika
        QCWorkbookProcessor atau dependency processing menghasilkan exception.

    Raises:
        SystemExit: Jika hanya salah satu dari --master atau --domain diberikan.
            argparse juga dapat menghasilkan SystemExit untuk invalid CLI usage.
    """
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