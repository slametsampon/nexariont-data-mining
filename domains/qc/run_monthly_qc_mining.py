from __future__ import annotations

import argparse
import re
from pathlib import Path

from src.monthly_mining_orchestrator import (
    MASTER_SOURCE_SHEET,
    MonthlyMiningOrchestrator,
    NoWeekFoldersFoundError,
)

"""CLI entry point untuk monthly QC Data Mining NEXARIONT.

Module ini menangani command-line interface untuk pemrosesan satu monthly source
root menjadi satu consolidated QC workbook.

Responsibility module dibatasi pada:
- parsing CLI;
- validasi input/master/output path;
- penentuan default output filename;
- pemanggilan MonthlyMiningOrchestrator;
- presentation progress dan final summary;
- process exit behavior.

Week/domain/workbook discovery, workbook processing, dan monthly consolidation
ditangani oleh application components di bawah MonthlyMiningOrchestrator.
"""

MASTER_BASENAME = "Master-Data.xlsx"


def safe_name(value: str) -> str:
    """Mengubah text menjadi nama file yang aman untuk output monthly workbook.

    Karakter yang tidak valid pada nama file diganti dengan underscore dan
    whitespace berurutan dinormalisasi menjadi satu underscore.

    Args:
        value (str): Text sumber, biasanya nama monthly source folder.

    Returns:
        str: Nama yang telah dinormalisasi. Mengembalikan "MONTH" jika hasil
        normalisasi kosong.
    """
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", value)
    cleaned = re.sub(r"\s+", "_", cleaned.strip())
    return cleaned or "MONTH"


def build_parser() -> argparse.ArgumentParser:
    """Membuat command-line parser untuk monthly QC mining.

    CLI mempertahankan source/output separation sehingga monthly source dapat
    berasal dari local maupun UNC/server path dan consolidated output dapat
    disimpan pada lokasi berbeda.

    Returns:
        argparse.ArgumentParser: Parser dengan argument --input-root,
        --output-dir, --master, --output-name, dan --overwrite.
    """
    parser = argparse.ArgumentParser(
        description=(
            "NEXARIONT QC monthly runner. "
            "Input monthly source folder and output folder can be different locations "
            "(e.g. input on server, output on local PC)."
        )
    )
    parser.add_argument(
        "--input-root",
        required=True,
        type=Path,
        help=(
            "Monthly source folder. May be local path or UNC/server path. "
            "Example: \\\\SERVER\\QC\\8-AGUSTUS 2026"
        ),
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help=(
            "Destination folder for the single consolidated monthly XLSX. "
            "May be different from input location."
        ),
    )
    parser.add_argument(
        "--master",
        type=Path,
        default=None,
        help=(
            "Optional canonical QC master path. "
            "Default: <repository>\\doc\\Master-Data.xlsx. "
            "Sampling Point data is read from worksheet 'Sampling-Point'."
        ),
    )
    parser.add_argument(
        "--output-name",
        default=None,
        help=(
            "Optional output filename. "
            "Default: QC_Data_Mining_<monthly-folder-name>.xlsx"
        ),
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow replacing an existing final monthly output file.",
    )
    return parser


def main() -> int:
    """Menjalankan monthly QC mining dari command line.

    Method memvalidasi monthly source root dan canonical master, menentukan
    final output filename, melindungi existing output kecuali --overwrite
    digunakan, lalu mendelegasikan execution ke MonthlyMiningOrchestrator.

    Progress dari orchestrator ditampilkan melalui callback print. Setelah
    processing selesai, method menampilkan execution summary termasuk workbook,
    candidate record, worksheet status, execution failure, output-read error,
    serta discovery condition.

    Returns:
        int: 0 jika monthly orchestration selesai dan final summary dapat
        ditampilkan.

    Raises:
        SystemExit: Jika input root tidak tersedia, master tidak tersedia,
            final output sudah ada tanpa --overwrite, tidak ditemukan folder
            minggu, atau argparse mendeteksi invalid CLI usage.
    """
    repo = Path(__file__).resolve().parent
    args = build_parser().parse_args()

    input_root = args.input_root.expanduser()
    output_dir = args.output_dir.expanduser()

    master = (
        args.master.expanduser()
        if args.master is not None
        else repo / "doc" / MASTER_BASENAME
    )

    if not input_root.is_dir():
        raise SystemExit(f"INPUT_ROOT_NOT_FOUND: {input_root}")

    if not master.is_file():
        raise SystemExit(
            "MASTER_NOT_FOUND:\n"
            f"  {master}\n\n"
            "Expected default location:\n"
            f"  {repo / 'doc' / MASTER_BASENAME}\n"
            "Or specify another file with --master."
        )

    output_dir.mkdir(parents=True, exist_ok=True)

    output_name = (
        args.output_name
        or f"QC_Data_Mining_{safe_name(input_root.name)}.xlsx"
    )
    if not output_name.casefold().endswith(".xlsx"):
        output_name += ".xlsx"

    final_output = output_dir / output_name

    if final_output.exists() and not args.overwrite:
        raise SystemExit(
            f"OUTPUT_EXISTS: {final_output}\n"
            "Use --overwrite only if replacement is intentional."
        )

    orchestrator = MonthlyMiningOrchestrator()

    try:
        result = orchestrator.run(
            input_root=input_root,
            output_dir=output_dir,
            master=master,
            final_output=final_output,
            progress=print,
        )
    except NoWeekFoldersFoundError as exc:
        raise SystemExit(str(exc)) from None

    totals = result.totals

    print("\n" + "=" * 72)
    print("NEXARIONT QC MONTHLY MINING COMPLETED")
    print("=" * 72)
    print(f"Input root            : {result.input_root}")
    print(f"Output directory      : {result.output_dir}")
    print(f"Master                : {result.master}")
    print(f"Master worksheet      : {MASTER_SOURCE_SHEET}")
    print(f"Month folder          : {result.month}")
    print(f"Weeks detected        : {result.weeks_detected}")
    print(f"Workbooks attempted   : {totals['workbooks_attempted']}")
    print(f"Candidate records     : {totals['candidate_records']}")
    print(f"Worksheets            : {totals['worksheets']}")
    print(f"Success               : {totals['success']}")
    print(f"Excluded              : {totals['excluded']}")
    print(f"Error                 : {totals['error']}")
    print(f"Execution failures    : {totals['exec_failed']}")
    print(f"Output read errors    : {totals['output_read_error']}")
    print(f"Missing domain folder : {totals['missing_domain_folder']}")
    print(f"Empty domain folder   : {totals['no_workbooks']}")
    print(f"Output file           : {result.final_output}")

    if (
        totals["error"]
        or totals["exec_failed"]
        or totals["output_read_error"]
        or totals["missing_domain_folder"]
        or totals["no_workbooks"]
    ):
        print("\nATTENTION: lihat sheet Run_Summary dan Processing_Log.")
    else:
        print("\nNo processing exception recorded.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
