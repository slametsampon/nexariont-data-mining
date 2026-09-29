from __future__ import annotations

import argparse
import re
from pathlib import Path

from src.monthly_mining_orchestrator import (
    MASTER_SOURCE_SHEET,
    MonthlyMiningOrchestrator,
    NoWeekFoldersFoundError,
)


MASTER_BASENAME = "Master-Data.xlsx"


def safe_name(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", value)
    cleaned = re.sub(r"\s+", "_", cleaned.strip())
    return cleaned or "MONTH"


def build_parser() -> argparse.ArgumentParser:
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
