import argparse
import sys
from pathlib import Path

from src.config import MiningConfig
from src.service import QCDataMiningService
from src.exporter import ExcelExporter
from src.sampling_point_master import SamplingPointMaster


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

    config = MiningConfig()

    if bool(args.master) != bool(args.domain):
        raise SystemExit("--master dan --domain harus diberikan bersama-sama")

    sampling_point_master = None
    if args.master:
        sampling_point_master = SamplingPointMaster.load(
            Path(args.master).resolve(),
            worksheet_name=config.sampling_point_master_sheet,
        )

    service = QCDataMiningService(
        config,
        sampling_point_master=sampling_point_master,
        domain=args.domain,
    )

    exporter = ExcelExporter(
        config
    )

    try:

        result = service.process_all_worksheets(
            input_file=input_file,
        )

        exporter.export_multi_worksheet(
            records=result.records,
            worksheet_results=result.worksheet_results,
            output_file=output_file,
            source_file=input_file,
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