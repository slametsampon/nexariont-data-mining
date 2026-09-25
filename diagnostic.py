import argparse
from pathlib import Path
from typing import Any

from src.config import MiningConfig
from src.parser import ShiftReportParser
from src.workbook_reader import WorkbookReader


class WorksheetDiagnostic:
    """
    Diagnostic runner untuk menguji compatibility
    existing ShiftReportParser terhadap seluruh worksheet.

    Diagnostic ini:
    - membaca seluruh worksheet;
    - mengecualikan worksheet non-visible;
    - menjalankan existing parser pada visible worksheet;
    - menampilkan jumlah record;
    - menampilkan sample record untuk review;
    - tidak menulis output Excel;
    - tidak menggabungkan records;
    - tidak mengubah parser;
    - tidak mengubah source workbook.
    """

    def __init__(
        self,
        config: MiningConfig,
    ):
        self.config = config

        self.parser = ShiftReportParser(
            config
        )

    def run(
        self,
        input_file: Path,
    ) -> None:

        reader = WorkbookReader(
            input_file
        )

        worksheet_info = (
            reader.get_worksheet_info()
        )

        print()
        print("QC MULTI-WORKSHEET DIAGNOSTIC")
        print("=" * 100)
        print(f"Input: {input_file}")
        print("=" * 100)

        total_worksheets = len(
            worksheet_info
        )

        visible_count = 0
        excluded_count = 0
        success_count = 0
        error_count = 0
        total_records = 0

        for info in worksheet_info:

            if not info.is_visible:

                excluded_count += 1

                self._print_result(
                    worksheet_name=info.name,
                    visibility="HIDDEN",
                    status="EXCLUDED",
                    records=0,
                    message="Non-visible worksheet",
                )

                print()

                continue

            visible_count += 1

            try:
                worksheet = (
                    reader.load_worksheet(
                        info.name
                    )
                )

                records = self.parser.parse(
                    worksheet
                )

                record_count = len(
                    records
                )

                success_count += 1
                total_records += record_count

                self._print_result(
                    worksheet_name=info.name,
                    visibility="VISIBLE",
                    status="SUCCESS",
                    records=record_count,
                    message="-",
                )

                self._print_record_samples(
                    records=records,
                    max_samples=5,
                )

            except Exception as exc:

                error_count += 1

                self._print_result(
                    worksheet_name=info.name,
                    visibility="VISIBLE",
                    status="ERROR",
                    records=0,
                    message=(
                        f"{type(exc).__name__}: "
                        f"{exc}"
                    ),
                )

                print()

        print("=" * 100)
        print("SUMMARY")
        print("-" * 100)

        print(
            f"Worksheets discovered : "
            f"{total_worksheets}"
        )

        print(
            f"Visible candidates    : "
            f"{visible_count}"
        )

        print(
            f"Excluded              : "
            f"{excluded_count}"
        )

        print(
            f"Success               : "
            f"{success_count}"
        )

        print(
            f"Error                 : "
            f"{error_count}"
        )

        print(
            f"Candidate records     : "
            f"{total_records}"
        )

        print("=" * 100)

    @staticmethod
    def _print_record_samples(
        records: list[Any],
        max_samples: int = 5,
    ) -> None:
        """
        Menampilkan sample hasil existing parser.

        Tujuannya hanya untuk diagnostic/review.

        Tidak:
        - mengubah SamplingRecord;
        - melakukan QC validation;
        - melakukan data cleansing;
        - melakukan mapping tambahan;
        - menulis data ke file.
        """

        if not records:
            print(
                "    Sample records: none"
            )
            print()
            return

        sample_count = min(
            len(records),
            max_samples,
        )

        print(
            f"    Sample records "
            f"({sample_count} of {len(records)}):"
        )

        for index, record in enumerate(
            records[:max_samples],
            start=1,
        ):

            print(
                f"      [{index}] "
                f"{record!r}"
            )

        print()

    @staticmethod
    def _print_result(
        worksheet_name: str,
        visibility: str,
        status: str,
        records: int,
        message: str,
    ) -> None:

        print(
            f"{worksheet_name:<22} "
            f"{visibility:<9} "
            f"{status:<9} "
            f"{records:>7}  "
            f"{message}"
        )


def parse_arguments():

    parser = argparse.ArgumentParser(
        description=(
            "Diagnostic compatibility parser "
            "untuk seluruh visible worksheet."
        )
    )

    parser.add_argument(
        "--input",
        required=True,
        type=Path,
        help=(
            "Path file input .xls atau .xlsx"
        ),
    )

    return parser.parse_args()


def main() -> None:

    args = parse_arguments()

    input_file = (
        args.input
        .expanduser()
        .resolve()
    )

    if not input_file.exists():
        raise FileNotFoundError(
            f"File tidak ditemukan: "
            f"{input_file}"
        )

    config = MiningConfig()

    diagnostic = WorksheetDiagnostic(
        config
    )

    diagnostic.run(
        input_file
    )


if __name__ == "__main__":
    main()