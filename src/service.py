from dataclasses import dataclass, field
from pathlib import Path

from .config import MiningConfig
from .exporter import ExcelExporter
from .models import SourcedSamplingRecord
from .parser import ShiftReportParser
from .workbook_reader import WorkbookReader
from .sampling_point_master import SamplingPointMaster


@dataclass(frozen=True)
class WorksheetProcessingResult:
    """
    Hasil processing untuk satu worksheet.

    status:
        SUCCESS  -> parser selesai tanpa exception.
        ERROR    -> parser menghasilkan exception.
        EXCLUDED -> worksheet tidak diproses.

    record_count adalah jumlah candidate records
    yang dihasilkan existing parser.

    SUCCESS tidak berarti QC validation/approval.
    """

    worksheet_name: str
    visibility: str
    status: str
    record_count: int
    message: str = ""


@dataclass
class MultiWorksheetProcessingResult:
    """
    Aggregate result dari multi-worksheet processing.

    records:
        Candidate records dengan source worksheet
        yang dipertahankan melalui SourcedSamplingRecord.

    worksheet_results:
        Processing result untuk setiap worksheet,
        termasuk worksheet yang excluded.
    """

    records: list[SourcedSamplingRecord] = field(
        default_factory=list
    )

    worksheet_results: list[
        WorksheetProcessingResult
    ] = field(
        default_factory=list
    )

    @property
    def total_records(self) -> int:
        return len(self.records)

    @property
    def success_count(self) -> int:
        return sum(
            1
            for result in self.worksheet_results
            if result.status == "SUCCESS"
        )

    @property
    def error_count(self) -> int:
        return sum(
            1
            for result in self.worksheet_results
            if result.status == "ERROR"
        )

    @property
    def excluded_count(self) -> int:
        return sum(
            1
            for result in self.worksheet_results
            if result.status == "EXCLUDED"
        )


class QCDataMiningService:
    """
    Application service / orchestration layer.

    Existing single-worksheet process_file() dipertahankan.

    Multi-worksheet capability tersedia melalui
    process_all_worksheets() dan belum menjadi
    production export path pada CP4.2.
    """

    def __init__(
        self,
        config: MiningConfig,
        sampling_point_master: SamplingPointMaster | None = None,
        domain: str | None = None,
    ):
        self.config = config

        self.parser = ShiftReportParser(
            config,
            sampling_point_master=sampling_point_master,
            domain=domain,
        )

        self.exporter = ExcelExporter(
            config
        )

    def process_file(
        self,
        input_file: Path,
        output_file: Path,
    ) -> int:
        """
        Existing single-worksheet production path.

        Dipertahankan pada CP4.2 untuk regression safety.
        """

        reader = WorkbookReader(
            input_file
        )

        worksheet = reader.load_worksheet(
            self.config.worksheet_name
        )

        records = self.parser.parse(
            worksheet
        )

        self.exporter.export(
            records=records,
            output_file=output_file,
            source_file=input_file,
        )

        return len(records)

    def process_all_worksheets(
        self,
        input_file: Path,
    ) -> MultiWorksheetProcessingResult:
        """
        Memproses seluruh worksheet berdasarkan
        visibility metadata dan configured worksheet
        exclusion.

        Non-visible:
            tidak diparse dan dicatat EXCLUDED.

        Visible tetapi termasuk configured exclusion:
            tidak diparse dan dicatat EXCLUDED.

        Visible dan tidak termasuk configured exclusion:
            diproses menggunakan existing parser.

        Method ini:
        - tidak menulis output Excel;
        - tidak mengubah source workbook;
        - tidak mengubah SamplingRecord;
        - mempertahankan source worksheet;
        - mengisolasi error per worksheet.
        """

        input_file = Path(
            input_file
        )

        if not input_file.exists():
            raise FileNotFoundError(
                f"File tidak ditemukan: {input_file}"
            )

        reader = WorkbookReader(
            input_file
        )

        processing_result = (
            MultiWorksheetProcessingResult()
        )

        excluded_worksheet_names = {
            name.strip().casefold()
            for name in self.config.excluded_worksheet_names
        }

        for worksheet_info in (
            reader.get_worksheet_info()
        ):

            # -------------------------------------------------
            # Source visibility exclusion
            # -------------------------------------------------

            if not worksheet_info.is_visible:

                processing_result.worksheet_results.append(
                    WorksheetProcessingResult(
                        worksheet_name=(
                            worksheet_info.name
                        ),
                        visibility="HIDDEN",
                        status="EXCLUDED",
                        record_count=0,
                        message=(
                            "Non-visible worksheet"
                        ),
                    )
                )

                continue

            # -------------------------------------------------
            # Configured business exclusion
            # -------------------------------------------------

            normalized_worksheet_name = (
                worksheet_info.name
                .strip()
                .casefold()
            )

            if (
                normalized_worksheet_name
                in excluded_worksheet_names
            ):

                processing_result.worksheet_results.append(
                    WorksheetProcessingResult(
                        worksheet_name=(
                            worksheet_info.name
                        ),
                        visibility="VISIBLE",
                        status="EXCLUDED",
                        record_count=0,
                        message=(
                            "Configured worksheet exclusion"
                        ),
                    )
                )

                continue

            # -------------------------------------------------
            # Candidate worksheet parsing
            # -------------------------------------------------

            try:
                worksheet = reader.load_worksheet(
                    worksheet_info.name
                )

                records = self.parser.parse(
                    worksheet
                )

                sourced_records = [
                    SourcedSamplingRecord(
                        source_worksheet=(
                            worksheet_info.name
                        ),
                        record=record,
                    )
                    for record in records
                ]

                processing_result.records.extend(
                    sourced_records
                )

                processing_result.worksheet_results.append(
                    WorksheetProcessingResult(
                        worksheet_name=(
                            worksheet_info.name
                        ),
                        visibility="VISIBLE",
                        status="SUCCESS",
                        record_count=len(records),
                        message="",
                    )
                )

            except Exception as exc:

                processing_result.worksheet_results.append(
                    WorksheetProcessingResult(
                        worksheet_name=(
                            worksheet_info.name
                        ),
                        visibility="VISIBLE",
                        status="ERROR",
                        record_count=0,
                        message=(
                            f"{type(exc).__name__}: "
                            f"{exc}"
                        ),
                    )
                )

        return processing_result
