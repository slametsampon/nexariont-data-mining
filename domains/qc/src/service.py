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
    """Ringkasan pemrosesan satu worksheet, termasuk worksheet yang dilewati.

    Status SUCCESS berarti parser selesai tanpa exception, bukan persetujuan
    atau validasi QC. Semua atribut merupakan argumen konstruktor dataclass
    frozen; message memiliki nilai bawaan string kosong.

    Attributes:
        worksheet_name (str): Nama worksheet sumber.
        visibility (str): "VISIBLE" atau "HIDDEN"; HIDDEN mencakup seluruh
            worksheet yang dinilai tidak visible.
        status (str): "SUCCESS", "ERROR", atau "EXCLUDED".
        record_count (int): Jumlah candidate record; 0 jika error atau excluded.
        message (str): Keterangan error atau alasan pengecualian.
    """

    worksheet_name: str
    visibility: str
    status: str
    record_count: int
    message: str = ""


@dataclass
class MultiWorksheetProcessingResult:
    """Kumpulan record dan ringkasan pemrosesan seluruh worksheet.

    Kedua daftar dapat diberikan pada konstruktor dataclass. Jika tidak
    diberikan, setiap instance mendapatkan daftar kosongnya sendiri.

    Attributes:
        records (list[SourcedSamplingRecord]): Candidate record beserta sumbernya.
        worksheet_results (list[WorksheetProcessingResult]): Hasil setiap
            worksheet, termasuk yang gagal atau dikecualikan.
        total_records (int): Properti jumlah record pada records.
        success_count (int): Properti jumlah worksheet berstatus SUCCESS.
        error_count (int): Properti jumlah worksheet berstatus ERROR.
        excluded_count (int): Properti jumlah worksheet berstatus EXCLUDED.

    Examples:
        >>> result = MultiWorksheetProcessingResult()
        >>> (result.total_records, result.error_count)
        (0, 0)
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
    """Mengatur pembacaan dan parsing data QC pada tingkat worksheet.

    process_file() memproses satu worksheet terkonfigurasi dan mengekspor hasil.
    process_all_worksheets() mengembalikan hasil gabungan untuk diekspor oleh
    pemanggil, dengan penyaringan worksheet dan isolasi error per worksheet.

    Args:
        config (MiningConfig): Aturan parser, pengecualian worksheet, dan ekspor.
        sampling_point_master (SamplingPointMaster | None): Referensi titik
            sampling opsional; default None.
        domain (str | None): Domain untuk pencocokan master; default None.

    Attributes:
        config (MiningConfig): Konfigurasi yang digunakan.
        parser (ShiftReportParser): Parser yang digunakan ulang antar-worksheet.
        exporter (ExcelExporter): Penulis hasil mode satu worksheet.

    Examples:
        service = QCDataMiningService(MiningConfig())
        result = service.process_all_worksheets(Path("laporan.xlsx"))
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
        """Memproses worksheet terkonfigurasi dan menyimpan hasil ke Excel.

        Args:
            input_file (Path): Workbook sumber .xls atau .xlsx.
            output_file (Path): Lokasi workbook .xlsx hasil.

        Returns:
            int: Jumlah candidate record yang diekspor.

        Raises:
            FileNotFoundError: Workbook sumber tidak ditemukan.
            ValueError: Format atau worksheet tidak sesuai, atau parser gagal
                menemukan blok data maupun tanggal sampling.
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

                diagnostics = list(
                    self.parser.last_diagnostics
                )

                message = " | ".join(
                    diagnostics
                )

                processing_result.worksheet_results.append(
                    WorksheetProcessingResult(
                        worksheet_name=(
                            worksheet_info.name
                        ),
                        visibility="VISIBLE",
                        status="SUCCESS",
                        record_count=len(records),
                        message=message,
                    )
                )

            except Exception as exc:

                diagnostics = list(
                    self.parser.last_diagnostics
                )

                message = (
                    f"{type(exc).__name__}: "
                    f"{exc}"
                )

                if diagnostics:
                    message = (
                        f"{message} | "
                        + " | ".join(diagnostics)
                    )

                processing_result.worksheet_results.append(
                    WorksheetProcessingResult(
                        worksheet_name=(
                            worksheet_info.name
                        ),
                        visibility="VISIBLE",
                        status="ERROR",
                        record_count=0,
                        message=message,
                    )
                )

        return processing_result
