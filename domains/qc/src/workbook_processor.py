from pathlib import Path

from .config import MiningConfig
from .exporter import ExcelExporter
from .sampling_point_master import SamplingPointMaster
from .service import (
    MultiWorksheetProcessingResult,
    QCDataMiningService,
)


class QCWorkbookProcessor:
    """Menjalankan ekstraksi dan ekspor seluruh worksheet satu workbook QC.

    Digunakan bersama oleh alur CLI dan orkestrasi bulanan. process() memuat
    master opsional, menjalankan service multi-worksheet, lalu menulis data,
    kontrol, dan log pemrosesan ke workbook tujuan.

    Args:
        config (MiningConfig | None): Konfigurasi pemrosesan; jika None,
            menggunakan MiningConfig().

    Attributes:
        config (MiningConfig): Konfigurasi efektif untuk service dan exporter.

    Examples:
        processor = QCWorkbookProcessor()
        result = processor.process(
            input_file=Path("laporan.xlsx"), output_file=Path("hasil.xlsx")
        )
    """

    def __init__(
        self,
        config: MiningConfig | None = None,
    ):
        self.config = config or MiningConfig()

    def process(
        self,
        input_file: Path,
        output_file: Path,
        master_file: Path | None = None,
        domain: str | None = None,
    ) -> MultiWorksheetProcessingResult:

        """Memproses workbook sumber dan mengekspor hasil multi-worksheet.

        Args:
            input_file (Path): Workbook .xls atau .xlsx sumber.
            output_file (Path): Lokasi workbook .xlsx hasil.
            master_file (Path | None): Workbook master opsional; default None.
            domain (str | None): Domain master; wajib diberikan bersama master_file.

        Returns:
            MultiWorksheetProcessingResult: Record dan status setiap worksheet.
            Kegagalan parsing per worksheet dicatat di hasil dan tidak menghentikan
            pemrosesan worksheet berikutnya.

        Raises:
            ValueError: Hanya salah satu dari master_file/domain diberikan, format
                sumber tidak didukung, atau worksheet master tidak ditemukan.
            FileNotFoundError: Workbook sumber atau master tidak ditemukan.
        """

        input_file = Path(input_file).resolve()
        output_file = Path(output_file).resolve()

        if (master_file is None) != (domain is None):
            raise ValueError(
                "--master dan --domain harus diberikan bersama-sama"
            )

        sampling_point_master = None

        if master_file is not None:
            sampling_point_master = SamplingPointMaster.load(
                Path(master_file).resolve(),
                worksheet_name=(
                    self.config.sampling_point_master_sheet
                ),
            )

        service = QCDataMiningService(
            self.config,
            sampling_point_master=sampling_point_master,
            domain=domain,
        )

        result = service.process_all_worksheets(
            input_file=input_file,
        )

        exporter = ExcelExporter(
            self.config
        )

        exporter.export_multi_worksheet(
            records=result.records,
            worksheet_results=result.worksheet_results,
            output_file=output_file,
            source_file=input_file,
        )

        return result