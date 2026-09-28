from pathlib import Path

from .config import MiningConfig
from .exporter import ExcelExporter
from .sampling_point_master import SamplingPointMaster
from .service import (
    MultiWorksheetProcessingResult,
    QCDataMiningService,
)


class QCWorkbookProcessor:
    """
    Shared application entry point untuk memproses
    satu QC workbook end-to-end.
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