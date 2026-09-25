from pathlib import Path

from .config import MiningConfig
from .workbook_reader import WorkbookReader
from .parser import ShiftReportParser
from .exporter import ExcelExporter


class QCDataMiningService:
    """
    Application service / orchestrator.

    Tidak menangani detail parsing dan tidak
    menangani detail formatting Excel.
    """

    def __init__(
        self,
        config: MiningConfig,
    ):
        self.config = config

        self.parser = ShiftReportParser(
            config
        )

        self.exporter = ExcelExporter(
            config
        )

    def process_file(
        self,
        input_file: Path,
        output_file: Path,
    ) -> int:

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