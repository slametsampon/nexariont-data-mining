from pathlib import Path
from typing import Sequence

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
)
from openpyxl.utils import get_column_letter

from .config import MiningConfig
from .models import (
    SamplingRecord,
    SourcedSamplingRecord,
)


class ExcelExporter:
    def __init__(
        self,
        config: MiningConfig,
    ):
        self.config = config

    # =====================================================
    # EXISTING SINGLE-WORKSHEET EXPORT
    # =====================================================

    def export(
        self,
        records: Sequence[SamplingRecord],
        output_file: Path,
        source_file: Path,
    ) -> None:
        """
        Existing single-worksheet export path.

        Dipertahankan untuk regression compatibility.
        """

        output_file = Path(
            output_file
        )

        output_file.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        workbook = Workbook()

        worksheet = workbook.active
        worksheet.title = (
            self.config.output_sheet_name
        )

        self._write_data(
            worksheet,
            records,
        )

        self._write_control_sheet(
            workbook,
            records,
            source_file,
        )

        workbook.save(
            output_file
        )

    # =====================================================
    # MULTI-WORKSHEET EXPORT — CP4.3
    # =====================================================

    def export_multi_worksheet(
        self,
        records: Sequence[SourcedSamplingRecord],
        output_file: Path,
        source_file: Path,
    ) -> None:
        """
        Export consolidated candidate records dari
        multi-worksheet processing.

        Source worksheet dipertahankan sebagai provenance
        pada setiap record.

        CP4.3:
        - menulis consolidated Data_Mining;
        - mempertahankan source worksheet;
        - membuat Control minimum multi-worksheet;
        - belum membuat Processing_Log.

        Processing_Log merupakan scope CP4.4.

        Candidate record bukan pernyataan QC validation
        atau approval.
        """

        output_file = Path(
            output_file
        )

        source_file = Path(
            source_file
        )

        output_file.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        workbook = Workbook()

        worksheet = workbook.active
        worksheet.title = (
            self.config.output_sheet_name
        )

        self._write_multi_worksheet_data(
            worksheet=worksheet,
            records=records,
        )

        self._write_multi_worksheet_control_sheet(
            workbook=workbook,
            records=records,
            source_file=source_file,
        )

        workbook.save(
            output_file
        )

    # =====================================================
    # EXISTING SINGLE-WORKSHEET DATA WRITER
    # =====================================================

    def _write_data(
        self,
        worksheet,
        records: Sequence[SamplingRecord],
    ) -> None:

        worksheet.append(
            list(
                self.config.output_headers
            )
        )

        for record in records:
            worksheet.append(
                record.to_row()
            )

        self._format_header(
            worksheet
        )

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        widths = {
            "A": 13,
            "B": 10,
            "C": 40,
            "D": 22,
            "E": 12,
            "F": 12,
            "G": 12,
            "H": 15,
        }

        for column, width in (
            widths.items()
        ):
            worksheet.column_dimensions[
                column
            ].width = width

        for cell in worksheet["A"][1:]:
            cell.number_format = (
                "dd/mm/yyyy"
            )

    # =====================================================
    # MULTI-WORKSHEET DATA WRITER
    # =====================================================

    def _write_multi_worksheet_data(
        self,
        worksheet,
        records: Sequence[SourcedSamplingRecord],
    ) -> None:
        """
        Menulis consolidated multi-worksheet records.

        Struktur:
        Source Worksheet
        Tanggal
        Jam
        Sampling Point (SP)
        Parameter
        Unit
        Min
        Max
        Nilai
        """

        headers = [
            "Source Worksheet",
            *self.config.output_headers,
        ]

        worksheet.append(
            list(headers)
        )

        for sourced_record in records:
            worksheet.append(
                sourced_record.to_row()
            )

        self._format_header(
            worksheet
        )

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        widths = {
            "A": 24,
            "B": 13,
            "C": 10,
            "D": 40,
            "E": 22,
            "F": 12,
            "G": 12,
            "H": 12,
            "I": 15,
        }

        for column, width in (
            widths.items()
        ):
            worksheet.column_dimensions[
                column
            ].width = width

        # Pada multi-worksheet export:
        #
        # A = Source Worksheet
        # B = Tanggal
        #
        # Existing single-worksheet export tetap
        # menggunakan kolom A sebagai Tanggal.
        for cell in worksheet["B"][1:]:
            cell.number_format = (
                "dd/mm/yyyy"
            )

    # =====================================================
    # SHARED HEADER FORMAT
    # =====================================================

    @staticmethod
    def _format_header(
        worksheet,
    ) -> None:

        fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )

        font = Font(
            color="FFFFFF",
            bold=True,
        )

        alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

        for cell in worksheet[1]:
            cell.fill = fill
            cell.font = font
            cell.alignment = alignment

    # =====================================================
    # EXISTING SINGLE-WORKSHEET CONTROL
    # =====================================================

    def _write_control_sheet(
        self,
        workbook,
        records,
        source_file,
    ) -> None:

        worksheet = workbook.create_sheet(
            self.config.control_sheet_name
        )

        controls = [
            (
                "Item",
                "Keterangan",
            ),
            (
                "Source file",
                source_file.name,
            ),
            (
                "Source worksheet",
                self.config.worksheet_name,
            ),
            (
                "Output records",
                len(records),
            ),
            (
                "OFF handling",
                (
                    "Tidak dibuat sebagai "
                    "record pengukuran"
                ),
            ),
            (
                "Total handling",
                (
                    "Bukan parameter; menjadi "
                    "batas pembacaan SP"
                ),
            ),
            (
                "Blank value",
                (
                    "Tidak dibuat sebagai "
                    "record pengukuran"
                ),
            ),
            (
                "STD '-'",
                (
                    "Dipertahankan sesuai source; "
                    "tidak dianggap 0"
                ),
            ),
        ]

        for row in controls:
            worksheet.append(
                row
            )

        self._format_header(
            worksheet
        )

        worksheet.column_dimensions[
            "A"
        ].width = 24

        worksheet.column_dimensions[
            "B"
        ].width = 65

    # =====================================================
    # MULTI-WORKSHEET CONTROL — CP4.3
    # =====================================================

    def _write_multi_worksheet_control_sheet(
        self,
        workbook,
        records: Sequence[SourcedSamplingRecord],
        source_file: Path,
    ) -> None:
        """
        Control minimum untuk multi-worksheet export.

        Tidak menyatakan candidate records sebagai
        QC validated atau approved.

        Detail processing per worksheet akan
        ditambahkan pada CP4.4 melalui Processing_Log.
        """

        worksheet = workbook.create_sheet(
            self.config.control_sheet_name
        )

        controls = [
            (
                "Item",
                "Keterangan",
            ),
            (
                "Source file",
                source_file.name,
            ),
            (
                "Processing mode",
                "Multi-worksheet",
            ),
            (
                "Candidate records",
                len(records),
            ),
            (
                "Record status",
                (
                    "Candidate parser output; "
                    "not QC validation/approval"
                ),
            ),
            (
                "Worksheet traceability",
                (
                    "Source Worksheet retained "
                    "per record"
                ),
            ),
            (
                "OFF handling",
                (
                    "Tidak dibuat sebagai "
                    "record pengukuran"
                ),
            ),
            (
                "Total handling",
                (
                    "Bukan parameter; menjadi "
                    "batas pembacaan SP"
                ),
            ),
            (
                "Blank value",
                (
                    "Tidak dibuat sebagai "
                    "record pengukuran"
                ),
            ),
            (
                "STD '-'",
                (
                    "Dipertahankan sesuai source; "
                    "tidak dianggap 0"
                ),
            ),
        ]

        for row in controls:
            worksheet.append(
                row
            )

        self._format_header(
            worksheet
        )

        worksheet.column_dimensions[
            "A"
        ].width = 24

        worksheet.column_dimensions[
            "B"
        ].width = 65
        