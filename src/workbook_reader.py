from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlrd
from openpyxl import load_workbook


@dataclass(frozen=True)
class WorksheetInfo:
    """
    Metadata worksheet untuk discovery dan visibility filtering.

    visibility menyimpan nilai visibility dari source workbook
    dalam bentuk integer internal:

        0 = visible
        1 = hidden
        2 = very hidden

    Worksheet hanya menjadi candidate parsing jika visible.
    """

    name: str
    visibility: int

    @property
    def is_visible(self) -> bool:
        return self.visibility == 0


class CellAdapter:
    """
    Minimal cell interface yang digunakan parser.
    """

    def __init__(
        self,
        value: Any,
        row: int,
        column: int,
    ):
        self.value = value
        self.row = row
        self.column = column


class XlsWorksheetAdapter:
    """
    Adapter agar worksheet xlrd (.xls) menyediakan
    interface minimum yang dibutuhkan ShiftReportParser.

    Tidak membuat file .xlsx sementara.
    Original .xls hanya dibaca.
    """

    def __init__(self, sheet: xlrd.sheet.Sheet):
        self._sheet = sheet

        self.max_row = sheet.nrows
        self.max_column = sheet.ncols

    def cell(
        self,
        row: int,
        column: int,
    ) -> CellAdapter:

        # Parser menggunakan index Excel mulai dari 1,
        # sedangkan xlrd menggunakan index mulai dari 0.
        row_index = row - 1
        col_index = column - 1

        if (
            row_index < 0
            or col_index < 0
            or row_index >= self._sheet.nrows
            or col_index >= self._sheet.ncols
        ):
            return CellAdapter(
                value=None,
                row=row,
                column=column,
            )

        xls_cell = self._sheet.cell(
            row_index,
            col_index,
        )

        value = self._convert_value(
            xls_cell
        )

        return CellAdapter(
            value=value,
            row=row,
            column=column,
        )

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
    ):

        if max_row is None:
            max_row = self.max_row

        if max_col is None:
            max_col = self.max_column

        for row_number in range(
            min_row,
            max_row + 1,
        ):

            row_cells = []

            for column_number in range(
                min_col,
                max_col + 1,
            ):

                row_cells.append(
                    self.cell(
                        row=row_number,
                        column=column_number,
                    )
                )

            yield tuple(row_cells)

    @staticmethod
    def _convert_value(
        cell: xlrd.sheet.Cell,
    ) -> Any:
        """
        Konversi minimum dan konservatif.

        Blank -> None

        Text seperti:
            <0.01
            -
            OFF
        tetap dipertahankan sebagai text.

        Numeric tetap numeric.
        """

        if cell.ctype in (
            xlrd.XL_CELL_EMPTY,
            xlrd.XL_CELL_BLANK,
        ):
            return None

        if cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value

        if cell.ctype == xlrd.XL_CELL_NUMBER:

            value = cell.value

            # 5.0 -> 5
            # 9.2282 -> 9.2282
            if float(value).is_integer():
                return int(value)

            return value

        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return bool(cell.value)

        if cell.ctype == xlrd.XL_CELL_ERROR:
            return cell.value

        return cell.value


class WorkbookReader:
    """
    Input reader/factory.

    Supported:
    - .xls  : xlrd + adapter
    - .xlsx : openpyxl

    Parser tidak perlu mengetahui format source.
    """

    def __init__(
        self,
        file_path: Path,
    ):
        self.file_path = Path(file_path)

    def get_worksheet_names(
        self,
    ) -> list[str]:
        """
        Mengembalikan seluruh nama worksheet dari workbook.

        Supported:
        - .xls  : xlrd
        - .xlsx : openpyxl

        Method ini hanya melakukan worksheet discovery.
        Tidak melakukan parsing data.
        """

        if not self.file_path.exists():
            raise FileNotFoundError(
                f"File tidak ditemukan: {self.file_path}"
            )

        extension = self.file_path.suffix.lower()

        if extension == ".xls":

            workbook = xlrd.open_workbook(
                filename=str(self.file_path),
                on_demand=True,
            )

            try:
                return workbook.sheet_names()

            finally:
                workbook.release_resources()

        if extension == ".xlsx":

            workbook = load_workbook(
                filename=self.file_path,
                data_only=True,
                read_only=True,
            )

            try:
                return list(
                    workbook.sheetnames
                )

            finally:
                workbook.close()

        raise ValueError(
            "Format input tidak didukung. "
            "Gunakan file .xls atau .xlsx"
        )

    def get_worksheet_info(
        self,
    ) -> list[WorksheetInfo]:
        """
        Mengembalikan metadata seluruh worksheet,
        termasuk visibility.

        Visible worksheet dapat menjadi candidate parsing.
        Hidden/non-visible worksheet tidak menjadi candidate.
        """

        if not self.file_path.exists():
            raise FileNotFoundError(
                f"File tidak ditemukan: {self.file_path}"
            )

        extension = self.file_path.suffix.lower()

        if extension == ".xls":

            workbook = xlrd.open_workbook(
                filename=str(self.file_path),
                on_demand=True,
            )

            try:
                return [
                    WorksheetInfo(
                        name=sheet.name,
                        visibility=sheet.visibility,
                    )
                    for sheet in workbook.sheets()
                ]

            finally:
                workbook.release_resources()

        if extension == ".xlsx":

            workbook = load_workbook(
                filename=self.file_path,
                data_only=True,
                read_only=True,
            )

            try:
                visibility_map = {
                    "visible": 0,
                    "hidden": 1,
                    "veryHidden": 2,
                }

                return [
                    WorksheetInfo(
                        name=sheet.title,
                        visibility=visibility_map.get(
                            sheet.sheet_state,
                            -1,
                        ),
                    )
                    for sheet in workbook.worksheets
                ]

            finally:
                workbook.close()

        raise ValueError(
            "Format input tidak didukung. "
            "Gunakan file .xls atau .xlsx"
        )

    def get_visible_worksheet_names(
        self,
    ) -> list[str]:
        """
        Mengembalikan hanya worksheet yang visible.

        Hidden/non-visible worksheet tidak menjadi
        candidate parsing.
        """

        return [
            worksheet.name
            for worksheet in self.get_worksheet_info()
            if worksheet.is_visible
        ]

    def load_worksheet(
        self,
        worksheet_name: str,
    ):

        if not self.file_path.exists():
            raise FileNotFoundError(
                f"File tidak ditemukan: {self.file_path}"
            )

        extension = self.file_path.suffix.lower()

        if extension == ".xls":
            return self._load_xls(
                worksheet_name
            )

        if extension == ".xlsx":
            return self._load_xlsx(
                worksheet_name
            )

        raise ValueError(
            "Format input tidak didukung. "
            "Gunakan file .xls atau .xlsx"
        )

    def _load_xls(
        self,
        worksheet_name: str,
    ) -> XlsWorksheetAdapter:

        workbook = xlrd.open_workbook(
            filename=str(self.file_path),
            on_demand=True,
        )

        if worksheet_name not in workbook.sheet_names():

            available = ", ".join(
                workbook.sheet_names()
            )

            raise ValueError(
                f"Worksheet '{worksheet_name}' "
                f"tidak ditemukan. "
                f"Worksheet tersedia: {available}"
            )

        sheet = workbook.sheet_by_name(
            worksheet_name
        )

        return XlsWorksheetAdapter(
            sheet
        )

    def _load_xlsx(
        self,
        worksheet_name: str,
    ):

        workbook = load_workbook(
            filename=self.file_path,
            data_only=True,
            read_only=True,
        )

        if worksheet_name not in workbook.sheetnames:

            available = ", ".join(
                workbook.sheetnames
            )

            raise ValueError(
                f"Worksheet '{worksheet_name}' "
                f"tidak ditemukan. "
                f"Worksheet tersedia: {available}"
            )

        return workbook[
            worksheet_name
        ]