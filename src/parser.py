import re

from datetime import datetime, date
from typing import Any

from .config import MiningConfig
from .models import SamplingRecord


class ShiftReportParser:
    """
    Parser QC Shift Report.

    Responsibility:
    - menemukan horizontal data blocks
    - membaca Sampling Date
    - mendeteksi Sampling Point
    - membaca Parameter
    - menghubungkan Parameter dengan Sampling Time
    - menghasilkan normalized SamplingRecord

    Parser tidak bertanggung jawab:
    - membuka file Excel
    - menentukan reader .xls / .xlsx
    - menulis file output

    Dengan demikian parser tidak bergantung langsung
    pada openpyxl maupun xlrd.
    """

    def __init__(self, config: MiningConfig):
        self.config = config

    # ---------------------------------------------------------
    # Public API
    # ---------------------------------------------------------

    def parse(
        self,
        worksheet,
    ) -> list[SamplingRecord]:

        block_columns = self._find_data_blocks(
            worksheet
        )

        if not block_columns:
            raise ValueError(
                "Tidak ditemukan header 'Item' "
                "pada worksheet."
            )

        records: list[SamplingRecord] = []

        for item_col in block_columns:

            sampling_date = self._extract_sampling_date(
                worksheet,
                item_col,
            )

            block_records = self._parse_block(
                worksheet=worksheet,
                item_col=item_col,
                sampling_date=sampling_date,
            )

            records.extend(block_records)

        return records

    # ---------------------------------------------------------
    # Data block detection
    # ---------------------------------------------------------

    def _find_data_blocks(
        self,
        worksheet,
    ) -> list[int]:
        """
        Mencari semua kolom yang memiliki header 'Item'.

        Contoh source:

        A = Item
        B = Min
        C = Max
        D = Unit
        E... = Sampling

        I = Item
        J = Min
        K = Max
        L = Unit
        M... = Sampling

        Dengan demikian satu worksheet dapat memiliki
        lebih dari satu horizontal block.
        """

        block_columns: list[int] = []

        for row in worksheet.iter_rows():

            for cell in row:

                if (
                    self._normalize_text(cell.value)
                    == self.config.item_header.lower()
                ):
                    block_columns.append(
                        cell.column
                    )

        return sorted(
            set(block_columns)
        )

    # ---------------------------------------------------------
    # Date extraction
    # ---------------------------------------------------------

    def _extract_sampling_date(
        self,
        worksheet,
        item_col: int,
    ) -> date:
        """
        Mencari tanggal sampling, misalnya:

            Samp. Date : 22/09/2026

        pada area header horizontal block.
        """

        search_end_col = min(
            worksheet.max_column,
            item_col + 7,
        )

        search_end_row = min(
            worksheet.max_row,
            20,
        )

        for row in worksheet.iter_rows(
            min_row=1,
            max_row=search_end_row,
            min_col=item_col,
            max_col=search_end_col,
        ):

            for cell in row:

                value = cell.value

                if not isinstance(
                    value,
                    str,
                ):
                    continue

                match = re.search(
                    r"(\d{1,2}/\d{1,2}/\d{4})",
                    value,
                )

                if match:

                    return datetime.strptime(
                        match.group(1),
                        "%d/%m/%Y",
                    ).date()

        raise ValueError(
            f"Sampling Date tidak ditemukan "
            f"untuk block column {item_col}."
        )

    # ---------------------------------------------------------
    # Block parser
    # ---------------------------------------------------------

    def _parse_block(
        self,
        worksheet,
        item_col: int,
        sampling_date: date,
    ) -> list[SamplingRecord]:

        records: list[SamplingRecord] = []

        row_number = self._find_first_data_row(
            worksheet,
            item_col,
        )

        while row_number <= worksheet.max_row:

            if self._is_sampling_point_row(
                worksheet,
                row_number,
                item_col,
            ):

                sampling_point = self._cell_value(
                    worksheet,
                    row_number,
                    item_col,
                )

                sampling_times = (
                    self._get_sampling_times(
                        worksheet,
                        row_number,
                        item_col,
                    )
                )

                row_number += 1

                while (
                    row_number
                    <= worksheet.max_row
                ):

                    parameter = self._cell_value(
                        worksheet,
                        row_number,
                        item_col,
                    )

                    # -----------------------------------------
                    # Blank Item -> akhir block SP
                    # -----------------------------------------

                    if self._is_blank(
                        parameter
                    ):
                        break

                    # -----------------------------------------
                    # Total -> akhir parameter SP
                    # -----------------------------------------

                    if self._is_total(
                        parameter
                    ):
                        break

                    # -----------------------------------------
                    # Sampling Point berikutnya
                    # -----------------------------------------

                    if self._is_sampling_point_row(
                        worksheet,
                        row_number,
                        item_col,
                    ):
                        break

                    parameter_records = (
                        self._create_parameter_records(
                            worksheet=worksheet,
                            row_number=row_number,
                            item_col=item_col,
                            sampling_date=sampling_date,
                            sampling_point=str(
                                sampling_point
                            ).strip(),
                            sampling_times=(
                                sampling_times
                            ),
                        )
                    )

                    records.extend(
                        parameter_records
                    )

                    row_number += 1

                continue

            row_number += 1

        return records

    # ---------------------------------------------------------
    # Sampling Point
    # ---------------------------------------------------------

    def _is_sampling_point_row(
        self,
        worksheet,
        row_number: int,
        item_col: int,
    ) -> bool:
        """
        Struktur Sampling Point pada source:

        Item/SP            Min Max Unit    Time...
        T-140 OH (...)                     07.00

        Rule:

        - Item/SP mempunyai nilai
        - Min kosong
        - Max kosong
        - Unit kosong
        - terdapat Sampling Time atau OFF
          pada kolom sampling

        Background cyan bukan satu-satunya rule parser.
        Ini memungkinkan parser bekerja pada .xls maupun
        .xlsx melalui worksheet adapter yang sesuai.
        """

        item = self._cell_value(
            worksheet,
            row_number,
            item_col,
        )

        if self._is_blank(item):
            return False

        if self._is_total(item):
            return False

        minimum = self._cell_value(
            worksheet,
            row_number,
            item_col + 1,
        )

        maximum = self._cell_value(
            worksheet,
            row_number,
            item_col + 2,
        )

        unit = self._cell_value(
            worksheet,
            row_number,
            item_col + 3,
        )

        if not (
            self._is_blank(minimum)
            and self._is_blank(maximum)
            and self._is_blank(unit)
        ):
            return False

        sampling_cells = []

        for offset in range(
            self.config.max_sampling_columns
        ):

            sampling_cells.append(
                self._cell_value(
                    worksheet,
                    row_number,
                    item_col + 4 + offset,
                )
            )

        return any(
            self._is_sampling_time(value)
            or self._is_off(value)
            for value in sampling_cells
        )

    # ---------------------------------------------------------
    # Sampling Time
    # ---------------------------------------------------------

    def _get_sampling_times(
        self,
        worksheet,
        row_number: int,
        item_col: int,
    ) -> list[tuple[int, str]]:
        """
        Membaca seluruh Sampling Time untuk satu
        Sampling Point.

        Contoh:

            07.00 | 11.00 | 13.00

        menghasilkan:

            [
                (0, "07.00"),
                (1, "11.00"),
                (2, "13.00"),
            ]

        OFF tidak menghasilkan sampling record.
        """

        sampling_times: list[
            tuple[int, str]
        ] = []

        for offset in range(
            self.config.max_sampling_columns
        ):

            value = self._cell_value(
                worksheet,
                row_number,
                item_col + 4 + offset,
            )

            if self._is_blank(value):
                continue

            # -----------------------------------------
            # OFF = tidak ada pengukuran
            # -----------------------------------------

            if self._is_off(value):
                continue

            if self._is_sampling_time(value):

                sampling_times.append(
                    (
                        offset,
                        self._format_time(
                            value
                        ),
                    )
                )

        return sampling_times

    # ---------------------------------------------------------
    # Parameter record creation
    # ---------------------------------------------------------

    def _create_parameter_records(
        self,
        worksheet,
        row_number: int,
        item_col: int,
        sampling_date: date,
        sampling_point: str,
        sampling_times: list[
            tuple[int, str]
        ],
    ) -> list[SamplingRecord]:

        records: list[
            SamplingRecord
        ] = []

        parameter = self._cell_value(
            worksheet,
            row_number,
            item_col,
        )

        minimum = self._clean_standard(
            self._cell_value(
                worksheet,
                row_number,
                item_col + 1,
            )
        )

        maximum = self._clean_standard(
            self._cell_value(
                worksheet,
                row_number,
                item_col + 2,
            )
        )

        unit = self._cell_value(
            worksheet,
            row_number,
            item_col + 3,
        )

        for (
            offset,
            sampling_time,
        ) in sampling_times:

            value = self._cell_value(
                worksheet,
                row_number,
                item_col + 4 + offset,
            )

            # -----------------------------------------
            # Tidak membuat record untuk nilai kosong
            # -----------------------------------------

            if self._is_blank(value):
                continue

            record = SamplingRecord(
                sampling_date=sampling_date,
                sampling_time=sampling_time,
                sampling_point=(
                    sampling_point
                ),
                parameter=str(
                    parameter
                ).strip(),
                unit=self._clean_value(
                    unit
                ),
                minimum=minimum,
                maximum=maximum,
                value=value,
            )

            records.append(
                record
            )

        return records

    # ---------------------------------------------------------
    # Header / data positioning
    # ---------------------------------------------------------

    def _find_first_data_row(
        self,
        worksheet,
        item_col: int,
    ) -> int:
        """
        Mencari row header 'Time'.

        Data/SP diasumsikan mulai pada row berikutnya
        sesuai struktur source yang sudah diverifikasi.
        """

        for row_number in range(
            1,
            min(
                30,
                worksheet.max_row,
            ) + 1,
        ):

            value = self._cell_value(
                worksheet,
                row_number,
                item_col + 4,
            )

            if (
                self._normalize_text(
                    value
                )
                == "time"
            ):
                return row_number + 1

        # ---------------------------------------------
        # Fallback konservatif.
        #
        # Tidak menginfer lokasi data lain apabila
        # header Time tidak ditemukan.
        # ---------------------------------------------

        return 1

    # ---------------------------------------------------------
    # Worksheet helper
    # ---------------------------------------------------------

    @staticmethod
    def _cell_value(
        worksheet,
        row: int,
        column: int,
    ) -> Any:
        """
        Mengakses nilai cell melalui worksheet interface.

        Interface minimum yang diperlukan:

            worksheet.max_row
            worksheet.max_column
            worksheet.cell(row=?, column=?).value
            worksheet.iter_rows(...)

        Dengan demikian parser tidak bergantung
        langsung pada openpyxl ataupun xlrd.
        """

        if column > worksheet.max_column:
            return None

        if row > worksheet.max_row:
            return None

        return worksheet.cell(
            row=row,
            column=column,
        ).value

    # ---------------------------------------------------------
    # Value classification
    # ---------------------------------------------------------

    @staticmethod
    def _is_blank(
        value: Any,
    ) -> bool:

        if value is None:
            return True

        if isinstance(
            value,
            str,
        ):
            return (
                value.strip() == ""
            )

        return False

    def _is_total(
        self,
        value: Any,
    ) -> bool:

        return (
            self._normalize_text(
                value
            )
            == self.config.total_keyword.lower()
        )

    def _is_off(
        self,
        value: Any,
    ) -> bool:

        return (
            self._normalize_text(
                value
            )
            == self.config.off_keyword.lower()
        )

    @staticmethod
    def _normalize_text(
        value: Any,
    ) -> str:

        if value is None:
            return ""

        return str(
            value
        ).strip().lower()

    # ---------------------------------------------------------
    # Sampling Time classification
    # ---------------------------------------------------------

    @staticmethod
    def _is_sampling_time(
        value: Any,
    ) -> bool:
        """
        Sampling time yang didukung:

            07.00
            7.00
            07:00
            7:00

        serta object time/datetime yang mempunyai
        attribute hour dan minute.
        """

        if value is None:
            return False

        # ---------------------------------------------
        # Python time/datetime
        # ---------------------------------------------

        if (
            hasattr(value, "hour")
            and hasattr(value, "minute")
        ):
            return True

        text = str(
            value
        ).strip()

        return bool(
            re.fullmatch(
                r"\d{1,2}[.:]\d{2}",
                text,
            )
        )

    @staticmethod
    def _format_time(
        value: Any,
    ) -> str:
        """
        Normalisasi separator waktu menjadi titik.

        07:00 -> 07.00
        """

        if (
            hasattr(value, "hour")
            and hasattr(value, "minute")
        ):

            return (
                f"{value.hour:02d}."
                f"{value.minute:02d}"
            )

        text = str(
            value
        ).strip()

        return text.replace(
            ":",
            ".",
        )

    # ---------------------------------------------------------
    # Value cleaning
    # ---------------------------------------------------------

    @staticmethod
    def _clean_standard(
        value: Any,
    ) -> Any:
        """
        Standard/SOC dipertahankan sesuai source.

        Contoh:

            0       -> 0
            0.25    -> 0.25
            "-"     -> "-"

        '-' tidak dikonversi menjadi 0 karena:

            '-' != 0

        Parser tidak menginfer batas operasi yang tidak
        dinyatakan pada source.
        """

        if value is None:
            return None

        if isinstance(
            value,
            str,
        ):

            value = value.strip()

            if value == "":
                return None

        return value

    @staticmethod
    def _clean_value(
        value: Any,
    ) -> Any:
        """
        Membersihkan whitespace tanpa mengubah
        meaning/value source.

        Contoh:

            " wt % "  -> "wt %"
            "<0.01"   -> "<0.01"
            "-"       -> "-"
        """

        if value is None:
            return None

        if isinstance(
            value,
            str,
        ):

            cleaned = value.strip()

            if cleaned == "":
                return None

            return cleaned

        return value