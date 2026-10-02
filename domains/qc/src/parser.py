import re

from datetime import datetime, date
from typing import Any

from .config import MiningConfig
from .models import SamplingRecord
from .sampling_point_master import SamplingPointMaster


class ShiftReportParser:
    """Mengubah blok laporan shift QC menjadi record pengukuran terstruktur.

    Parser menemukan blok horizontal, membaca tanggal, mengenali titik
    sampling, dan mengaitkan parameter dengan waktu sampling. Worksheet cukup
    menyediakan cell(), iter_rows(), max_row, dan max_column; pembukaan serta
    penulisan file dilakukan oleh komponen lain.

    Args:
        config (MiningConfig): Label header dan aturan pembacaan data.
        sampling_point_master (SamplingPointMaster | None): Referensi identitas
            terkontrol; default None. Bersama domain, memungkinkan pencarian
            blok tanpa header Item.
        domain (str | None): Domain pencocokan master; default None.

    Attributes:
        config (MiningConfig): Konfigurasi parser.
        sampling_point_master (SamplingPointMaster | None): Referensi opsional.
        domain (str | None): Domain sumber untuk pencocokan identitas.

    Examples:
        reader = WorkbookReader(Path("laporan.xlsx"))
        parser = ShiftReportParser(MiningConfig())
        records = parser.parse(reader.load_worksheet("shift-pagi"))
    """

    def __init__(
        self,
        config: MiningConfig,
        sampling_point_master: SamplingPointMaster | None = None,
        domain: str | None = None,
    ):
        self.config = config
        self.sampling_point_master = sampling_point_master
        self.domain = domain
        self.last_diagnostics: list[str] = []

    # ---------------------------------------------------------
    # Public API
    # ---------------------------------------------------------

    def parse(
        self,
        worksheet,
    ) -> list[SamplingRecord]:

        """Membaca seluruh blok data dalam satu worksheet.

        Args:
            worksheet: Worksheet openpyxl atau adapter dengan antarmuka sel
                dan iterasi baris yang digunakan parser.

        Returns:
            list[SamplingRecord]: Candidate record dari seluruh blok yang ditemukan.

        Raises:
            ValueError: Blok data tidak ditemukan, tanggal sampling tidak tersedia,
                atau teks tanggal tidak dapat dikonversi menjadi tanggal yang valid.
        """

        # Diagnostics are per worksheet/parse call.
        self.last_diagnostics = []

        block_columns = self._find_data_blocks(
            worksheet
        )

        if not block_columns:
            raise ValueError(
                "Tidak ditemukan header 'Item' "
                "pada worksheet."
            )

        # Record controlled recognition/mapping diagnostics before extraction.
        # This does not change extraction decisions or source data.
        self._collect_sampling_point_diagnostics(
            worksheet=worksheet,
            block_columns=block_columns,
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
        Mencari seluruh kolom horizontal data block yang terkontrol.

        Header ``Item`` tetap menjadi discovery mechanism utama. Bila
        controlled Sampling Point Master dan domain tersedia, kolom yang
        berisi Sampling Point yang exact-match terhadap Master digabungkan
        dengan header-derived block. Ini mendukung mixed worksheet ketika
        sebagian horizontal block memiliki header ``Item`` dan block lain
        tidak memilikinya.

        Tanpa controlled Master/domain, behavior existing dipertahankan:
        hanya block yang mempunyai header ``Item`` yang dikembalikan. Tidak
        ada fuzzy/pattern inference.
        """

        block_columns: set[int] = set()

        # -----------------------------------------------------
        # Header-based discovery.
        # -----------------------------------------------------

        for row in worksheet.iter_rows():
            for cell in row:
                if (
                    self._normalize_text(cell.value)
                    == self.config.item_header.lower()
                ):
                    block_columns.add(
                        cell.column
                    )

        # -----------------------------------------------------
        # Tanpa controlled Master/domain, pertahankan behavior
        # existing: hanya header-derived blocks.
        # -----------------------------------------------------

        if (
            self.sampling_point_master is None
            or self.domain is None
        ):
            return sorted(block_columns)

        # -----------------------------------------------------
        # Controlled Master augmentation.
        #
        # Digabungkan dengan header-derived blocks agar mixed
        # worksheet dapat dikenali. Sampling Point yang tidak
        # match terhadap Master tidak diinfer/fuzzy-matched.
        # -----------------------------------------------------

        for row in worksheet.iter_rows():
            for cell in row:
                if (
                    self.sampling_point_master.match(
                        self.domain,
                        cell.value,
                    )
                    is not None
                ):
                    block_columns.add(
                        cell.column
                    )

        return sorted(block_columns)

    # ---------------------------------------------------------
    # Extraction diagnostics / observability
    # ---------------------------------------------------------

    def _collect_sampling_point_diagnostics(
        self,
        worksheet,
        block_columns: list[int],
    ) -> None:
        """Collect non-destructive Sampling Point recognition diagnostics.

        The diagnostic scan is intentionally separate from extraction logic.
        It reports two conditions that previously could remain silent:

        - structurally plausible Sampling Point rows that do not match the
          controlled Master; and
        - source Sampling Point labels that are mapped to a different Master
          output identity.

        Candidate columns are limited to already discovered blocks plus columns
        whose ``Time`` header is found in the same position expected by the
        existing parser. Diagnostics never create records and never change the
        set of parsed block columns.
        """

        if (
            self.sampling_point_master is None
            or self.domain is None
        ):
            return

        discovered_columns = set(block_columns)
        diagnostic_columns = set(discovered_columns)

        # A Time header is existing structural evidence of a possible block.
        # It is used only for diagnostic coverage, never to add extraction blocks.
        for item_col in range(1, max(1, worksheet.max_column - 3)):
            time_col = item_col + 4
            if time_col > worksheet.max_column:
                break
            for row_number in range(1, min(30, worksheet.max_row) + 1):
                if (
                    self._normalize_text(
                        self._cell_value(
                            worksheet,
                            row_number,
                            time_col,
                        )
                    )
                    == "time"
                ):
                    diagnostic_columns.add(item_col)
                    break

        diagnostics: list[str] = []
        seen: set[tuple] = set()
        undiscovered_columns_reported: set[int] = set()

        for item_col in sorted(diagnostic_columns):
            for row_number in range(1, worksheet.max_row + 1):
                if not self._is_structural_sampling_point_candidate(
                    worksheet,
                    row_number,
                    item_col,
                ):
                    continue

                source_value = self._cell_value(
                    worksheet,
                    row_number,
                    item_col,
                )
                source_text = str(source_value).strip()
                location = f"R{row_number}C{item_col}"

                matched = self.sampling_point_master.match(
                    self.domain,
                    source_value,
                )

                if matched is None:
                    key = ("CANDIDATE_UNMATCHED", location, source_text)
                    if key not in seen:
                        diagnostics.append(
                            "REVIEW SP_CANDIDATE_UNMATCHED "
                            f"{location} source={source_text!r}"
                        )
                        seen.add(key)

                    if (
                        item_col not in discovered_columns
                        and item_col not in undiscovered_columns_reported
                    ):
                        diagnostics.append(
                            "WARNING BLOCK_NOT_DISCOVERED "
                            f"C{item_col} candidate={location} "
                            f"source={source_text!r}"
                        )
                        undiscovered_columns_reported.add(item_col)
                    continue

                output_identity = self.sampling_point_master.output_identity(
                    self.domain,
                    source_value,
                    matched,
                )
                canonical_identity = str(output_identity).strip()

                if source_text != canonical_identity:
                    key = (
                        "MAPPED",
                        location,
                        source_text,
                        canonical_identity,
                    )
                    if key not in seen:
                        diagnostics.append(
                            "TRACE SP_MAPPED "
                            f"{location} source={source_text!r} "
                            f"-> {canonical_identity!r}"
                        )
                        seen.add(key)
                else:
                    key = (
                        "RECOGNIZED",
                        location,
                        source_text,
                        canonical_identity,
                    )
                    if key not in seen:
                        diagnostics.append(
                            "TRACE SP_RECOGNIZED "
                            f"{location} source={source_text!r} "
                            f"-> {canonical_identity!r}"
                        )
                        seen.add(key)

        self.last_diagnostics = diagnostics

    def _is_structural_sampling_point_candidate(
        self,
        worksheet,
        row_number: int,
        item_col: int,
    ) -> bool:
        """Return the pre-Master structural Sampling Point signature only.

        This helper intentionally mirrors the structural rule that existed
        before Master recognition became the gate. It is used solely for
        diagnostics, so an unmatched row is reported for review rather than
        silently converted into a Sampling Point.
        """

        item = self._cell_value(
            worksheet,
            row_number,
            item_col,
        )

        if self._is_blank(item) or self._is_total(item):
            return False

        # Verified diagnostic false-positives:
        # ``pH`` and ``NPG+Free water`` are QC parameter labels,
        # not Sampling Point identities.
        # These exclusions are diagnostic-only; extraction logic is unchanged.
        normalized_item = self._normalize_text(item)
        if normalized_item in {
            "ph",
            "npg+free water",
        }:
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

        sampling_cells = [
            self._cell_value(
                worksheet,
                row_number,
                item_col + 4 + offset,
            )
            for offset in range(self.config.max_sampling_columns)
        ]

        return any(
            self._is_sampling_time(value)
            or self._is_off(value)
            for value in sampling_cells
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

        report_date = self._extract_report_level_date(
            worksheet
        )

        if report_date is not None:
            return report_date

        raise ValueError(
            f"Sampling Date tidak ditemukan "
            f"untuk block column {item_col}."
        )

    @classmethod
    def _extract_report_level_date(
        cls,
        worksheet,
    ) -> date | None:
        """
        Fallback untuk source yang menyatakan tanggal pada
        field report-level 'Date/Day'.

        Contoh source yang sudah diverifikasi:

            Date/Day : Senin, 24 Agustus 2026

        Method ini hanya membaca tanggal textual Indonesia
        yang terikat pada field Date/Day. Tidak menginfer
        tanggal dari filename, worksheet name, file timestamp,
        atau cell lain yang tidak terikat pada field tersebut.
        """

        search_end_row = min(
            worksheet.max_row,
            20,
        )

        for row_number in range(
            1,
            search_end_row + 1,
        ):

            for column_number in range(
                1,
                worksheet.max_column + 1,
            ):

                value = cls._cell_value(
                    worksheet,
                    row_number,
                    column_number,
                )

                if not isinstance(
                    value,
                    str,
                ):
                    continue

                normalized = (
                    cls._normalize_text(value)
                    .rstrip(":")
                    .strip()
                )

                if normalized != "date/day":
                    continue

                search_end_col = min(
                    worksheet.max_column,
                    column_number + 7,
                )

                for date_col in range(
                    column_number,
                    search_end_col + 1,
                ):

                    candidate = cls._cell_value(
                        worksheet,
                        row_number,
                        date_col,
                    )

                    parsed_date = (
                        cls._parse_indonesian_textual_date(
                            candidate
                        )
                    )

                    if parsed_date is not None:
                        return parsed_date

        return None

    @staticmethod
    def _parse_indonesian_textual_date(
        value: Any,
    ) -> date | None:
        """
        Parse textual date Indonesia yang digunakan pada
        report-level Date/Day.

        Contoh:
            Senin, 24 Agustus 2026
            Minggu, 30 Agustus 2026
            24 Agustus 2026
        """

        if not isinstance(
            value,
            str,
        ):
            return None

        match = re.search(
            r"\b(\d{1,2})\s+"
            r"(Januari|Februari|Maret|April|Mei|Juni|"
            r"Juli|Agustus|September|Oktober|November|Desember)"
            r"\s+(\d{4})\b",
            value,
            flags=re.IGNORECASE,
        )

        if not match:
            return None

        month_numbers = {
            "januari": 1,
            "februari": 2,
            "maret": 3,
            "april": 4,
            "mei": 5,
            "juni": 6,
            "juli": 7,
            "agustus": 8,
            "september": 9,
            "oktober": 10,
            "november": 11,
            "desember": 12,
        }

        day = int(match.group(1))
        month = month_numbers[match.group(2).lower()]
        year = int(match.group(3))

        try:
            return date(year, month, day)
        except ValueError:
            return None

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

                if self.sampling_point_master is not None and self.domain is not None:
                    matched_sp = self.sampling_point_master.match(self.domain, sampling_point)
                    if matched_sp is not None:
                        sampling_point = self.sampling_point_master.output_identity(
                            self.domain,
                            sampling_point,
                            matched_sp,
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

        # Controlled master recognition takes precedence when configured.
        # No fuzzy inference: SamplingPointMaster.match() is conservative.
        if self.sampling_point_master is not None and self.domain is not None:
            return self.sampling_point_master.match(self.domain, item) is not None

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