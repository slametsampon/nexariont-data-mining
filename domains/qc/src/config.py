from dataclasses import dataclass, field


@dataclass(frozen=True)
class MiningConfig:
    """Konfigurasi tetap untuk pembacaan, parsing, dan ekspor data QC.

    Dataclass bersifat frozen. Setiap atribut dapat diberikan sebagai argumen
    konstruktor; nilai bawaan mengikuti format laporan QC yang digunakan proyek.

    Attributes:
        sampling_point_master_sheet (str): Worksheet master; default "QA Review".
        worksheet_name (str): Worksheet mode tunggal; default "shift-pagi".
        excluded_worksheet_names (tuple[str, ...]): Nama worksheet yang dilewati
            pada mode multi-worksheet, tanpa membedakan kapitalisasi.
        item_header (str): Penanda kolom item; default "Item".
        std_header (str): Label standar yang disediakan konfigurasi; default "STD".
        unit_header (str): Label satuan; default "Unit".
        off_keyword (str): Penanda kondisi tidak beroperasi; default "OFF".
        total_keyword (str): Penanda baris total; default "Total".
        max_sampling_columns (int): Batas kolom waktu per blok; default 4.
        output_sheet_name (str): Nama worksheet hasil; default "Data_Mining".
        control_sheet_name (str): Nama worksheet kontrol; default "Control".
        output_headers (tuple[str, ...]): Delapan header keluaran, berurutan:
            Tanggal, Jam, Sampling Point (SP), Parameter, Unit, Min, Max, Nilai.

    Examples:
        >>> config = MiningConfig(worksheet_name="shift-malam")
        >>> config.max_sampling_columns
        4
    """

    sampling_point_master_sheet: str = "QA Review"
    worksheet_name: str = "shift-pagi"

    # Visible worksheets yang secara eksplisit bukan
    # candidate QC data mining.
    excluded_worksheet_names: tuple[str, ...] = (
        "CMKS",
        "QA Rekap",
        "QA Recap",
    )
    # Header source
    item_header: str = "Item"
    std_header: str = "STD"
    unit_header: str = "Unit"

    # Business/parser rules
    off_keyword: str = "OFF"
    total_keyword: str = "Total"

    # Jumlah maksimum kolom sampling time
    # pada satu blok.
    #
    # File contoh mempunyai beberapa kolom waktu di sebelah kanan Unit.
    max_sampling_columns: int = 4

    # Output
    output_sheet_name: str = "Data_Mining"
    control_sheet_name: str = "Control"

    output_headers: tuple[str, ...] = field(
        default=(
            "Tanggal",
            "Jam",
            "Sampling Point (SP)",
            "Parameter",
            "Unit",
            "Min",
            "Max",
            "Nilai",
        )
    )