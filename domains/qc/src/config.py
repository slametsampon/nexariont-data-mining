from dataclasses import dataclass, field


@dataclass(frozen=True)
class MiningConfig:
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