from dataclasses import dataclass, field


@dataclass(frozen=True)
class MiningConfig:
    worksheet_name: str = "shift-pagi"

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