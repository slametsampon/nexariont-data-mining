from dataclasses import dataclass
from datetime import date
from typing import Any


@dataclass
class SamplingRecord:
    """Satu hasil pengukuran QC dalam format delapan kolom.

    Semua atribut merupakan argumen konstruktor dataclass. Nilai pengukuran
    dan batas standar mempertahankan tipe hasil parser, termasuk teks khusus.

    Attributes:
        sampling_date (date): Tanggal pengambilan sampel.
        sampling_time (str): Waktu sampling yang telah diformat parser.
        sampling_point (str): Identitas titik sampling.
        parameter (str): Nama parameter pengujian.
        unit (Any): Satuan pengukuran.
        minimum (Any): Batas minimum standar.
        maximum (Any): Batas maksimum standar.
        value (Any): Nilai hasil pengukuran.

    Examples:
        >>> from datetime import date
        >>> record = SamplingRecord(
        ...     date(2026, 9, 22), "08:00", "SP-01", "pH", "-", 6, 9, 7.2
        ... )
        >>> record.to_row()[-1]
        7.2
    """

    sampling_date: date
    sampling_time: str
    sampling_point: str
    parameter: str
    unit: Any
    minimum: Any
    maximum: Any
    value: Any

    def to_row(self) -> list:
        """Mengubah record menjadi delapan nilai sesuai urutan header ekspor.

        Returns:
            list: Tanggal, waktu, titik sampling, parameter, satuan, minimum,
            maksimum, dan nilai pengukuran.
        """

        return [
            self.sampling_date,
            self.sampling_time,
            self.sampling_point,
            self.parameter,
            self.unit,
            self.minimum,
            self.maximum,
            self.value,
        ]


@dataclass(frozen=True)
class SourcedSamplingRecord:
    """Record QC beserta nama worksheet asal untuk penelusuran data.

    Wrapper frozen ini tidak mengubah isi SamplingRecord. Record yang dibungkus
    tetap mutable; frozen hanya membatasi penggantian atribut wrapper.

    Attributes:
        source_worksheet (str): Nama worksheet sumber; argumen konstruktor.
        record (SamplingRecord): Record yang dibungkus; argumen konstruktor.

    Examples:
        >>> from datetime import date
        >>> record = SamplingRecord(
        ...     date(2026, 9, 22), "08:00", "SP-01", "pH", "-", 6, 9, 7.2
        ... )
        >>> SourcedSamplingRecord("shift-pagi", record).to_row()[0]
        'shift-pagi'
    """

    source_worksheet: str
    record: SamplingRecord

    def to_row(self) -> list:
        """Menambahkan nama worksheet di depan delapan nilai record.

        Returns:
            list: Sembilan nilai dengan source_worksheet sebagai kolom pertama.
        """

        return [
            self.source_worksheet,
            *self.record.to_row(),
        ]