from dataclasses import dataclass
from datetime import date
from typing import Any


@dataclass
class SamplingRecord:
    """
    Normalized QC sampling record.

    Model ini mempertahankan contract existing parser
    dan existing exporter.
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
        """
        Existing eight-column QC output representation.
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
    """
    Wrapper untuk mempertahankan source worksheet
    dari sebuah SamplingRecord.

    SamplingRecord tetap menjadi normalized QC record.
    Source worksheet merupakan provenance/traceability
    yang ditambahkan pada orchestration layer.

    Wrapper ini tidak mengubah isi SamplingRecord.
    """

    source_worksheet: str
    record: SamplingRecord

    def to_row(self) -> list:
        """
        Representasi row dengan source worksheet
        sebagai kolom provenance pertama.

        Method ini disiapkan untuk CP4 berikutnya.
        Existing exporter belum menggunakan method ini
        pada CP4.1.
        """

        return [
            self.source_worksheet,
            *self.record.to_row(),
        ]