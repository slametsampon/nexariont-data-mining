from dataclasses import dataclass
from datetime import date
from typing import Any


@dataclass
class SamplingRecord:
    sampling_date: date
    sampling_time: str
    sampling_point: str
    parameter: str
    unit: Any
    minimum: Any
    maximum: Any
    value: Any

    def to_row(self) -> list:
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