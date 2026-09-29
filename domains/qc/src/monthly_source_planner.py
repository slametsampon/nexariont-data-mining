from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


DOMAIN_ROUTING = (
    ("NPG", "NPG"),
    ("Octanol", "Octanol"),
    ("Syn Gas", "Syngas"),
    ("Utility", "Utility"),
    ("WWT", "wwt"),
)

ROMAN_WEEK = {
    "I": 1,
    "II": 2,
    "III": 3,
    "IV": 4,
    "V": 5,
}


def _norm_name(value: str) -> str:
    return " ".join(str(value).strip().casefold().split())


@dataclass(frozen=True)
class MonthlyWorkItem:
    month: str
    week: str
    folder_name: str
    domain: str
    source_workbook: str
    source_path: Path

    @property
    def provenance(self) -> tuple[str, str, str, str]:
        return (
            self.month,
            self.week,
            self.domain,
            self.source_workbook,
        )


@dataclass(frozen=True)
class MonthlyDiscoveryCondition:
    month: str
    week: str
    folder_name: str
    domain: str
    source_path: Path | None
    status: str
    message: str


@dataclass(frozen=True)
class MonthlyDomainPlan:
    month: str
    week: str
    folder_name: str
    domain: str
    domain_dir: Path | None
    work_items: tuple[MonthlyWorkItem, ...]
    condition: MonthlyDiscoveryCondition | None = None


@dataclass(frozen=True)
class MonthlyPlan:
    month: str
    weeks: tuple[Path, ...]
    domain_plans: tuple[MonthlyDomainPlan, ...]

    @property
    def work_items(self) -> tuple[MonthlyWorkItem, ...]:
        return tuple(
            item
            for domain_plan in self.domain_plans
            for item in domain_plan.work_items
        )

    @property
    def discovery_conditions(self) -> tuple[MonthlyDiscoveryCondition, ...]:
        return tuple(
            domain_plan.condition
            for domain_plan in self.domain_plans
            if domain_plan.condition is not None
        )


class MonthlySourcePlanner:
    """Translate one monthly source folder into a deterministic work plan."""

    def __init__(self, domain_routing=DOMAIN_ROUTING):
        self.domain_routing = tuple(domain_routing)

    def plan(self, input_root: Path) -> MonthlyPlan:
        input_root = Path(input_root)
        weeks = tuple(self._find_week_dirs(input_root))
        domain_plans: list[MonthlyDomainPlan] = []

        for week in weeks:
            for folder_name, parser_domain in self.domain_routing:
                domain_dir = self._find_child_dir(week, folder_name)

                if domain_dir is None:
                    condition = MonthlyDiscoveryCondition(
                        month=input_root.name,
                        week=week.name,
                        folder_name=folder_name,
                        domain=parser_domain,
                        source_path=None,
                        status="MISSING_DOMAIN_FOLDER",
                        message=f"Folder '{folder_name}' tidak ditemukan.",
                    )
                    domain_plans.append(
                        MonthlyDomainPlan(
                            month=input_root.name,
                            week=week.name,
                            folder_name=folder_name,
                            domain=parser_domain,
                            domain_dir=None,
                            work_items=(),
                            condition=condition,
                        )
                    )
                    continue

                source_files = self._find_source_workbooks(domain_dir)

                if not source_files:
                    condition = MonthlyDiscoveryCondition(
                        month=input_root.name,
                        week=week.name,
                        folder_name=folder_name,
                        domain=parser_domain,
                        source_path=domain_dir,
                        status="NO_WORKBOOKS",
                        message="Tidak ada source .xls/.xlsx.",
                    )
                    domain_plans.append(
                        MonthlyDomainPlan(
                            month=input_root.name,
                            week=week.name,
                            folder_name=folder_name,
                            domain=parser_domain,
                            domain_dir=domain_dir,
                            work_items=(),
                            condition=condition,
                        )
                    )
                    continue

                work_items = tuple(
                    MonthlyWorkItem(
                        month=input_root.name,
                        week=week.name,
                        folder_name=folder_name,
                        domain=parser_domain,
                        source_workbook=source.name,
                        source_path=source,
                    )
                    for source in source_files
                )

                domain_plans.append(
                    MonthlyDomainPlan(
                        month=input_root.name,
                        week=week.name,
                        folder_name=folder_name,
                        domain=parser_domain,
                        domain_dir=domain_dir,
                        work_items=work_items,
                    )
                )

        return MonthlyPlan(
            month=input_root.name,
            weeks=weeks,
            domain_plans=tuple(domain_plans),
        )

    @staticmethod
    def _week_sort_key(path: Path):
        name = path.name.strip()
        match = re.search(r"(?i)\bminggu\s+([ivx]+|\d+)\b", name)
        if match:
            token = match.group(1).upper()
            if token.isdigit():
                return (0, int(token), _norm_name(name))
            if token in ROMAN_WEEK:
                return (0, ROMAN_WEEK[token], _norm_name(name))
        return (1, 999, _norm_name(name))

    @classmethod
    def _find_week_dirs(cls, root: Path) -> list[Path]:
        weeks = [
            path
            for path in root.iterdir()
            if path.is_dir() and _norm_name(path.name).startswith("minggu ")
        ]
        return sorted(weeks, key=cls._week_sort_key)

    @staticmethod
    def _find_child_dir(parent: Path, expected_name: str) -> Path | None:
        expected = _norm_name(expected_name)
        for path in parent.iterdir():
            if path.is_dir() and _norm_name(path.name) == expected:
                return path
        return None

    @staticmethod
    def _find_source_workbooks(domain_dir: Path) -> list[Path]:
        files = []
        for path in domain_dir.rglob("*"):
            if not path.is_file():
                continue
            if path.name.startswith("~$"):
                continue
            if path.suffix.casefold() not in {".xls", ".xlsx"}:
                continue
            files.append(path)

        return sorted(
            files,
            key=lambda path: _norm_name(str(path.relative_to(domain_dir))),
        )
