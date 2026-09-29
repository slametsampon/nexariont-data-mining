from __future__ import annotations

import tempfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .monthly_source_planner import MonthlySourcePlanner
from .monthly_workbook_consolidator import MonthlyWorkbookConsolidator
from .workbook_processor import QCWorkbookProcessor


MASTER_SOURCE_SHEET = "Sampling-Point"
CORE_MASTER_SHEET = "QA Review"


class NoWeekFoldersFoundError(ValueError):
    pass


@dataclass(frozen=True)
class MonthlyRunResult:
    input_root: Path
    output_dir: Path
    master: Path
    final_output: Path
    month: str
    weeks_detected: int
    totals: Counter


class MonthlyMiningOrchestrator:
    """Coordinate the monthly QC mining application flow."""

    def __init__(
        self,
        planner: MonthlySourcePlanner | None = None,
        processor: QCWorkbookProcessor | None = None,
    ):
        self.planner = planner or MonthlySourcePlanner()
        self.processor = processor or QCWorkbookProcessor()

    def run(
        self,
        input_root: Path,
        output_dir: Path,
        master: Path,
        final_output: Path,
    ) -> MonthlyRunResult:
        input_root = Path(input_root)
        output_dir = Path(output_dir)
        master = Path(master)
        final_output = Path(final_output)

        plan = self.planner.plan(input_root)
        if not plan.weeks:
            raise NoWeekFoldersFoundError(
                f"NO_WEEK_FOLDERS_FOUND under: {input_root}"
            )

        consolidator = MonthlyWorkbookConsolidator()
        source_counter = 0

        with tempfile.TemporaryDirectory(prefix="nexariont_qc_monthly_") as tmp:
            tmp_dir = Path(tmp)
            adapted_master = self._build_core_master_adapter(
                master,
                tmp_dir / "_Master-Data_core_adapter.xlsx",
            )

            for domain_plan in plan.domain_plans:
                if domain_plan.condition is not None:
                    consolidator.record_discovery_condition(domain_plan.condition)
                    continue

                for item in domain_plan.work_items:
                    source_counter += 1
                    temp_output = tmp_dir / f"{source_counter:04d}.xlsx"

                    try:
                        self.processor.process(
                            input_file=item.source_path,
                            output_file=temp_output,
                            master_file=adapted_master,
                            domain=item.domain,
                        )
                    except Exception as exc:
                        consolidator.record_execution_failure(
                            item,
                            f"ERROR: {exc}",
                        )
                        continue

                    if not temp_output.exists():
                        consolidator.record_execution_failure(
                            item,
                            "Output file tidak terbentuk.",
                        )
                        continue

                    consolidator.consolidate_workbook(
                        item,
                        temp_output,
                    )

        totals = consolidator.finalize(final_output)

        return MonthlyRunResult(
            input_root=input_root,
            output_dir=output_dir,
            master=master,
            final_output=final_output,
            month=plan.month,
            weeks_detected=len(plan.weeks),
            totals=totals,
        )

    @staticmethod
    def _build_core_master_adapter(
        master_path: Path,
        adapter_path: Path,
    ) -> Path:
        wb_source = load_workbook(
            master_path,
            data_only=True,
            read_only=True,
        )
        try:
            if MASTER_SOURCE_SHEET not in wb_source.sheetnames:
                raise ValueError(
                    f"Worksheet master '{MASTER_SOURCE_SHEET}' tidak ditemukan "
                    f"dalam {master_path}"
                )

            ws_source = wb_source[MASTER_SOURCE_SHEET]

            wb_adapter = Workbook()
            ws_adapter = wb_adapter.active
            ws_adapter.title = CORE_MASTER_SHEET

            for row in ws_source.iter_rows(values_only=True):
                ws_adapter.append(list(row))

            wb_adapter.save(adapter_path)
            return adapter_path
        finally:
            wb_source.close()
