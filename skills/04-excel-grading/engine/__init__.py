"""xlsx_s06 — S06 Excel實作 評分 engine, kept with its skill (skills/04-excel-grading/engine).

Loaded by kd_mcp.core.grading.engines from skill.toml's [grading] engine_path.
Needs mcp-server/src on sys.path for kd_mcp's grading models (no MCP runtime).

extract_submission: extract objective facts from a submission (also a CLI)
extract_rubric    : built-in point table + reader for 評閱標準.xlsx
score_submission  : references/scoring_checks.md rules -> GradeResult
results           : writes the 單科 評分結果 workbook (評分細項 + 評分結果表)
"""
from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from kd_mcp.core.grading.models import GradeResult, StudentInfo
from .extract_submission import inspect_submission
from .extract_rubric import compare_with_builtin, read_rubric_file
from . import results
from .score_submission import score_facts

if TYPE_CHECKING:
    from kd_mcp.core.registry import SkillSpec


class XlsxS06Engine:
    name = "xlsx_s06"
    formats = ("xlsx",)
    results = results    # Meta, default_output_name, meta_from_existing, write_results

    def inspect(self, path: Path, spec: "SkillSpec") -> dict:
        # the question file lets the engine detect altered raw data (3.x / 4.x)
        return inspect_submission(path, spec.asset("question"))

    def score(self, facts: dict, student: StudentInfo, spec: "SkillSpec") -> GradeResult:
        result = score_facts(facts, student)
        result.pass_score = spec.pass_score
        return result

    def rubric(self, spec: "SkillSpec", rubric_path: Path | None = None) -> dict:
        path = rubric_path or spec.asset("rubric")
        if path is None:
            raise FileNotFoundError(f"{spec.id} 的 skill.toml 未設定 [assets] rubric")
        items = read_rubric_file(path)
        return {
            "rubric_path": str(path),
            "items": items,
            "base_total": sum(i["points"] for i in items),
            "max_total_with_bonus": sum(i["points"] + i["bonus"] for i in items),
            "diffs_from_scorer": compare_with_builtin(items),
            "pass_score": spec.pass_score,
        }


ENGINE = XlsxS06Engine()
