"""xlsx_s06 — S06 Excel實作 評分 engine (skills/04-excel-grading).

inspection: extract objective facts from a submission (port of inspect_submission.py)
rubric    : built-in point table + reader for 評閱標準.xlsx
scoring   : references/scoring_checks.md rules -> GradeResult
"""
from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from ...models import GradeResult, StudentInfo
from .inspection import inspect_submission
from .rubric import compare_with_builtin, read_rubric_file
from .scoring import score_facts

if TYPE_CHECKING:
    from ....registry import SkillSpec


class XlsxS06Engine:
    name = "xlsx_s06"
    formats = ("xlsx",)

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
