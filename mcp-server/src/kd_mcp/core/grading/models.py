"""Grading data model shared by excel_grading / word_grading / powerpoint_grading.

Nothing here knows about a specific exam: each domain package builds
ItemScore lists from its own rubric and scoring rules.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path

PASS_SCORE = 60
SUMMARY_LIMIT = 150


@dataclass
class ItemScore:
    section: str          # 大項, e.g. "考題1.薪資"
    code: str             # 小項代號, e.g. "1.1"
    title: str
    score: float
    max_score: float
    note: str             # 評閱結果說明
    needs_review: bool = False
    is_bonus: bool = False  # 額外加分項：沒拿到不算「待改善」

    @property
    def completion_ratio(self) -> float:
        return round(self.score / self.max_score, 4) if self.max_score else 0.0

    def to_dict(self) -> dict:
        return {**asdict(self), "completion_ratio": self.completion_ratio}


@dataclass
class StudentInfo:
    dept: str | None
    emp_id: str | None
    name: str | None
    filename_valid: bool
    filename: str


@dataclass
class GradeResult:
    student: StudentInfo
    items: list[ItemScore]
    warnings: list[str] = field(default_factory=list)
    pass_score: float = PASS_SCORE

    @property
    def total(self) -> float:
        return round(sum(i.score for i in self.items), 2)

    @property
    def passed(self) -> bool:
        return self.total >= self.pass_score

    def section_totals(self) -> dict[str, dict]:
        out: dict[str, dict] = {}
        for i in self.items:
            s = out.setdefault(i.section, {"score": 0.0, "max_score": 0.0})
            s["score"] += i.score
            s["max_score"] += i.max_score
        return out

    def to_dict(self) -> dict:
        return {
            "student": asdict(self.student),
            "total": self.total,
            "passed": self.passed,
            "pass_score": self.pass_score,
            "section_totals": self.section_totals(),
            "items": [i.to_dict() for i in self.items],
            "needs_review": [i.code for i in self.items if i.needs_review],
            "warnings": self.warnings,
            "draft_summary": draft_summary(self),
        }


def parse_filename(path: str | Path) -> StudentInfo:
    """「部門代號-工號-姓名.<ext>」 -> StudentInfo. Anything that does not split
    into exactly three non-empty parts is reported invalid (never guessed)."""
    p = Path(path)
    parts = [s.strip() for s in p.stem.split("-")]
    if len(parts) == 3 and all(parts):
        return StudentInfo(parts[0], parts[1], parts[2], True, p.name)
    return StudentInfo(None, None, None, False, p.name)


def apply_overrides(result: GradeResult, overrides: dict | None) -> GradeResult:
    """overrides: {"3.4": {"score": 9, "note": "..."}} — the model's or a
    human's final call on a needs_review item. Clears needs_review."""
    for code, ov in (overrides or {}).items():
        item = next((i for i in result.items if i.code == code), None)
        if item is None:
            raise ValueError(f"Unknown item code: {code}")
        if "score" in ov:
            item.score = max(0.0, min(float(ov["score"]), item.max_score))
        if ov.get("note"):
            item.note = ov["note"]
        item.needs_review = False
    return result


def draft_summary(result: GradeResult, limit: int = SUMMARY_LIMIT) -> str:
    """A factual starting point for 評閱結果簡述 (≤150 chars): strengths first,
    then what to improve. The model is expected to polish it."""
    full = [i for i in result.items if i.max_score and i.score >= i.max_score]
    # only base-point items count as 待改善; a missed bonus is not a fault
    weak = [i for i in result.items if i.score < i.max_score and not i.is_bonus]
    sections = result.section_totals()
    good_sections = [s.split(".", 1)[-1] for s, v in sections.items() if v["score"] >= v["max_score"]]
    strengths = []
    if good_sections:
        strengths.append(f"{'、'.join(good_sections)}全部完成")
    bonus = [i.title.replace("(額外加分)", "") for i in full if i.is_bonus]
    if bonus:
        strengths.append(f"取得{'、'.join(bonus)}加分")
    text = f"總分{result.total:g}分，{'合格' if result.passed else '不合格'}。"
    if strengths:
        text += "，".join(strengths) + "。"
    if weak:
        text += "待改善：" + "、".join(f"{i.code}{i.title}" for i in weak) + "。"
    return text if len(text) <= limit else text[: limit - 1] + "…"
