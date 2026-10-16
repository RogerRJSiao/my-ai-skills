"""What a grading engine must provide. One engine can serve several skills
(the skill-specific bits — asset files, pass score — come from SkillSpec)."""
from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Protocol

from ..models import GradeResult, StudentInfo

if TYPE_CHECKING:
    from ...registry import SkillSpec


class Engine(Protocol):
    name: str
    formats: tuple[str, ...]          # file extensions it can read, e.g. ("xlsx",)

    def inspect(self, path: Path, spec: "SkillSpec") -> dict:
        """Objective facts about one submission. Never computes scores."""

    def score(self, facts: dict, student: StudentInfo, spec: "SkillSpec") -> GradeResult:
        """Apply the skill's scoring rules to `facts`."""

    def rubric(self, spec: "SkillSpec", rubric_path: Path | None = None) -> dict:
        """The rubric as data, plus differences from the engine's built-in table."""
