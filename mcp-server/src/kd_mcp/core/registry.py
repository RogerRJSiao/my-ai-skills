"""Discover skills from skills/*/skill.toml and filter them by domain.

skill.toml layout:

    [skill]   id, title, domain, format, course_code?, practice_item?
    [grading] engine, pass_score?
    [assets]  name = "relative/path" ...   (relative to the skill folder)
"""
from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from pathlib import Path

MANIFEST = "skill.toml"
REQUIRED_SKILL_KEYS = ("id", "title", "domain", "format")


@dataclass(frozen=True)
class SkillSpec:
    id: str
    title: str
    domain: str
    formats: tuple[str, ...]     # "pbix|pbip" -> ("pbix", "pbip")
    root: Path
    engine: str
    pass_score: float = 60
    course_code: str | None = None
    practice_item: str | None = None
    assets: dict[str, Path] = field(default_factory=dict)

    def asset(self, key: str) -> Path | None:
        return self.assets.get(key)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "title": self.title,
            "domain": self.domain,
            "formats": list(self.formats),
            "engine": self.engine,
            "pass_score": self.pass_score,
            "course_code": self.course_code,
            "practice_item": self.practice_item,
        }


@dataclass(frozen=True)
class RegistryIssue:
    skill_dir: str
    domain: str | None
    message: str


def _parse(manifest: Path, known_engines: set[str] | None) -> tuple[SkillSpec | None, list[str], str | None]:
    try:
        data = tomllib.loads(manifest.read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError) as e:
        return None, [f"無法讀取 {MANIFEST}：{e}"], None

    skill = data.get("skill", {})
    grading = data.get("grading", {})
    domain = skill.get("domain")
    problems = [f"[skill] 缺少 {k}" for k in REQUIRED_SKILL_KEYS if not skill.get(k)]
    if not grading.get("engine"):
        problems.append("[grading] 缺少 engine")
    elif known_engines is not None and grading["engine"] not in known_engines:
        problems.append(f"未知的 engine「{grading['engine']}」")
    if skill.get("id") and skill["id"] != manifest.parent.name:
        problems.append(f"id「{skill['id']}」與資料夾名稱「{manifest.parent.name}」不同")

    root = manifest.parent.resolve()
    assets = {}
    for key, rel in data.get("assets", {}).items():
        p = (root / rel).resolve()
        if not p.is_relative_to(root):
            problems.append(f"asset {key} 指到 skill 資料夾以外：{rel}")
        elif not p.exists():
            problems.append(f"asset {key} 不存在：{rel}")
        else:
            assets[key] = p
    if problems:
        return None, problems, domain

    spec = SkillSpec(
        id=skill["id"],
        title=skill["title"],
        domain=domain,
        formats=tuple(f.strip().lower() for f in str(skill["format"]).split("|") if f.strip()),
        root=root,
        engine=grading["engine"],
        pass_score=float(grading.get("pass_score", 60)),
        course_code=skill.get("course_code"),
        practice_item=skill.get("practice_item"),
        assets=assets,
    )
    return spec, [], domain


def scan(skills_dir: Path, known_engines: set[str] | None = None) -> tuple[list[SkillSpec], list[RegistryIssue]]:
    """Parse every skills/*/skill.toml. Folders without a manifest are plain
    skills (no MCP support) and are skipped silently."""
    specs, issues = [], []
    if not skills_dir.is_dir():
        return specs, [RegistryIssue(str(skills_dir), None, "skills 目錄不存在")]
    for manifest in sorted(skills_dir.glob(f"*/{MANIFEST}")):
        spec, problems, domain = _parse(manifest, known_engines)
        if spec:
            specs.append(spec)
        issues += [RegistryIssue(manifest.parent.name, domain, m) for m in problems]
    return specs, issues


def for_domain(specs: list[SkillSpec], domain: str) -> dict[str, SkillSpec]:
    return {s.id: s for s in specs if s.domain == domain}
