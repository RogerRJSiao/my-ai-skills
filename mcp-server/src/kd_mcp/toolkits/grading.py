"""Grading toolkit: one set of grading tools that serves every skill of a
domain (training_internal = 內訓, training_external = 外訓, ...).

Which skills exist, which file formats they take and which engine grades
them all come from skills/*/skill.toml (core.registry); this module only
exposes them over MCP. Mount it with a prefix per server:

    register(mcp, domain="training_internal", prefix="gi", data_root=..., output=...)
    -> gi_list_skills, gi_get_rubric, gi_list_submissions,
       gi_inspect_submission, gi_grade_submission, gi_write_results
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Annotated, Literal

from mcp.server.mcpserver import MCPServer
from pydantic import BaseModel, Field

from ..core.config import default_skills_dir
from ..core.grading import apply_overrides, parse_filename
from ..core.grading.engines import engine_status, get_engine
from ..core.logging import AuditLogger
from ..core.registry import for_domain, scan
from ..core.security import PathPolicy
from ..runtime.tooling import audited


class ItemOverride(BaseModel):
    score: float | None = Field(None, description="Final score for this item (clamped to 0..max).")
    note: str | None = Field(None, description="Replacement 評閱結果說明 for this item.")


class StudentInput(BaseModel):
    submission_path: str = Field(description="Path to the student's 部門代號-工號-姓名.<ext> file.")
    summary: str | None = Field(
        None, description="評閱結果簡述 (≤150 chars; ~100 strengths + ~50 to improve). "
                          "Omit to use the server's draft_summary."
    )
    overrides: dict[str, ItemOverride] | None = Field(
        None, description='Per-item final calls, e.g. {"3.4": {"score": 9, "note": "..."}}.'
    )


SkillArg = Annotated[str, Field(description="Skill id from <prefix>_list_skills, e.g. 04-excel-grading.")]


def instructions(prefix, domain):
    return f"""\
Grading tools for the {domain} skills of the my-ai-skills repo.

Start with {prefix}_list_skills to see which skills this server grades, then
{prefix}_list_submissions and {prefix}_grade_submission per file. Before
{prefix}_write_results, ask the user once for 截止日 and 執行者 (and confirm
課程代碼/實作項目, which default to the skill's skill.toml) unless appending
to an existing result file; confirm any file whose filename_valid is false;
resolve every needs_review item and pass the decision as overrides.
Paths are relative to the server's data root. {prefix}_write_results saves
into the batch's 03_results/ when the batch folder exists (pass batch, e.g.
202612), otherwise into the server's output folder; nothing else is written.
"""


def _overrides(raw):
    return {k: v.model_dump(exclude_none=True) for k, v in (raw or {}).items()}


def register(mcp, *, domain, prefix, data_root, output, skills_dir=None, extra_read_dirs=(), audit=None):
    skills_dir = Path(skills_dir or default_skills_dir()).resolve()
    policy = PathPolicy.build(data_root=data_root, output=output, skills_dir=skills_dir,
                              extra_read_dirs=extra_read_dirs)
    all_specs, all_issues = scan(skills_dir)
    specs = for_domain(all_specs, domain)
    issues = [i for i in all_issues if i.domain in (domain, None)]

    def spec_of(skill):
        spec = specs.get(skill)
        if spec is None:
            available = "、".join(specs) or "（無）"
            raise ValueError(f"skill「{skill}」不屬於本 server（domain={domain}）。可用的 skill：{available}")
        return spec

    def grade(spec, submission_path, overrides):
        engine = get_engine(spec.engine, spec.engine_path)
        sub = policy.read(submission_path)
        if sub.suffix.lower().lstrip(".") not in spec.formats:
            raise ValueError(f"{sub.name} 不是 {spec.id} 接受的格式（{'/'.join(spec.formats)}）")
        result = engine.score(engine.inspect(sub, spec), parse_filename(sub), spec)
        return apply_overrides(result, overrides)

    #--把函式登記成 MCP tool，加上前綴
    def tool(name):
        full = f"{prefix}_{name}"

        def deco(fn):
            return mcp.tool(name=full)(audited(full, audit)(fn))
        return deco

    @tool("list_skills")
    def list_skills() -> dict:
        """Skills this server can grade (from skills/*/skill.toml), with their
        accepted file formats, default 課程代碼/實作項目 and whether the grading
        engine is installed. `issues` lists skill.toml problems to fix."""
        out = []
        for spec in specs.values():
            ok, reason = engine_status(spec.engine, spec.engine_path)
            out.append({**spec.to_dict(), "engine_available": ok, "engine_error": reason})
        return {
            "domain": domain,
            "skills": out,
            "issues": [{"skill_dir": i.skill_dir, "message": i.message} for i in issues],
        }

    @tool("get_rubric")
    def get_rubric(
        skill: SkillArg,
        rubric_path: Annotated[str | None, Field(
            description="評閱標準 file to read instead of the skill's bundled one.")] = None,
    ) -> dict:
        """Read the skill's 評閱標準 and report point differences from the
        engine's built-in table. If `diffs_from_scorer` is non-empty, tell the
        user the rubric changed (scores still follow the built-in table)."""
        spec = spec_of(skill)
        return get_engine(spec.engine, spec.engine_path).rubric(spec, policy.read(rubric_path) if rubric_path else None)

    @tool("list_submissions")
    def list_submissions(
        skill: SkillArg,
        folder: Annotated[str, Field(description="Folder containing student files.")],
    ) -> dict:
        """List the skill's submission files in a folder and parse
        部門代號/工號/姓名 from each filename. Files with filename_valid=false must
        be confirmed with the user before grading (never guess the fields)."""
        spec = spec_of(skill)
        root = policy.read(folder)
        if not root.is_dir():
            raise NotADirectoryError(str(root))
        files = sorted(
            p for p in root.iterdir()
            if p.is_file() and p.suffix.lower().lstrip(".") in spec.formats and not p.name.startswith("~$")
        )
        return {
            "skill": spec.id,
            "folder": str(root),
            "submissions": [{"path": str(p), **parse_filename(p).__dict__} for p in files],
        }

    @tool("inspect_submission")
    def inspect_submission(
        skill: SkillArg,
        submission_path: Annotated[str, Field(description="Student file.")],
    ) -> dict:
        """Raw facts the engine extracts from one submission (for 04: pivot
        table, charts, freeze panes, borders, sort/filter, formulas, raw-data
        check). Use it to double-check an item; grade_submission calls it."""
        spec = spec_of(skill)
        return get_engine(spec.engine, spec.engine_path).inspect(policy.read(submission_path), spec)

    @tool("grade_submission")
    def grade_submission(
        skill: SkillArg,
        submission_path: Annotated[str, Field(description="Student file (部門代號-工號-姓名.<ext>).")],
        overrides: Annotated[dict[str, ItemOverride] | None, Field(
            description='Optional per-item final calls, e.g. {"3.4": {"score": 9}}.')] = None,
    ) -> dict:
        """Score one submission with the skill's engine.

        Returns every item (score, max_score, completion_ratio, 評閱結果說明 note),
        section totals, total, passed, warnings, a draft 評閱結果簡述, and
        `needs_review`: item codes the SOP says must be confirmed by opening
        the file. Resolve those before writing results. Writes nothing."""
        return grade(spec_of(skill), submission_path, _overrides(overrides)).to_dict()

    @tool("write_results")
    def write_results(
        skill: SkillArg,
        students: Annotated[list[StudentInput], Field(description="Students to grade and write.")],
        deadline: Annotated[str | None, Field(description="截止日, e.g. 2026-10-31.")] = None,
        course_code: Annotated[str | None, Field(description="課程代碼. Default: skill.toml.")] = None,
        practice_item: Annotated[str | None, Field(description="實作項目. Default: skill.toml.")] = None,
        grader: Annotated[str | None, Field(description="執行者 (評閱者姓名或代號).")] = None,
        result_path: Annotated[str | None, Field(
            description="Existing 評分結果.xlsx to append to. Default: the skill's blank template.")] = None,
        output_path: Annotated[str | None, Field(
            description="Where to save (in the output folder or a batch's 03_results/). "
                        "Default: <課程代碼>_<梯次>/03_results/ when that batch folder exists, else <output>/; "
                        "file name <課程代碼>_<Excel|Word|PPT>_評分結果<梯次yyyymm>_<yyyymmdd_HHMMSS>.xlsx")] = None,
        batch: Annotated[str | None, Field(
            description="梯次 yyyymm, e.g. 202612: picks <課程代碼>_<梯次>/03_results/ and the file name. "
                        "Default: 截止日's year+month.")] = None,
        on_duplicate: Annotated[Literal["error", "replace", "append"], Field(
            description="What to do if a student (工號+姓名) is already in the result file.")] = "error",
        overwrite: Annotated[bool, Field(description="Allow replacing an existing output file.")] = False,
    ) -> dict:
        """Grade each student and write 活頁1 評分細項 (one row per item) and
        活頁2 評分結果表 (one row per student, with 副檔名) into the 單科 評分結果 workbook.
        分數加總 and 是否合格 are formulas covering the whole 明細 sheet."""
        import openpyxl

        spec = spec_of(skill)
        res = get_engine(spec.engine, spec.engine_path).results
        Meta, default_output_name, meta_from_existing = res.Meta, res.default_output_name, res.meta_from_existing
        _write = res.write_results
        if result_path:
            base = policy.read(result_path)
        else:
            base = spec.asset("result_template")
            if base is None:
                raise FileNotFoundError(f"{spec.id} 的 skill.toml 未設定 [assets] result_template")

        meta = Meta(deadline, course_code, practice_item, grader)
        sources = {k: "argument" for k, v in (("course_code", course_code),
                                                ("practice_item", practice_item)) if v}
        if meta.missing() and result_path:
            existing = meta_from_existing(openpyxl.load_workbook(base))
            if existing:
                for k in ("course_code", "practice_item"):
                    if not getattr(meta, k) and getattr(existing, k):
                        sources[k] = "existing result file"
                meta = Meta(*(a or b for a, b in zip(meta.__dict__.values(), existing.__dict__.values())))
        for k in ("course_code", "practice_item"):
            if not getattr(meta, k) and getattr(spec, k):
                setattr(meta, k, getattr(spec, k))
                sources[k] = "skill.toml"
        if meta.missing():
            raise ValueError(f"缺少必要資訊：{'、'.join(meta.missing())}，請先向使用者確認。")

        graded = []
        for s in students:
            result = grade(spec, s.submission_path, _overrides(s.overrides))
            if not result.student.filename_valid:
                raise ValueError(
                    f"檔名「{result.student.filename}」無法解析出部門代號/工號/姓名，請先向使用者確認或更正檔名。"
                )
            graded.append((result, s.summary or result.to_dict()["draft_summary"]))

        if output_path:
            out = policy.write(output_path)
        else:
            name = default_output_name(meta, spec.formats[0], batch)
            # the batch in the name (given, or from 截止日) picks <課程代碼>_<梯次>/03_results/
            m = re.search(r"評分結果(\d{6})_", name)
            results_dir = Path(data_root).resolve() / f"{meta.course_code}_{m.group(1)}" / "03_results" if m else None
            out = policy.write((results_dir if results_dir and results_dir.is_dir() else Path(output)) / name)
        if out.exists() and not overwrite:
            raise FileExistsError(f"{out} 已存在；若要覆寫請設定 overwrite=true，或指定其他 output_path。")

        summary = _write(base, out, graded, meta, on_duplicate, pass_score=spec.pass_score,
                         ext=spec.formats[0])
        summary["meta_sources"] = sources
        summary["unresolved_needs_review"] = {
            f"{r.student.emp_id}-{r.student.name}": [i.code for i in r.items if i.needs_review]
            for r, _ in graded if any(i.needs_review for i in r.items)
        }
        return summary
