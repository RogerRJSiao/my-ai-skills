"""Write graded students into the 單科 評分結果 workbook (活頁1 評分細項 + 活頁2 評分結果表).

Part of the xlsx_s06 engine: the MCP toolkit reaches it as ENGINE.results.
The template path is passed in, and only the two sheets' fixed headers are
assumed. Each skill writes its own
workbook (one format per file); 活頁2 副檔名 tells the formats apart when the
three subjects are merged later.

Follows SKILL.md 第3/4步:
- append after the existing rows, never overwrite other students;
- 分數加總 / 是否合格 are formulas, and every 活頁2 row's range is rewritten
  so it covers the new 活頁1 last row.

分數加總 uses SUMIFS on 工號 + 姓名 (not SUMIF on 姓名 alone) because two
students can share a name, and the known cases already share a 工號.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from kd_mcp.core.grading.models import PASS_SCORE, SUMMARY_LIMIT, GradeResult

DETAIL_SHEET = "評分細項"
SUMMARY_SHEET = "評分結果表"
DETAIL_HEADERS = ["截止日", "課程代碼", "實作項目", "部門代號", "工號", "姓名",
                  "大項", "小項代號", "得分", "完成比例", "執行者", "評閱結果說明"]
SUMMARY_HEADERS = ["截止日", "課程代碼", "實作項目", "部門代號", "工號", "姓名", "副檔名",
                   "檔案數", "分數加總", "是否合格", "執行者", "評閱結果簡述(150字以內)"]


@dataclass
class Meta:
    deadline: str | None
    course_code: str | None
    practice_item: str | None
    grader: str | None

    def missing(self) -> list[str]:
        labels = {"deadline": "截止日", "course_code": "課程代碼",
                  "practice_item": "實作項目", "grader": "執行者"}
        return [labels[k] for k, v in self.__dict__.items() if not v]


class DuplicateStudentError(ValueError):
    pass


def _deadline_value(s):
    if isinstance(s, (dt.date, dt.datetime)):
        return s
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return dt.datetime.strptime(str(s), fmt)
        except ValueError:
            pass
    return s


def _last_row(ws) -> int:
    """Last row with anything in column A..F (ws.max_row counts styled blanks)."""
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, 7)):
            return r
    return 1


def meta_from_existing(wb) -> Meta | None:
    ws = wb[SUMMARY_SHEET]
    if _last_row(ws) < 2:
        return None
    v = [ws.cell(row=2, column=c).value for c in (1, 2, 3, 11)]
    deadline = v[0].strftime("%Y-%m-%d") if isinstance(v[0], (dt.date, dt.datetime)) else v[0]
    return Meta(deadline, v[1], v[2], v[3])


def _check_headers(wb):
    for name, headers in ((DETAIL_SHEET, DETAIL_HEADERS), (SUMMARY_SHEET, SUMMARY_HEADERS)):
        if name not in wb.sheetnames:
            raise ValueError(f"結果活頁簿缺少工作表「{name}」")
        actual = [wb[name].cell(row=1, column=c).value for c in range(1, len(headers) + 1)]
        if actual != headers:
            raise ValueError(f"「{name}」表頭與標準格式不同：{actual}")


def _student_rows(ws, emp_id, name) -> list[int]:
    return [
        r for r in range(2, _last_row(ws) + 1)
        if str(ws.cell(row=r, column=5).value) == str(emp_id) and ws.cell(row=r, column=6).value == name
    ]


def _copy_style(ws, src_row: int, dst_row: int, ncols: int):
    for c in range(1, ncols + 1):
        src, dst = ws.cell(row=src_row, column=c), ws.cell(row=dst_row, column=c)
        if src.has_style:
            dst._style = src._style


SUBJECT_LABEL = {"xlsx": "Excel", "docx": "Word", "pptx": "PPT"}


def default_output_name(meta: Meta, ext: str, batch: str | None = None,
                        now: dt.datetime | None = None) -> str:
    """<課程代碼>_<Excel|Word|PPT>_評分結果<yyyymm>_<yyyymmdd_HHMMSS>.xlsx.
    yyyymm is the batch (梯次); without it, the 截止日's year and month."""
    now = now or dt.datetime.now()
    if not batch:
        deadline = _deadline_value(meta.deadline)
        batch = f"{deadline:%Y%m}" if isinstance(deadline, (dt.date, dt.datetime)) else f"{now:%Y%m}"
    return (f"{meta.course_code}_{SUBJECT_LABEL.get(ext, ext)}_評分結果{batch}"
            f"_{now:%Y%m%d_%H%M%S}.xlsx")


def write_results(
    base_workbook: Path,
    output_path: Path,
    graded: list[tuple[GradeResult, str]],
    meta: Meta,
    on_duplicate: str = "error",
    pass_score: float = PASS_SCORE,
    ext: str | None = None,
) -> dict:
    """graded: [(GradeResult, 評閱結果簡述)]. on_duplicate: error | replace | append.
    ext: the skill's format, written to 活頁2 副檔名 (e.g. "xlsx")."""
    if on_duplicate not in ("error", "replace", "append"):
        raise ValueError("on_duplicate must be error, replace or append")

    wb = openpyxl.load_workbook(base_workbook)
    _check_headers(wb)
    ws_d, ws_s = wb[DETAIL_SHEET], wb[SUMMARY_SHEET]

    replaced = []
    for result, _ in graded:
        st = result.student
        rows_s = _student_rows(ws_s, st.emp_id, st.name)
        if not rows_s:
            continue
        if on_duplicate == "error":
            raise DuplicateStudentError(
                f"{st.emp_id}-{st.name} 已存在於結果檔，請改用 on_duplicate='replace' 或 'append'"
            )
        if on_duplicate == "replace":
            for ws in (ws_d, ws_s):
                for r in reversed(_student_rows(ws, st.emp_id, st.name)):
                    ws.delete_rows(r)
            replaced.append(f"{st.emp_id}-{st.name}")

    deadline = _deadline_value(meta.deadline)
    head = [deadline, meta.course_code, meta.practice_item]
    style_row_d = 2 if _last_row(ws_d) >= 2 else None
    style_row_s = 2 if _last_row(ws_s) >= 2 else None

    for result, summary in graded:
        st = result.student
        for item in result.items:
            r = _last_row(ws_d) + 1
            values = head + [st.dept, st.emp_id, st.name, item.section, item.code,
                             item.score, item.completion_ratio, meta.grader, item.note]
            for c, v in enumerate(values, start=1):
                ws_d.cell(row=r, column=c, value=v)
            if style_row_d:
                _copy_style(ws_d, style_row_d, r, len(DETAIL_HEADERS))
            ws_d.cell(row=r, column=10).number_format = "0%"
            if isinstance(deadline, dt.datetime):
                ws_d.cell(row=r, column=1).number_format = "yyyy/mm/dd"

        r = _last_row(ws_s) + 1
        values = head + [st.dept, st.emp_id, st.name, ext, 1, None, None, meta.grader,
                         (summary or "")[:SUMMARY_LIMIT]]
        for c, v in enumerate(values, start=1):
            ws_s.cell(row=r, column=c, value=v)
        if style_row_s:
            _copy_style(ws_s, style_row_s, r, len(SUMMARY_HEADERS))
        if isinstance(deadline, dt.datetime):
            ws_s.cell(row=r, column=1).number_format = "yyyy/mm/dd"

    # Rewrite every 活頁2 formula so its range covers the new 活頁1 last row.
    last_d = max(_last_row(ws_d), 2)
    rng = lambda col: f"{DETAIL_SHEET}!${col}$2:${col}${last_d}"  # noqa: E731
    for r in range(2, _last_row(ws_s) + 1):
        ws_s.cell(row=r, column=9, value=(
            f"=SUMIFS({rng('I')},{rng('E')},E{r},{rng('F')},F{r})"
        ))
        ws_s.cell(row=r, column=10, value=f'=IF(I{r}>={pass_score:g},"合格","不合格")')

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return {
        "output_path": str(output_path),
        "detail_rows": _last_row(ws_d) - 1,
        "summary_rows": _last_row(ws_s) - 1,
        "replaced": replaced,
        "students": [
            {
                "dept": res.student.dept,
                "emp_id": res.student.emp_id,
                "name": res.student.name,
                "total": res.total,
                "passed": res.passed,
            }
            for res, _ in graded
        ],
        "note": "分數加總/是否合格為公式，Excel 開啟時才會計算；上方 total/passed 為伺服器端核對值。",
    }
