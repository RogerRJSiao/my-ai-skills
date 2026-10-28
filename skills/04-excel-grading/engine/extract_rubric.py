"""The S06 rubric: the built-in point table the scorer uses, and a reader for
評閱標準.xlsx so a newer rubric file can be compared against it."""
from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from pathlib import Path

import openpyxl

# Asset file names live in skills/04-excel-grading/skill.toml.


@dataclass(frozen=True)
class RubricItem:
    section: str       # 大項, e.g. "考題1.薪資"
    code: str          # 小項代號, e.g. "1.1"
    title: str         # 小項名稱, e.g. "作答狀態判斷"
    points: float      # 基本配分
    bonus: float       # 額外加分配分

    @property
    def max_score(self) -> float:
        return self.points + self.bonus

    def to_dict(self) -> dict:
        return {**asdict(self), "max_score": self.max_score}


# Mirrors the table at the top of references/scoring_checks.md.
BUILTIN_RUBRIC: tuple[RubricItem, ...] = (
    RubricItem("考題1.薪資", "1.1", "作答狀態初步判斷", 5, 0),
    RubricItem("考題1.薪資", "1.2", "欄列標籤選取", 10, 0),
    RubricItem("考題1.薪資", "1.3", "樞紐值的彙總方式", 10, 0),
    RubricItem("考題1.薪資", "1.4", "資料型別顯示設定", 5, 0),
    RubricItem("考題1.薪資", "1.5", "從表格取得資料集(額外加分)", 0, 6),
    RubricItem("考題2.成績", "2.1", "凍結窗格設定", 4, 0),
    RubricItem("考題2.成績", "2.2", "標頭跨欄置中設定", 4, 0),
    RubricItem("考題2.成績", "2.3", "Excel公式與表格範圍設定", 9, 0),
    RubricItem("考題2.成績", "2.4", "排序與篩選設定", 10, 0),
    RubricItem("考題2.成績", "2.5", "資料格線樣式設計", 3, 0),
    RubricItem("考題2.成績", "2.6", "處理有問題儲存格(額外加分)", 0, 3),
    RubricItem("考題3.車輛銷售", "3.1", "作答狀態初步判斷", 5, 0),
    RubricItem("考題3.車輛銷售", "3.2", "圖表座標軸設定", 15, 0),
    RubricItem("考題3.車輛銷售", "3.3", "圖表圖例設定", 5, 0),
    RubricItem("考題3.車輛銷售", "3.4", "圖表資訊美化(額外加分)", 0, 9),
    RubricItem("考題4.武將能力", "4.1", "作答狀態初步判斷", 5, 0),
    RubricItem("考題4.武將能力", "4.2", "圖表座標設定", 10, 0),
    RubricItem("考題4.武將能力", "4.3", "圖表資訊美化(額外加分)", 0, 6),
)


def builtin_rubric_map() -> dict[str, RubricItem]:
    return {item.code: item for item in BUILTIN_RUBRIC}


_CODE_RE = re.compile(r"^\s*(\d+\.\d+)\s*[-－]?\s*(.*?)\s*[。.]?\s*$")


def _num(v) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def read_rubric_file(path: str | Path) -> list[dict]:
    """Read 評閱標準.xlsx (first sheet, header on row 1).

    Expected columns: 大項 | 小項 | 評分標準 | 配分 | 額外加分配分 | ..."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None:
            continue
        m = _CODE_RE.match(str(row[1]))
        code, title = (m.group(1), m.group(2)) if m else (str(row[1]).strip(), "")
        items.append({
            "section": (row[0] or "").strip(),
            "code": code,
            "title": title,
            "criteria": row[2],
            "points": _num(row[3]),
            "bonus": _num(row[4]) if len(row) > 4 else 0.0,
        })
    wb.close()
    return items


def compare_with_builtin(file_items: list[dict]) -> list[dict]:
    """List differences in points/bonus/item set between a rubric file and
    the built-in table. An empty list means the scorer is in sync."""
    builtin = builtin_rubric_map()
    diffs = []
    seen = set()
    for it in file_items:
        seen.add(it["code"])
        b = builtin.get(it["code"])
        if b is None:
            diffs.append({"code": it["code"], "issue": "not in built-in scorer"})
            continue
        if (it["points"], it["bonus"]) != (b.points, b.bonus):
            diffs.append({
                "code": it["code"],
                "issue": "points changed",
                "file": {"points": it["points"], "bonus": it["bonus"]},
                "builtin": {"points": b.points, "bonus": b.bonus},
            })
    for code in builtin:
        if code not in seen:
            diffs.append({"code": code, "issue": "missing from rubric file"})
    return diffs
