"""Apply references/scoring_checks.md to the facts from xlsx_s06.inspection.

Every item gets a score, a 評閱結果說明 note, and a `needs_review` flag for
the cases the SOP says a human (or the model) must confirm, such as a chart
Data Table that might count as 運算列表資料標籤. Callers can override any
item afterwards (see grading.apply_overrides).
"""
from __future__ import annotations

import re

from ...models import GradeResult, ItemScore, StudentInfo
from .rubric import builtin_rubric_map

VEHICLE_CATEGORIES = {"福特", "賓士", "裕隆", "三菱"}
VEHICLE_SERIES = {"一月", "二月", "三月"}
GENERAL_CATEGORIES = {"統帥", "武力", "智力", "政治", "魅力"}
GENERAL_SERIES = {"趙雲", "周瑜", "程遠志", "劉禪"}


# ── helpers ────────────────────────────────────────────────────────────────

def _clean_set(values) -> list[str]:
    return [str(v).strip() for v in (values or []) if v is not None]


def _exact_set(values, expected: set[str]) -> bool:
    vals = _clean_set(values)
    return len(vals) == len(expected) and set(vals) == expected


def _num_eq(v, target: float) -> bool:
    try:
        return v is not None and float(v) == target
    except ValueError:
        return False


class _Builder:
    def __init__(self):
        self.rubric = builtin_rubric_map()
        self.items: list[ItemScore] = []

    def add(self, code: str, score: float, note: str, needs_review: bool = False):
        r = self.rubric[code]
        score = max(0.0, min(float(score), r.max_score))
        self.items.append(ItemScore(r.section, code, r.title, score, r.max_score, note,
                                    needs_review, is_bonus=r.points == 0))

    def zero_all(self, codes, note: str):
        for code in codes:
            self.add(code, 0, note)


# ── 考題1 ──────────────────────────────────────────────────────────────────

def _score_salary(b: _Builder, pivot: dict | None):
    if not pivot:
        b.zero_all(("1.1", "1.2", "1.3", "1.4", "1.5"),
                   "【薪資】活頁未建立樞紐表（已排除【考題】活頁範例），考題1全部0分。")
        return

    b.add("1.1", 5, "【薪資】活頁已建立樞紐表。")

    rows = set(_clean_set(pivot.get("row_visible_labels")))
    cols = _clean_set(pivot.get("col_visible_labels"))
    row_ok = {"男", "女"} <= rows
    if cols == ["業務"]:
        col_ok, col_note, col_review = True, "欄標籤只篩出「業務」", False
    elif _exact_set(cols, {"業務", "會計", "門市"}):
        col_ok, col_note, col_review = True, "欄標籤完整顯示業務/會計/門市（折衷微調，給滿分）", False
    else:
        col_ok, col_note, col_review = False, f"欄標籤為 {cols or '未設定'}，未只篩出「業務」", bool(cols)
    row_note = "列標籤同時有男、女" if row_ok else f"列標籤為 {sorted(rows) or '未設定'}，缺男或女"
    b.add("1.2", (5 if col_ok else 0) + (5 if row_ok else 0), f"{col_note}；{row_note}。", col_review)

    subtotal = (pivot.get("data_fields") or [{}])[0].get("subtotal")
    if subtotal == "average":
        b.add("1.3", 10, "值的彙總方式為平均值。")
    elif subtotal in (None, "sum"):  # Excel omits subtotal when it is sum
        b.add("1.3", 5, "值的彙總方式為加總（非平均值），給一半。")
    else:
        b.add("1.3", 0, f"值的彙總方式為「{subtotal}」，非平均值。")

    fmts = pivot.get("value_number_formats") or []
    if any("$" in f for f in fmts):
        b.add("1.4", 5, "樞紐值顯示為含「$」的貨幣格式。")
    elif fmts and all(f == "General" or re.fullmatch(r"[#,0.]+(_\))?", f) for f in fmts):
        b.add("1.4", 2.5, f"樞紐值為通用格式／數值格式 {fmts}，給一半。")
    else:
        b.add("1.4", 0, f"樞紐值格式為 {fmts or '無法判讀'}，非貨幣格式。", not fmts)

    if pivot.get("source_is_named_table"):
        b.add("1.5", 6, f"樞紐表來源為表格「{pivot.get('source_name_or_range')}」，額外加分。")
    else:
        b.add("1.5", 0, f"樞紐表來源為儲存格範圍 {pivot.get('source_name_or_range')}，未先設為表格。")


# ── 考題2 ──────────────────────────────────────────────────────────────────

def _score_grades(b: _Builder, g: dict | None):
    if not g:
        b.zero_all(("2.1", "2.2", "2.3", "2.4", "2.5", "2.6"), "找不到【成績】活頁，考題2全部0分。")
        return

    fp = g.get("freeze_panes")
    m = re.fullmatch(r"([A-Z]+)(\d+)", fp or "")
    col_ok = bool(m) and m.group(1) == "B"
    row_ok = bool(m) and m.group(2) == "3"
    known = fp in (None, "A1", "B3", "A3", "B1")
    if col_ok and row_ok:
        b.add("2.1", 4, "凍結窗格設在B3，科目列與姓名欄皆固定。")
    else:
        parts = []
        if row_ok:
            parts.append("只固定科目列")
        if col_ok:
            parts.append("只固定姓名欄")
        note = f"凍結窗格為 {fp or '未設定'}" + (f"（{'、'.join(parts)}）" if parts else "")
        b.add("2.1", (2 if row_ok else 0) + (2 if col_ok else 0), note + "。", not known)

    if g.get("a1_center_or_merged"):
        if g.get("header_text_preserved"):
            b.add("2.2", 4, "A1:G1 已跨欄置中，標題文字維持「經濟甲班 期中考成績表」。")
        else:
            b.add("2.2", 2, f"A1:G1 已跨欄置中，但標題文字被改為「{g.get('a1_value')}」。")
    else:
        b.add("2.2", 0, "A1:G1 未設定跨欄置中。")

    avg_row = g.get("avg_row_index")
    half_g = 4.5 if g.get("g3_g15_average_formulas_ok") else 0
    half_row = 4.5 if (g.get("row_avg_formulas_ok") and g.get("row_avg_integer_format")) else 0
    notes = ["G3:G15 以AVERAGE計算個人平均" if half_g else "G3:G15 未全部以AVERAGE計算個人平均"]
    if not avg_row:
        notes.append("找不到「班平均」列")
    elif half_row:
        notes.append(f"第{avg_row}列班平均以AVERAGE計算且顯示整數")
    elif g.get("row_avg_formulas_ok"):
        notes.append(f"第{avg_row}列班平均有AVERAGE但未設為整數格式")
    else:
        notes.append(f"第{avg_row}列班平均未全部以AVERAGE計算")
    b.add("2.3", half_g + half_row, "；".join(notes) + "。")

    sorts = g.get("sort_conditions") or []
    sort_ok = any(s.get("descending") and (s.get("ref") or "").startswith("G") for s in sorts)
    filters = g.get("filter_conditions") or []
    filter_ok = any(f.get("operator") == "greaterThan" and _num_eq(f.get("val"), 70) for f in filters)
    filter_review = bool(filters) and not filter_ok
    sort_review = bool(sorts) and not sort_ok
    notes = [
        "個人平均已由大到小排序" if sort_ok else ("有排序但非個人平均由大到小" if sorts else "未排序"),
        "已篩選個人平均 >70" if filter_ok else (f"篩選條件為 {filters}" if filters else "未設定 >70 篩選"),
    ]
    if filter_ok and len(g.get("hidden_rows") or []) == 0:
        notes.append("但篩選後沒有隱藏任何列，請確認")
        filter_review = True
    b.add("2.4", (5 if sort_ok else 0) + (5 if filter_ok else 0), "；".join(notes) + "。",
          sort_review or filter_review)

    if g.get("borders_complete"):
        b.add("2.5", 3, "A1:G16 範圍皆有框線。")
    else:
        missing = g.get("missing_border_cells") or []
        b.add("2.5", 0, f"A1:G16 有 {len(missing)} 格未設框線（例：{', '.join(missing[:5])}）。")

    cell = g.get("g_avg_row_cell") or {}
    if avg_row and cell.get("is_empty"):
        review = not g.get("g3_g15_average_formulas_ok")
        b.add("2.6", 3, f"G{avg_row} 未使用公式，額外加分。"
              + ("（但 G3:G15 也未作答，請確認是否仍給分）" if review else ""), review)
    else:
        b.add("2.6", 0, f"G{avg_row or 16} 有公式或數值（{cell.get('value')}），不給額外加分。")


# ── 考題3 / 考題4 ──────────────────────────────────────────────────────────

def _chart_gate(chart: dict | None, raw: dict | None) -> str | None:
    """Return a reason to zero the whole section, or None if gradable."""
    if raw and not raw.get("unchanged"):
        cells = "、".join(d["cell"] for d in raw.get("diffs", []))
        return f"A欄或第2列原始資料被更動（{cells}），整大題0分。"
    if chart is None:
        return "活頁中沒有圖表，整大題0分。"
    if chart.get("literal_only") or not chart.get("is_dynamic"):
        return "圖表為貼上的固定數值（非動態參照儲存格），整大題0分。"
    return None


def _score_vehicle(b: _Builder, chart: dict | None, raw: dict | None):
    codes = ("3.1", "3.2", "3.3", "3.4")
    reason = _chart_gate(chart, raw)
    if reason:
        b.zero_all(codes, reason)
        return

    if chart.get("chart_type") == "barChart":
        b.add("3.1", 5, "已建立動態的直條圖／橫條圖。")
    else:
        b.add("3.1", 0, f"圖表類型為 {chart.get('chart_type')}，非直條圖或橫條圖。", True)

    cats = _clean_set(chart.get("categories"))
    cat_ok = _exact_set(cats, VEHICLE_CATEGORIES)
    axis_ok = _num_eq(chart.get("y_max"), 2000) and _num_eq(chart.get("y_major_unit"), 500)
    grid_ok = bool(chart.get("gridlines_none"))
    notes = [
        "類別軸為四個廠牌" if cat_ok else f"類別軸為 {cats}（應為福特/賓士/裕隆/三菱）",
        "數值軸最大值2000、主要單位500" if axis_ok
        else f"數值軸最大值 {chart.get('y_max')}、主要單位 {chart.get('y_major_unit')}（應為2000/500）",
        "主要格線已設為無" if grid_ok else "主要格線未移除",
    ]
    b.add("3.2", 5 * (cat_ok + axis_ok + grid_ok), "；".join(notes) + "。")

    series = _clean_set(chart.get("series_names"))
    if _exact_set(series, VEHICLE_SERIES):
        b.add("3.3", 5, "圖例只有一月、二月、三月。")
    else:
        b.add("3.3", 0, f"圖例為 {series}（應只有一月/二月/三月）。")

    pts, notes, review = 0, [], False
    if chart.get("has_data_labels"):
        pts += 3
        notes.append("有資料標籤")
    elif chart.get("has_data_table"):
        review = True
        notes.append("未加逐點資料標籤，但有「運算列表(資料表)」，請人工確認是否給分")
    else:
        notes.append("無資料標籤")
    if chart.get("has_chart_title"):
        pts += 3
        notes.append("有圖表標題")
    else:
        notes.append("無圖表標題")
    if chart.get("has_x_axis_title") and chart.get("has_y_axis_title"):
        pts += 3
        notes.append("縱橫座標軸皆有標題")
    else:
        notes.append("座標軸標題不完整")
    b.add("3.4", pts, "；".join(notes) + "。", review)


def _score_general(b: _Builder, chart: dict | None, raw: dict | None):
    codes = ("4.1", "4.2", "4.3")
    reason = _chart_gate(chart, raw)
    if reason:
        b.zero_all(codes, reason)
        return

    if chart.get("chart_type") == "radarChart":
        b.add("4.1", 5, "已建立動態的雷達圖。")
    else:
        b.add("4.1", 0, f"圖表類型為 {chart.get('chart_type')}，非雷達圖。", True)

    cats = _clean_set(chart.get("categories"))
    series = _clean_set(chart.get("series_names"))
    cat_ok = _exact_set(cats, GENERAL_CATEGORIES)
    ser_ok = _exact_set(series, GENERAL_SERIES)
    notes = [
        "座標軸為統帥/武力/智力/政治/魅力" if cat_ok else f"座標軸為 {cats}",
        "圖例為趙雲/周瑜/程遠志/劉禪" if ser_ok else f"圖例為 {series}",
    ]
    b.add("4.2", 5 * (cat_ok + ser_ok), "；".join(notes) + "。")

    pts, notes, review = 0, [], False
    if chart.get("has_chart_title"):
        pts += 3
        notes.append("有圖表標題")
    else:
        notes.append("無圖表標題")
    symbols = [s for s in chart.get("marker_symbols") or [] if s]
    if any(s != "none" for s in symbols):
        pts += 3
        notes.append("使用含資料點標記的雷達圖")
    elif not symbols and chart.get("radar_style") == "marker":
        review = True
        notes.append("radarStyle=marker 但未指定標記符號，請人工確認是否有顯示資料點標記")
    else:
        notes.append("未顯示資料點標記")
    b.add("4.3", pts, "；".join(notes) + "。", review)


# ── entry points ───────────────────────────────────────────────────────────

def score_facts(facts: dict, student: StudentInfo) -> GradeResult:
    b = _Builder()
    warnings = []
    if not student.filename_valid:
        warnings.append(f"檔名「{student.filename}」不符合「部門代號-工號-姓名.xlsx」格式，請向使用者確認。")
    if facts.get("missing_sheets"):
        warnings.append(f"找不到工作表：{'、'.join(facts['missing_sheets'])}，請確認檔案是否正確。")
    if "車輛銷售_raw_data_check" not in facts:
        warnings.append("未提供原始題目檔，無法檢查原始資料是否被更動。")
    for key in ("chart_車輛銷售", "chart_武將能力"):
        if (facts.get(key) or {}).get("chart_count", 0) > 1:
            warnings.append(f"{key} 有多張圖表，已只評分其中動態且類型正確的那一張。")

    _score_salary(b, facts.get("pivot_薪資"))
    _score_grades(b, facts.get("成績"))
    _score_vehicle(b, facts.get("chart_車輛銷售"), facts.get("車輛銷售_raw_data_check"))
    _score_general(b, facts.get("chart_武將能力"), facts.get("武將能力_raw_data_check"))
    return GradeResult(student, b.items, warnings)
