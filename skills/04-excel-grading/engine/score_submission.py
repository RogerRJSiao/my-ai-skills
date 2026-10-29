"""Apply references/scoring_checks.md to the facts from engine/extract_submission.py.

The standard is the 評閱標準 xlsx (S06_excel_scoring_criteria_v<latest>.xlsx);
scoring_checks.md and this module are written from it. When they disagree,
the xlsx wins: fix the md and this code, never grade by the code as it stands.

Every item gets a score, a 評閱結果說明 note, and a `needs_review` flag for
the cases the SOP says a human (or the model) must confirm, such as a chart
Data Table that might count as 運算列表資料標籤. Callers can override any
item afterwards (see grading.apply_overrides).
"""
from __future__ import annotations

import re

from kd_mcp.core.grading.models import GradeResult, ItemScore, StudentInfo
from .extract_rubric import builtin_rubric_map

VEHICLE_CATEGORIES = {"福特", "賓士", "裕隆", "三菱"}
VEHICLE_SERIES = {"一月", "二月", "三月"}
GENERAL_CATEGORIES = {"統帥", "武力", "智力", "政治", "魅力"}
GENERAL_SERIES = {"趙雲", "周瑜", "程遠志", "劉禪"}
DEPARTMENTS = {"業務", "會計", "門市"}

# chart XML element -> name used in the 評閱結果說明
CHART_NAMES = {
    "barChart": "長條圖／直條圖", "radarChart": "雷達圖", "lineChart": "折線圖", "line3DChart": "立體折線圖",
    "pieChart": "圓餅圖", "pie3DChart": "立體圓餅圖", "doughnutChart": "環圈圖", "ofPieChart": "子母圓餅圖",
    "areaChart": "區域圖", "area3DChart": "立體區域圖", "scatterChart": "散佈圖", "bubbleChart": "泡泡圖",
    "stockChart": "股票圖", "surfaceChart": "曲面圖", "surface3DChart": "立體曲面圖",
}


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


def _number_format_kind(fmt: str) -> str:
    """currency: 「$」 before the digits (rubric 1.4); number: 通用格式 or a plain
    數值 format; other: anything else. Only the positive section is read."""
    section = fmt.split(";")[0]
    digits = re.search(r"[0#]", section)
    if "$" in section and digits and section.index("$") < digits.start():
        return "currency"
    if fmt == "General" or re.fullmatch(r"[#,0.]+(_\))?", section):
        return "number"
    return "other"


class _Builder:
    def __init__(self):
        self.rubric = builtin_rubric_map()
        self.items: list[ItemScore] = []

    def add(self, code: str, score: float, note: str, needs_review: bool = False):
        r = self.rubric[code]
        score = max(0.0, min(float(score), r.max_score))
        self.items.append(ItemScore(r.section, code, r.title, score, r.max_score, note,
                                    needs_review, is_bonus=r.points == 0))

    def zero_all(self, codes, note: str, needs_review: bool = False):
        for code in codes:
            self.add(code, 0, note, needs_review)


# ── 考題1 ──────────────────────────────────────────────────────────────────

_REJECT_NOTES = {
    "wrong_range": "來源範圍是【{source_sheet}】{source_range}，不是 A1:J13",
    "other_sheet": "來源是【{source_sheet}】{source_range}，不是【薪資】A1:J13",
    "unresolved": "來源「{source_name_or_range}」無法判讀是哪個活頁與範圍",
}


def _score_salary(b: _Builder, pivot: dict | None, rejected: list[dict], sheet_problem: str | None):
    codes = ("1.1", "1.2", "1.3", "1.4", "1.5")
    if sheet_problem:
        b.zero_all(codes, f"{sheet_problem}，考題1全部0分。")
        return
    if not pivot:
        if rejected:
            found = "；".join(f"【{r['host_sheet']}】的樞紐表" + _REJECT_NOTES[r["reason"]].format(**r)
                             for r in rejected)
            review = any(r["reason"] == "unresolved" for r in rejected)
            b.zero_all(codes, f"沒有以【薪資】A1:J13 為來源的樞紐表（{found}），考題1全部0分。"
                       + ("請人工確認來源。" if review else ""), review)
        else:
            b.zero_all(codes, "未建立以【薪資】資料為來源的樞紐表（已排除【考題】活頁範例），考題1全部0分。")
        return

    b.add("1.1", 5, f"已在【{pivot.get('host_sheet')}】活頁建立以【薪資】A1:J13 為來源的樞紐表。")

    rows = _clean_set(pivot.get("row_visible_labels"))
    cols = _clean_set(pivot.get("col_visible_labels"))
    if not cols or not set(cols) <= DEPARTMENTS:
        b.add("1.2", 0, f"欄標籤為 {cols or '未設定'}，不是業務、會計、門市，直接給0分。")
    elif len(rows) != 2 or set(rows) != {"男", "女"}:
        b.add("1.2", 0, f"列標籤為 {rows or '未設定'}，不是男、女，直接給0分。")
    elif cols == ["業務"]:
        b.add("1.2", 10, "欄標籤只篩出「業務」；列標籤為男、女。")
    else:
        b.add("1.2", 5, f"欄標籤為 {cols}，未只留下「業務」，扣5分；列標籤為男、女。")

    subtotal = (pivot.get("data_fields") or [{}])[0].get("subtotal")
    if subtotal == "average":
        b.add("1.3", 10, "值的彙總方式為平均值。")
    elif subtotal in (None, "sum"):  # Excel omits subtotal when it is sum
        b.add("1.3", 5, "值的彙總方式為加總（非平均值），給一半。")
    else:
        b.add("1.3", 0, f"值的彙總方式為「{subtotal}」，非平均值。")

    fmts = pivot.get("value_number_formats") or []
    kinds = {_number_format_kind(f) for f in fmts}
    if "currency" in kinds:
        b.add("1.4", 5, "樞紐值顯示為「$」在數值前面的貨幣格式。")
    elif kinds == {"number"}:
        b.add("1.4", 2.5, f"樞紐值為通用格式／數值格式 {fmts}，給一半。")
    else:
        b.add("1.4", 0, f"樞紐值格式為 {fmts or '無法判讀'}，非「$」在數值前面的貨幣格式。", not fmts)

    if pivot.get("source_is_named_table"):
        b.add("1.5", 6, f"樞紐表來源為表格「{pivot.get('source_name_or_range')}」，額外加分。")
    else:
        b.add("1.5", 0, f"樞紐表來源為 {pivot.get('source_name_or_range')}，未先把資料設為表格。")


# ── 考題2 ──────────────────────────────────────────────────────────────────

def _score_grades(b: _Builder, g: dict | None, sheet_problem: str | None):
    if not g:
        b.zero_all(("2.1", "2.2", "2.3", "2.4", "2.5", "2.6"),
                   f"{sheet_problem or '找不到【成績】活頁'}，考題2全部0分。")
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

    # 合併後置中 on exactly A1:G1 and the title text are scored separately (2 + 2)
    a1_merge = g.get("a1_merged_range")
    if g.get("a1_merge_and_center"):
        merge_pts, merge_note = 2, "A1:G1 已合併後置中"
    elif a1_merge and a1_merge != "A1:G1":
        merge_pts, merge_note = 0, f"合併範圍為 {a1_merge}，不是剛好 A1:G1"
    elif a1_merge and g.get("a1_horizontal") == "centerContinuous":
        merge_pts, merge_note = 0, "A1:G1 已合併，但對齊是跨欄置中，不是合併後置中"
    elif a1_merge:
        merge_pts, merge_note = 0, "A1:G1 已合併但未置中"
    elif g.get("a1_horizontal") == "centerContinuous":
        merge_pts, merge_note = 0, "A1 只設跨欄置中、未合併儲存格"
    else:
        merge_pts, merge_note = 0, "A1:G1 未合併後置中"
    if g.get("header_text_preserved"):
        title_pts, title_note = 2, "標題文字為「經濟甲班 期中考成績表」"
    else:
        title_pts, title_note = 0, f"標題文字被改為「{g.get('a1_value')}」"
    b.add("2.2", merge_pts + title_pts, f"{merge_note}；{title_note}。")

    avg_row = g.get("avg_row_index")
    half_g = 4.5 if g.get("g3_g15_average_formulas_ok") else 0
    half_row = 4.5 if (g.get("row_avg_formulas_ok") and g.get("row_avg_integer_format")) else 0
    notes = ["G3:G15 以AVERAGE計算個人平均" if half_g else "G3:G15 未全部以AVERAGE計算個人平均"]
    if not avg_row:
        notes.append("找不到「班平均」列")
    elif half_row:
        notes.append(f"第{avg_row}列班平均以AVERAGE計算且只保留到整數位")
    elif g.get("row_avg_formulas_ok"):
        notes.append(f"第{avg_row}列班平均有AVERAGE但未只保留到整數位")
    elif g.get("row_avg_uses_subtotal"):
        notes.append(f"第{avg_row}列班平均用SUBTOTAL(101)取代AVERAGE，此半不給分")
    else:
        notes.append(f"第{avg_row}列班平均未全部以AVERAGE計算")
    b.add("2.3", half_g + half_row, "；".join(notes) + "。")

    # 排序看結果：G3:G15 由大到小；沒有快取值時才退回看 sortState
    sorts = g.get("sort_conditions") or []
    state_ok = any(s.get("descending") and (s.get("ref") or "").startswith("G") for s in sorts)
    desc = g.get("g_values_descending")
    sort_ok = desc if desc is not None else state_ok
    sort_review = desc is None and not state_ok
    if desc is not None:
        sort_note = "G3:G15 個人平均已由大到小排序" if desc else "G3:G15 個人平均未由大到小排序"
    else:
        sort_note = ("排序條件為個人平均由大到小（檔案沒有計算結果，無法核對數值）" if state_ok
                     else "檔案沒有計算結果，無法確認G3:G15是否由大到小排序，請人工確認")

    filters = g.get("filter_conditions") or []
    # 數字篩選 on 個人平均 (column G); > 70 and >= 70 both count
    gt70 = [f for f in filters if f.get("column") == "G" and _num_eq(f.get("val"), 70)
            and f.get("operator") in ("greaterThan", "greaterThanOrEqual")]
    visible = g.get("visible_data_rows")
    filter_ok = bool(gt70) and visible == 9
    if gt70:
        op = ">70" if gt70[0]["operator"] == "greaterThan" else "≥70"
        filter_note = (f"已篩選個人平均 {op}，留下9列資料" if filter_ok
                       else f"篩選條件為個人平均 {op}，但最後留下 {visible} 列（應為8位學生加班平均共9列）")
    elif any(f.get("operator") == "values" for f in filters):
        filter_note = "個人平均是勾選數值清單篩選，不是「數字篩選 > 70」"
    elif any(f.get("column") != "G" for f in filters):
        filter_note = "篩選條件設在個人平均以外的欄位"
    elif filters:
        filter_note = f"篩選條件為 {filters}"
    else:
        filter_note = "未設定 >70 篩選"
    b.add("2.4", (5 if sort_ok else 0) + (5 if filter_ok else 0), f"{sort_note}；{filter_note}。",
          sort_review)

    if g.get("borders_complete"):
        b.add("2.5", 3, "A1:G16 範圍皆有框線。")
    else:
        missing = g.get("missing_border_cells") or []
        b.add("2.5", 0, f"A1:G16 有 {len(missing)} 格未設框線（例：{', '.join(missing[:5])}）。")

    cell = g.get("g16_cell") or {}
    if cell.get("is_empty"):
        b.add("2.6", 3, "G16 未填入公式、數值或文字，額外加分。")
    else:
        b.add("2.6", 0, f"G16 有公式、數值或文字（{cell.get('value')}），不給額外加分。")


# ── 考題3 / 考題4 ──────────────────────────────────────────────────────────

def _chart_name(chart: dict) -> str:
    types = chart.get("chart_types") or []
    return "＋".join(CHART_NAMES.get(t, t) for t in types) or "無法辨識的圖表"


def _chart_gate(chart: dict | None, raw: dict | None, expected_type: str, type_name: str,
                chartex: int = 0, chartex_ok: bool = False) -> tuple[str, bool] | None:
    """Return (reason, needs_review) to zero the whole section, or None if gradable.

    Only a dynamic chart of the expected type counts; any other chart is
    ignored (rubric 3.1 / 4.1). chartex counts the sheet's Excel 2016+ charts,
    which are not inspected; with chartex_ok (考題3 accepts 直方圖) a sheet
    whose only candidate is such a chart is zeroed for review."""
    if raw and not raw.get("unchanged"):
        cells = "、".join(d["cell"] for d in raw.get("diffs", []))
        return f"A欄或第2列原始資料被更動（{cells}），整大題0分。", False
    if chart is None:
        problem = "活頁中沒有圖表"
    elif chart.get("literal_only") or not chart.get("is_dynamic"):
        problem = "圖表為貼上的固定數值（非動態參照儲存格）"
    elif chart.get("chart_type") != expected_type:
        problem = f"圖表類型為{_chart_name(chart)}，不是{type_name}"
    else:
        return None
    if chartex and chartex_ok:
        lead = "活頁只有" if chart is None else f"{problem}；另有"
        return (f"{lead} Excel 2016 以後的新式圖表（如直方圖），程式無法判讀，"
                "暫給0分，請人工確認。"), True
    if chartex and chart is None:
        return f"活頁只有直方圖等新式圖表，不是{type_name}，整大題0分。", False
    return f"{problem}，整大題0分。", False


_RAW_UNCHECKED = "（未提供題目檔，未檢查A欄及第2列原始資料是否被更動，請人工確認）"


def _score_vehicle(b: _Builder, chart: dict | None, raw: dict | None, chartex: int = 0,
                   raw_checked: bool = True, sheet_problem: str | None = None):
    codes = ("3.1", "3.2", "3.3", "3.4")
    if sheet_problem:
        b.zero_all(codes, f"{sheet_problem}，考題3全部0分。")
        return
    gate = _chart_gate(chart, raw, "barChart", "長條圖／直條圖／直方圖", chartex, chartex_ok=True)
    if gate:
        for code in codes:
            b.add(code, 0, *gate)
        return

    b.add("3.1", 5, "已建立動態的" + ("立體" if chart.get("is_3d") else "")
          + ("橫條圖。" if chart.get("bar_dir") == "bar" else "直條圖。")
          + ("" if raw_checked else _RAW_UNCHECKED), not raw_checked)

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


def _score_general(b: _Builder, chart: dict | None, raw: dict | None, chartex: int = 0,
                   raw_checked: bool = True, sheet_problem: str | None = None):
    codes = ("4.1", "4.2", "4.3")
    if sheet_problem:
        b.zero_all(codes, f"{sheet_problem}，考題4全部0分。")
        return
    # 4.x accepts radar charts only, so a 直方圖 (chartEx) is never reviewed
    gate = _chart_gate(chart, raw, "radarChart", "雷達圖", chartex)
    if gate:
        for code in codes:
            b.add(code, 0, *gate)
        return

    b.add("4.1", 5, "已建立動態的雷達圖。" + ("" if raw_checked else _RAW_UNCHECKED), not raw_checked)

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

def _sheet_problems(facts: dict) -> dict[str, str]:
    """Expected sheet -> why it could not be graded (ambiguous or missing name)."""
    out = {s: f"有多個工作表名稱含「{s}」（{'、'.join(found)}），無法判斷要評哪一個"
           for s, found in (facts.get("ambiguous_sheets") or {}).items()}
    out.update({s: f"找不到【{s}】活頁" for s in facts.get("missing_sheets") or []})
    return out


def score_facts(facts: dict, student: StudentInfo) -> GradeResult:
    b = _Builder()
    warnings = []
    if not student.filename_valid:
        warnings.append(f"檔名「{student.filename}」不符合「部門代號-工號-姓名.xlsx」格式，請向使用者確認。")
    if facts.get("missing_sheets"):
        warnings.append(f"找不到工作表：{'、'.join(facts['missing_sheets'])}，請確認檔案是否正確。")
    for expected, found in (facts.get("ambiguous_sheets") or {}).items():
        warnings.append(f"有多個工作表名稱含「{expected}」（{'、'.join(found)}），無法判斷要評哪一個，"
                        f"該大題以0分計，請向使用者確認。")
    renamed = {e: n for e, n in (facts.get("sheet_names") or {}).items() if n != e}
    if renamed:
        warnings.append("工作表名稱有前後綴，已對應為："
                        + "、".join(f"{e}→「{n}」" for e, n in renamed.items()) + "。")
    raw_checked = "車輛銷售_raw_data_check" in facts
    if not raw_checked:
        warnings.append("未提供原始題目檔，無法檢查原始資料是否被更動。")
    for key in ("chart_車輛銷售", "chart_武將能力"):
        if (facts.get(key) or {}).get("chart_count", 0) > 1:
            warnings.append(f"{key} 有多張圖表，已只評分其中動態且類型正確的那一張。")

    problems = _sheet_problems(facts)
    chartex = facts.get("chartex_count") or {}
    _score_salary(b, facts.get("pivot_薪資"), facts.get("pivot_薪資_rejected") or [], problems.get("薪資"))
    _score_grades(b, facts.get("成績"), problems.get("成績"))
    _score_vehicle(b, facts.get("chart_車輛銷售"), facts.get("車輛銷售_raw_data_check"),
                   chartex.get("車輛銷售", 0), raw_checked, problems.get("車輛銷售"))
    _score_general(b, facts.get("chart_武將能力"), facts.get("武將能力_raw_data_check"),
                   chartex.get("武將能力", 0), raw_checked, problems.get("武將能力"))
    return GradeResult(student, b.items, warnings)
