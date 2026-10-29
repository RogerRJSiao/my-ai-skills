#!/usr/bin/env python3
"""
check_rubric.py — check that references/scoring_checks.md and engine/ still
follow the 評閱標準 (S06_excel_scoring_criteria_v<latest>.xlsx).

Order of precedence: 評閱標準 xlsx (latest version) > scoring_checks.md > engine/ py.
When this script reports a difference, fix scoring_checks.md and engine/ to
match the xlsx, never the other way round.

The batch folder is <data root>/grading/<domain>/<課程代碼>_<梯次>/, e.g.
D:\\claude_data\\grading\\training_internal\\S06_<yyyymm>\\ (resolved like
init_batch.py; --batch is required). The 評閱標準 compared against is the
latest <batch>/01_template/S06_excel_scoring_criteria_v*.xlsx, or the file
given with --rubric; the answer and question files for check 4 come from the
same folder. Inputs are only read; the results are written as Markdown to
<batch>/02_work/rubric/S06_excel_scoring_criteria_<yyyymmdd_HHMMSS>.md (the
folder is created when missing) and nothing else is written. Checks:
  1. 評閱標準版本    which file is compared; skill.toml must point to the latest
                     version in assets, and assets must hold the batch's file
                     unchanged (the MCP server grades with assets)
  2. 配分與小項      xlsx items vs engine/extract_rubric.py BUILTIN_RUBRIC vs the
                     table at the top of scoring_checks.md: item set, 大項,
                     小項名稱, 配分, 額外加分, and the 100 / 124 totals
  3. 關鍵數值        values written in the criteria text (ranges, cells, labels,
                     thresholds) vs the engine's constants and scoring_checks.md
  4. 標準解答檔試評  grades the answer file with the engine: every base item
                     must get full marks (needs mcp-server/src for kd_mcp,
                     skipped when it is not there)
What the criteria text means (e.g. when 1.2 is zero) cannot be checked by a
script: read the xlsx against scoring_checks.md for that (the report lists the
criteria text; --show-text prints it too). Exit code 0 = all passed.

Usage:
    python check_rubric.py --batch <yyyymm>    # 比對 S06_<yyyymm>/01_template 的評閱標準
    python check_rubric.py --root D:\\claude_data --batch <yyyymm>
    python check_rubric.py --batch <yyyymm> --rubric D:\\path\\S06_excel_scoring_criteria_v2.xlsx
    python check_rubric.py --batch <yyyymm> --show-text
"""
import re
import ast
import sys
import argparse
import tomllib
import datetime as dt
import importlib.util
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPTS_DIR.parent
ENGINE_DIR = SKILL_DIR / "engine"
SCORING_CHECKS = SKILL_DIR / "references" / "scoring_checks.md"
MCP_SRC = SKILL_DIR.parent.parent / "mcp-server" / "src"   # only in the repo checkout

sys.path.insert(0, str(SCRIPTS_DIR))
from init_batch import GRADING_DIR, course_code, default_root, domain   # noqa: E402

VERSION_RE = re.compile(r"_v(\d+(?:\.\d+)*)\.xlsx$", re.IGNORECASE)


class Report:
    """Prints each result as soon as it is known and keeps it for the md report."""

    def __init__(self):
        self.failed = 0
        self.sections: list[tuple[str, list[tuple[str, str]]]] = []

    def section(self, title):
        self.sections.append((title, []))
        print(f"\n{title}", flush=True)

    def _add(self, kind, msg, line):
        self.sections[-1][1].append((kind, msg))
        print(line, flush=True)

    def ok(self, msg):
        self._add("OK", msg, f"  [OK]   {msg}")

    def fail(self, msg):
        self.failed += 1
        self._add("FAIL", msg, f"  [FAIL] {msg}")

    def info(self, msg):
        self._add("INFO", msg, f"         {msg}")


# ── loading ───────────────────────────────────────────────────────────────

def load_module(name: str, path: Path):
    """Import one engine file on its own (it must not need kd_mcp)."""
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module   # dataclasses look their module up here
    spec.loader.exec_module(module)
    return module


def module_constants(path: Path) -> dict:
    """Module-level NAME = <literal> assignments of a .py file, read without
    importing it (score_submission.py needs kd_mcp)."""
    out = {}
    for node in ast.parse(path.read_text(encoding="utf-8")).body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            try:
                out[node.targets[0].id] = ast.literal_eval(node.value)
            except ValueError:
                pass
    return out


def version_key(path: Path) -> tuple:
    m = VERSION_RE.search(path.name)
    return tuple(int(p) for p in m.group(1).split(".")) if m else ()


def asset_names() -> dict[str, str]:
    """skill.toml [assets] file names, e.g. {"rubric": "S06_excel_scoring_criteria_v1.xlsx"}."""
    assets = tomllib.loads((SKILL_DIR / "skill.toml").read_text(encoding="utf-8")).get("assets", {})
    return {k: Path(v).name for k, v in assets.items()}


def versions_of(folder: Path, configured_name: str) -> list[Path]:
    """Every version of a template in folder, e.g. S06_excel_scoring_criteria_v*.xlsx."""
    stem = VERSION_RE.sub("", configured_name)
    if not folder.is_dir():
        return []
    return sorted((p for p in folder.iterdir() if p.is_file() and p.name.startswith(stem + "_v")
                   and VERSION_RE.search(p.name) and not p.name.startswith("~$")), key=version_key)


def md_table(md: str) -> dict[str, dict]:
    """The 小項 table at the top of scoring_checks.md: code -> {section, points, bonus}."""
    def num(cell):
        m = re.match(r"\s*(\d+(?:\.\d+)?)", cell)
        return float(m.group(1)) if m else 0.0

    out = {}
    for line in md.splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) == 4 and re.fullmatch(r"\d+\.\d+", cells[0]):
            out[cells[0]] = {"section": cells[1], "points": num(cells[2]), "bonus": num(cells[3])}
    return out


def norm_title(t: str) -> str:
    return re.sub(r"[。\s]", "", str(t or "")).replace("（", "(").replace("）", ")")


def fmt(v) -> str:
    if isinstance(v, (set, frozenset, list, tuple)):
        return "、".join(sorted(str(x) for x in v))
    return str(v)


# ── 1. version ────────────────────────────────────────────────────────────

def batch_dir(args) -> Path:
    """<data root>/grading/<domain>/<課程代碼>_<梯次>/, resolved like init_batch.py."""
    root_str, _ = (args.root, "--root") if args.root else default_root()
    return Path(root_str) / GRADING_DIR / domain() / f"{course_code()}_{args.batch}"


def pick_rubric(args, names: dict, rep: Report) -> Path | None:
    """The 評閱標準 to compare against: --rubric, else the latest version in
    the batch's 01_template (the batch's own standard)."""
    rep.section("1. 評閱標準版本")
    if args.rubric:
        path = Path(args.rubric)
        if not path.is_file():
            rep.fail(f"找不到 {path}")
            return None
        rep.ok(f"使用指定的評閱標準：{path}")
        return path

    assets = versions_of(SKILL_DIR / "assets" / "01_template", names["rubric"])
    if assets:
        rep.info(f"assets/01_template：{', '.join(p.name for p in assets)}")
        if assets[-1].name == names["rubric"]:
            rep.ok(f"skill.toml 指向 assets 最新版本 {assets[-1].name}")
        else:
            rep.fail(f"skill.toml 指向 {names['rubric']}，但 assets 最新版本是 {assets[-1].name}")

    folder = batch_dir(args) / "01_template"
    batch = versions_of(folder, names["rubric"])
    if not batch:
        rep.fail(f"梯次 {folder} 沒有 {VERSION_RE.sub('', names['rubric'])}_v*.xlsx")
        return None
    latest = batch[-1]
    rep.info(f"梯次 01_template：{', '.join(p.name for p in batch)}")
    # the MCP server grades with skill.toml's copy in assets, so it must be the batch's standard too
    same = next((a for a in assets if a.name == latest.name), None)
    if same is None:
        rep.fail(f"assets/01_template 沒有 {latest.name}：MCP 的 gi_grade_submission 用的是 skill.toml 指定的 "
                 f"{names['rubric']}，請把梯次這份放進 assets 並更新 skill.toml")
    elif same.read_bytes() != latest.read_bytes():
        rep.fail(f"assets 的 {latest.name} 與梯次內容不同：MCP 的 gi_grade_submission 用的是 assets 那份，"
                 f"請把梯次的檔案複製到 assets")
    else:
        rep.ok(f"assets 有相同內容的 {latest.name}（MCP 評分用的是同一份）")
    rep.ok(f"本次比對（梯次最新版本）：{latest}")
    return latest


# ── 2. points ─────────────────────────────────────────────────────────────

def check_points(items: list[dict], builtin: dict, table: dict, md: str, rep: Report):
    rep.section("2. 配分與小項（xlsx ↔ engine/extract_rubric.py ↔ scoring_checks.md 配分表）")
    codes = [it["code"] for it in items]
    for name, other in (("engine", builtin), ("scoring_checks.md", table)):
        missing, extra = sorted(set(codes) - set(other)), sorted(set(other) - set(codes))
        if missing or extra:
            rep.fail(f"{name} 的小項與 xlsx 不同：缺 {missing or '無'}，多 {extra or '無'}")
        else:
            rep.ok(f"{name} 的小項與 xlsx 相同（{len(codes)} 項）")

    bad = 0
    for it in items:
        b, t = builtin.get(it["code"]), table.get(it["code"])
        want = (it["section"], it["points"], it["bonus"])
        if b and (b.section, b.points, b.bonus) != want:
            bad += 1
            rep.fail(f"{it['code']} engine 為 {b.section} {b.points:g}/{b.bonus:g}，"
                     f"xlsx 為 {it['section']} {it['points']:g}/{it['bonus']:g}")
        if b and norm_title(b.title) != norm_title(it["title"]):
            bad += 1
            rep.fail(f"{it['code']} 小項名稱 engine 為「{b.title}」，xlsx 為「{it['title']}」")
        if t and (t["section"], t["points"], t["bonus"]) != want:
            bad += 1
            rep.fail(f"{it['code']} scoring_checks.md 為 {t['section']} {t['points']:g}/{t['bonus']:g}，"
                     f"xlsx 為 {it['section']} {it['points']:g}/{it['bonus']:g}")
    if not bad:
        rep.ok("每個小項的大項、名稱、配分、額外加分都一致")

    base = sum(it["points"] for it in items)
    top = base + sum(it["bonus"] for it in items)
    m = re.search(r"基本配分總和\s*=\s*(\d+)；含額外加分上限\s*=\s*(\d+)", md)
    if m and (float(m.group(1)), float(m.group(2))) == (base, top):
        rep.ok(f"基本配分總和 {base:g}、含額外加分上限 {top:g}，scoring_checks.md 一致")
    else:
        rep.fail(f"xlsx 基本配分總和 {base:g}、上限 {top:g}；scoring_checks.md 寫的是 "
                 f"{m.groups() if m else '（找不到這句）'}")


# ── 3. key values ─────────────────────────────────────────────────────────

def cells_of(rng: str) -> set[str]:
    m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", rng)
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    cols = [chr(c) for c in range(ord(c1), ord(c2) + 1)]   # single-letter columns are enough here
    return {f"{c}{r}" for c in cols for r in range(r1, r2 + 1)}


def grab(text: str, pattern: str, group: int = 1):
    m = re.search(pattern, text or "")
    return m.group(group) if m else None


def label_set(value: str | None):
    return {v.strip() for v in value.split("、")} if value else None


def check_values(crit: dict, es, ss: dict, md: str, rep: Report):
    rep.section("3. 關鍵數值（xlsx 評分標準文字 ↔ engine 常數 ↔ scoring_checks.md）")

    def check(code, label, xlsx_val, py_val, needles):
        name = f"{code} {label}"
        if xlsx_val is None:
            rep.fail(f"{name}：在 xlsx 的評分標準文字裡找不到，xlsx 可能改寫了，請人工比對並更新本腳本")
            return
        problems = []
        if py_val != xlsx_val:
            problems.append(f"engine 為 {fmt(py_val)}")
        missing = [n for n in needles if n not in md]
        if missing:
            problems.append(f"scoring_checks.md 沒有寫到 {missing}")
        if problems:
            rep.fail(f"{name}：xlsx 為 {fmt(xlsx_val)}，" + "；".join(problems))
        else:
            rep.ok(f"{name}：{fmt(xlsx_val)}")

    rng = grab(crit.get("1.1"), r"【薪資】活頁的\s*([A-Z]+\d+:[A-Z]+\d+)")
    check("1.1", "樞紐表來源範圍", rng, es.SALARY_SOURCE_RANGE, [f"【薪資】活頁的 {rng}"])
    rng = grab(crit.get("1.5"), r"【薪資】活頁的\s*([A-Z]+\d+:[A-Z]+\d+)")
    check("1.5", "設為表格的範圍", rng, es.SALARY_SOURCE_RANGE, [f"【薪資】{rng}"])

    t = crit.get("1.2")
    depts = grab(t, r"「欄標籤」若不是「([^」]+)」")
    check("1.2", "欄標籤可用的部門", label_set(depts), ss.get("DEPARTMENTS"), [f"「{depts}」"])
    keep = grab(t, r"只留下「([^」]+)」")
    check("1.2", "欄標籤只留下", keep, ss.get("KEEP_DEPARTMENT"), [f"只留下「{keep}」"])
    rows = grab(t, r"「列標籤」不是「([^」]+)」")
    check("1.2", "列標籤", label_set(rows), ss.get("ROW_LABELS"), [f"「{rows}」"])

    cell = grab(crit.get("2.1"), r"凍結窗格設在\s*([A-Z]+\d+)")
    check("2.1", "凍結窗格", cell, ss.get("FREEZE_CELL"), [f'"{cell}"'])

    t = crit.get("2.2")
    rng = grab(t, r"對\s*([A-Z]+\d+:[A-Z]+\d+)\s*設定")
    check("2.2", "合併後置中範圍", rng, es.GRADES_TITLE_RANGE, [rng])
    title = grab(t, r"標題是「([^」]+)」")
    check("2.2", "標題文字", title, es.GRADES_TITLE, [f"「{title}」"])

    t = crit.get("2.4")
    col = grab(t, r"「個人平均」\s*([A-Z]+)\d+:")
    check("2.4", "個人平均欄", col, ss.get("AVG_COLUMN"), [f'column == "{col}"'])
    value = grab(t, r"「大於(\d+)」")
    check("2.4", "數字篩選門檻", int(value) if value else None, ss.get("FILTER_VALUE"), [f">{value}", f"≥{value}"])
    left = grab(t, r"只留下(\d+)列")
    check("2.4", "篩選後留下的列數", int(left) if left else None, ss.get("FILTERED_ROWS"), [f"{left}列"])

    rng = grab(crit.get("2.5"), r"的\s*([A-Z]+\d+:[A-Z]+\d+)\s*範圍")
    check("2.5", "框線範圍", rng, es.GRADES_BORDER_RANGE, [rng])
    cell = grab(crit.get("2.6"), r"的\s*([A-Z]+\d+)\s*單格")
    check("2.6", "留空的儲存格", cell, es.GRADES_EMPTY_CELL, [cell])

    for code, sheet in (("3.1", "車輛銷售"), ("4.1", "武將能力")):
        m = re.search(r"第A欄[（(]\s*([A-Z]+\d+:[A-Z]+\d+)\s*[）)]及第2列[（(]\s*([A-Z]+\d+:[A-Z]+\d+)\s*[）)]",
                      crit.get(code) or "")
        col_a, row2 = es.RAW_DATA_CELLS[sheet]
        xlsx_cells = cells_of(m.group(1)) | cells_of(m.group(2)) if m else None
        check(code, f"{sheet}原始資料比對範圍",
              fmt(xlsx_cells) if m else None, fmt(set(col_a) | set(row2)),
              [f"比對 {m.group(1)} 與 {m.group(2)}"] if m else [])

    t = crit.get("3.2")
    cats = grab(t, r"水平座標軸只有「([^」]+)」")
    check("3.2", "類別軸", label_set(cats), ss.get("VEHICLE_CATEGORIES"),
          [f'"{c}"' for c in label_set(cats) or []])
    y_max = grab(t, r"「最大值」設為\s*(\d+)")
    check("3.2", "數值軸最大值", int(y_max) if y_max else None, ss.get("VEHICLE_Y_MAX"), [y_max or ""])
    unit = grab(t, r"「主要單位」\s*(\d+)")
    check("3.2", "主要單位", int(unit) if unit else None, ss.get("VEHICLE_Y_MAJOR_UNIT"), [unit or ""])
    series = grab(crit.get("3.3"), r"圖例項目只有「([^」]+)」")
    check("3.3", "圖例", label_set(series), ss.get("VEHICLE_SERIES"),
          [f'"{s}"' for s in label_set(series) or []])

    t = crit.get("4.2")
    cats = grab(t, r"水平座標軸只有「([^」]+)」")
    check("4.2", "類別軸", label_set(cats), ss.get("GENERAL_CATEGORIES"),
          [f'"{c}"' for c in label_set(cats) or []])
    series = grab(t, r"圖例項目只有「([^」]+)」")
    check("4.2", "圖例", label_set(series), ss.get("GENERAL_SERIES"),
          [f'"{s}"' for s in label_set(series) or []])


# ── 4. answer file ────────────────────────────────────────────────────────

def check_answer(folder: Path, names: dict, items: list[dict], rep: Report):
    rep.section("4. 標準解答檔試評（engine 給標準解答檔打分，基本配分應全拿）")
    answers = versions_of(folder, names["answer"])
    questions = versions_of(folder, names["question"])
    if not answers or not questions:
        rep.fail(f"{folder} 找不到標準解答檔或題目檔")
        return
    answer, question = answers[-1], questions[-1]
    if not MCP_SRC.is_dir():
        rep.info(f"略過：找不到 {MCP_SRC}（engine 計分需要 kd_mcp，只在 repo 裡能執行）")
        return
    sys.path[:0] = [str(MCP_SRC), str(SKILL_DIR)]
    try:
        from engine.extract_submission import inspect_submission
        from engine.score_submission import score_facts
        from kd_mcp.core.grading.models import parse_filename
    except ImportError as e:
        rep.info(f"略過：無法載入 engine（{e}）")
        return
    rep.info(f"標準解答檔 {answer.name}，題目檔 {question.name}")
    result = score_facts(inspect_submission(answer, question), parse_filename(answer))
    base = {it["code"] for it in items if it["points"] > 0}
    short = [i for i in result.items if i.code in base and i.score < i.max_score]
    for i in short:
        rep.fail(f"{i.code} 只拿 {i.score:g}/{i.max_score:g}：{i.note}")
    if not short:
        rep.ok(f"基本配分全拿，總分 {result.total:g}（額外加分項依解答檔的作法，不要求拿滿）")


# ── md report ─────────────────────────────────────────────────────────────

MD_MARK = {"OK": "✅", "FAIL": "❌", "INFO": "ℹ️"}


def md_cell(text) -> str:
    return str(text).replace("|", "\\|").replace("\r", "").replace("\n", "<br>")


def write_report(path: Path, rep: Report, rubric: Path | None, items: list[dict], now: dt.datetime):
    """The results as Markdown: summary, every check, then the criteria text
    for the manual reading of what each item means."""
    status = "全部通過" if not rep.failed else f"有 {rep.failed} 項未通過"
    lines = [
        f"# 評閱標準一致性檢查 {now:%Y-%m-%d %H:%M:%S}",
        "",
        "標準的優先順序：**評閱標準 xlsx（最新版本）＞ `references/scoring_checks.md` ＞ `engine/` 的 py**。"
        "有 ❌ 的項目，要以 xlsx 為準修改 scoring_checks.md 與 engine。",
        "",
        "| 項目 | 內容 |",
        "|---|---|",
        f"| 結果 | {'✅' if not rep.failed else '❌'} {status} |",
        f"| 評閱標準 | {md_cell(rubric) if rubric else '（找不到）'} |",
        f"| scoring_checks.md | {md_cell(SCORING_CHECKS)} |",
        f"| engine | {md_cell(ENGINE_DIR)} |",
        "| 未自動檢查 | 評分標準文字的語意（例如何時直接0分），請對照下方「評分標準文字」與 scoring_checks.md |",
    ]
    for title, entries in rep.sections:
        lines += ["", f"## {title}", "", "| 結果 | 說明 |", "|---|---|"]
        lines += [f"| {MD_MARK[kind]} | {md_cell(msg)} |" for kind, msg in entries]
    if items:
        lines += ["", "## 評分標準文字（語意請人工對照 scoring_checks.md）", "",
                  "| 小項 | 名稱 | 配分 | 額外加分 | 評分標準 |", "|---|---|---|---|---|"]
        lines += [f"| {it['code']} | {md_cell(it['title'])} | {it['points']:g} | {it['bonus']:g} "
                  f"| {md_cell(it['criteria'] or '')} |" for it in items]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# ── main ──────────────────────────────────────────────────────────────────

def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(
        description="Check scoring_checks.md and engine/ against the 評閱標準 xlsx.")
    ap.add_argument("--batch", required=True,
                    help="梯次(yyyymm)：比對 <資料根目錄>/grading/<domain>/<課程代碼>_<梯次>/01_template 最新版本的評閱標準，"
                         "結果寫成 <梯次>/02_work/rubric/<評閱標準檔名>_<yyyymmdd_HHMMSS>.md")
    ap.add_argument("--rubric", default=None, help="改用指定的評閱標準檔（例如使用者這次提供、尚未放進梯次的版本）")
    ap.add_argument("--root", default=None, help="資料根目錄；省略時同 init_batch.py（local_config.json → 內建預設值）")
    ap.add_argument("--show-text", action="store_true", help="在終端機列出每個小項的評分標準文字，方便人工比對語意")
    args = ap.parse_args()
    now = dt.datetime.now()

    if not batch_dir(args).is_dir():
        sys.exit(f"找不到批次資料夾：{batch_dir(args)}。請先執行 init_batch.py --batch {args.batch}")

    rep = Report()
    names = asset_names()
    rubric = pick_rubric(args, names, rep)
    items = []
    if rubric is not None:
        rubric_mod = load_module("extract_rubric", ENGINE_DIR / "extract_rubric.py")
        es = load_module("extract_submission", ENGINE_DIR / "extract_submission.py")
        ss = module_constants(ENGINE_DIR / "score_submission.py")
        md = SCORING_CHECKS.read_text(encoding="utf-8")
        items = rubric_mod.read_rubric_file(rubric)
        crit = {it["code"]: str(it["criteria"] or "") for it in items}

        check_points(items, rubric_mod.builtin_rubric_map(), md_table(md), md, rep)
        check_values(crit, es, ss, md, rep)
        check_answer(batch_dir(args) / "01_template", names, items, rep)

        if args.show_text:
            print("\n評分標準文字（語意請人工對照 scoring_checks.md）")
            for it in items:
                print(f"\n  {it['code']} {it['title']}（{it['points']:g}／額外 {it['bonus']:g}）"
                      f"\n    {crit[it['code']]}")

    print(f"\n{'全部通過' if not rep.failed else f'有 {rep.failed} 項未通過'}"
          "（評分標準文字的語意仍需人工對照 scoring_checks.md）", flush=True)

    stem = VERSION_RE.sub("", names["rubric"])     # S06_excel_scoring_criteria
    out = batch_dir(args) / "02_work" / "rubric" / f"{stem}_{now:%Y%m%d_%H%M%S}.md"
    write_report(out, rep, rubric, items, now)
    print(f"報告：{out}", flush=True)
    sys.exit(1 if rep.failed else 0)


if __name__ == "__main__":
    main()
