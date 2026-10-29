#!/usr/bin/env python3
"""
copy_to_work.py — copy the correctly filed submissions into the batch's work
area, ready for grading. By default all three file types are handled; --type
picks just one.

    00_input/submissions/<單位-工號-姓名>/<單位-工號-姓名>.<ext>
        -> 02_work/<Excel|Word|PPT>/<單位-工號-姓名>.<ext>

A file is copied only when all of these hold:
  1. it sits exactly one folder below submissions/
  2. that folder is named 單位-6碼工號-姓名
  3. the file is named 單位-6碼工號-姓名.<ext> (xlsx, docx or pptx)
  4. the folder's 單位-工號-姓名 is one row (單位, 工號, 姓名) of the roster
     00_input/roster/<課程>_三科_學員名冊<梯次>.xlsx, exactly as written there
     (單位 case included: p5c00 is not P5C00)
  5. the file name uses the same 單位-工號-姓名 as its folder
姓名 may not contain "-", "." or whitespace. Everything else is listed with
the reason it was skipped. Sources are only read; existing files in 02_work
are never overwritten. Archives (.zip/.7z/.rar) are never extracted: they are
listed as skipped, and the student must re-submit the file itself.

Usage:
    python copy_to_work.py --batch 202701                  # xlsx + docx + pptx
    python copy_to_work.py --batch 202701 --type docx --dry-run
    python copy_to_work.py --root D:\\claude_data --batch 202701 --type xlsx
"""
import re
import sys
import shutil
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl                                                         # noqa: E402
from init_batch import GRADING_DIR, course_code, default_root, domain   # noqa: E402

# file type -> 02_work sub-folder (same names as check_batch.WORK_SUBDIRS)
WORK_DIR_BY_TYPE = {"xlsx": "Excel", "docx": "Word", "pptx": "PPT"}
ARCHIVE_EXTS = {".zip", ".7z", ".rar"}           # listed, never extracted
STEM = r"[A-Za-z0-9]+-\d{6}-[^-.\s]+"           # 單位-6碼工號-姓名
FOLDER_RE = re.compile(rf"{STEM}")
ROSTER_HEADERS = ["id", "姓名", "工號", "單位"]


def roster_keys(roster_dir: Path, course: str, batch: str) -> tuple[Path, set[str]]:
    """The batch roster and its 單位-工號-姓名 combinations, exactly as written there."""
    name = f"{course}_三科_學員名冊{batch}.xlsx"
    path = roster_dir / name
    if not path.is_file():
        sys.exit(f"找不到學員名冊 {path}，無法判斷哪些檔案合規")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = wb.worksheets[0].iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows, ())]
    if header[:4] != ROSTER_HEADERS:
        sys.exit(f"{name} 的表頭應為 {ROSTER_HEADERS}，實際為 {header[:4]}")
    keys = {f"{str(r[3]).strip()}-{str(r[2]).strip()}-{str(r[1]).strip()}"
            for r in rows if r and all(r[i] not in (None, "") for i in (1, 2, 3))}
    wb.close()
    if not keys:
        sys.exit(f"{name} 沒有學員資料")
    return path, keys


def plan(submissions: Path, work: Path, ext: str, roster: set[str]):
    """[(src, dst or None, action)] for every file of type ext under submissions/."""
    file_re = re.compile(rf"{STEM}\.{re.escape(ext)}", re.IGNORECASE)
    out, taken = [], set()
    for src in sorted(p for p in submissions.rglob("*") if p.is_file()):
        if src.suffix.lower() != f".{ext}" or src.name.startswith("~$"):
            continue
        rel = src.relative_to(submissions)
        if len(rel.parts) != 2:
            out.append((src, None, "skip: 不在 submissions/<單位-工號-姓名>/ 這一層"))
        elif not FOLDER_RE.fullmatch(rel.parts[0]):
            out.append((src, None, "skip: 資料夾名稱不是 單位-工號-姓名"))
        elif not file_re.fullmatch(src.name):
            out.append((src, None, "skip: 檔名不是 單位-工號-姓名." + ext))
        elif rel.parts[0] not in roster:
            out.append((src, None, "skip: 資料夾的 單位-工號-姓名 不在學員名冊"))
        elif src.stem != rel.parts[0]:
            out.append((src, None, "skip: 檔名與資料夾的 單位-工號-姓名 不一致"))
        elif (work / src.name).exists() or src.name.lower() in taken:
            out.append((src, work / src.name, "skip: 02_work 已有同名檔案"))
        else:
            taken.add(src.name.lower())
            out.append((src, work / src.name, "copy"))
    return out


def archives(submissions: Path) -> list[Path]:
    """Archives anywhere under submissions/: listed once, never extracted."""
    return sorted(p for p in submissions.rglob("*") if p.is_file() and p.suffix.lower() in ARCHIVE_EXTS)


def main():
    #--設定CLI
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="Copy correctly filed submissions into 02_work for grading.")
    ap.add_argument("--root", default=None,
                    help="資料根目錄，例如 D:\\claude_data；省略時同 init_batch.py（local_config.json → 內建預設值）")
    ap.add_argument("--batch", required=True, help="梯次(yyyymm)，例如 202701")
    ap.add_argument("--course", default=None, help="課程代碼，預設讀 skill.toml 的 course_code")
    ap.add_argument("--type", default="all", choices=["all", *WORK_DIR_BY_TYPE],
                    help="檔案種類：xlsx→02_work/Excel、docx→Word、pptx→PPT；預設 all（三種都執行）")
    ap.add_argument("--dry-run", action="store_true", help="只列出會做的事，不複製任何檔案")
    args = ap.parse_args()

    #--找到梯次資料夾
    root_str, root_source = (args.root, "--root") if args.root else default_root()
    course = args.course or course_code()
    dest = Path(root_str) / GRADING_DIR / domain() / f"{course}_{args.batch}"
    submissions = dest / "00_input" / "submissions"
    types = list(WORK_DIR_BY_TYPE) if args.type == "all" else [args.type]
    print(f"梯次：{dest}（資料根目錄來源：{root_source}）")
    if not submissions.is_dir():
        sys.exit(f"找不到 {submissions}。請先執行 init_batch.py --batch {args.batch} 並放入學員作業")
    for t in types:   # check every target before copying anything
        if not (dest / "02_work" / WORK_DIR_BY_TYPE[t]).is_dir():
            sys.exit(f"找不到 {dest / '02_work' / WORK_DIR_BY_TYPE[t]}。請先執行 init_batch.py --batch {args.batch}")
    roster_path, roster = roster_keys(dest / "00_input" / "roster", course, args.batch)
    print(f"學員名冊：{roster_path.name}（{len(roster)} 人）")
    verb = "would copy" if args.dry_run else "copied"

    #--壓縮檔只列出一次，不解壓縮
    for a in archives(submissions):
        print(f"  skip: 壓縮檔，不解壓縮 {a.relative_to(submissions)}")

    total = 0
    for t in types:
        work = dest / "02_work" / WORK_DIR_BY_TYPE[t]
        #--逐一判斷並複製（來源只讀，02_work 已有的檔案不覆寫）
        print(f"\n{'[dry-run] ' if args.dry_run else ''}{submissions.relative_to(dest)} -> {work.relative_to(dest)}（{t}）")
        items = plan(submissions, work, t, roster)
        for src, dst, action in items:
            print(f"  {action:<8} {src.relative_to(submissions)}")
            if action == "copy" and not args.dry_run:
                shutil.copy2(src, dst)

        #--統計
        copied = sum(a == "copy" for _, _, a in items)
        total += copied
        print(f"{verb} {copied} file(s), skipped {len(items) - copied} file(s)")

        #--提醒 02_work 已有、但不在本次合規清單的檔案（只提醒，不刪除）
        ok = {src.name.lower() for src, _, a in items
              if a == "copy" or a == "skip: 02_work 已有同名檔案"}
        stray = sorted(p.name for p in work.iterdir() if p.is_file() and p.name.lower() not in ok)
        if stray:
            print(f"注意：02_work/{work.name} 有 {len(stray)} 份不合規的檔案（不會自動刪除）：{', '.join(stray)}")

    if len(types) > 1:
        print(f"\n合計 {verb} {total} file(s)")


if __name__ == "__main__":
    main()
