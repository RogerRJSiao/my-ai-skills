#!/usr/bin/env python3
"""
init_batch.py — set up one grading batch folder from this skill's assets/.

Copies the whole assets/ tree (00_input / 01_template / 02_work / 03_results)
into <data root>/grading/<domain>/<課程代碼>_<梯次>/, e.g.
D:\\claude_data\\grading\\training_internal\\S06_202609\\, so every template lands
in the folder it is used from. domain and 課程代碼 come from skill.toml.
Folders git does not keep (empty ones such as 00_input/submissions) are
created as well, and so are any missing parents such as grading/<domain>/.
Existing files are never overwritten (--force is disabled), so it is safe
to re-run on a batch that already holds rosters, submissions or results.
.gitkeep placeholders are not copied.

The data root (shared by every skill) is taken from, in order:
  1. --root
  2. scripts/local_config.json (per-machine, not in git):
       {"data_root": "D:\\\\claude_data"}
  3. DEFAULT_DATA_ROOT below

Usage:
    python init_batch.py --batch 202609
    python init_batch.py --root D:\\claude_data --batch 202609
    python init_batch.py --batch 202609 --dry-run
"""
import sys
import json
import shutil
import argparse
import tomllib
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
ASSETS = SKILL_DIR / "assets"
LOCAL_CONFIG = Path(__file__).resolve().parent / "local_config.json"
DEFAULT_DATA_ROOT = r"D:\claude_data"   # used when there is no local_config.json
GRADING_DIR = "grading"                  # <data root>/grading/<domain>/<課程>_<梯次>/

# The batch layout agreed for training_internal/<課程>_<梯次>/. Created even
# when the folder is empty in assets/ (git does not track empty folders).
LAYOUT = [
    "00_input/manifest_source",
    "00_input/roster",
    "00_input/submissions",
    "01_template",
    "02_work/Excel",
    "02_work/Word",
    "02_work/PPT",
    "03_results",
]
SKIP_NAMES = {".gitkeep"}


def skill_meta() -> dict:
    return tomllib.loads((SKILL_DIR / "skill.toml").read_text(encoding="utf-8")).get("skill", {})


def course_code() -> str:
    code = skill_meta().get("course_code")
    if not code:
        sys.exit("skill.toml 沒有設定 [skill] course_code，請用 --course 指定。")
    return code


def domain() -> str:
    d = skill_meta().get("domain")
    if not d:
        sys.exit("skill.toml 沒有設定 [skill] domain（例如 training_internal）。")
    return d


def default_root() -> tuple[str, str]:
    """(data root, where it came from). A config file that exists but is
    broken is an error rather than a silent fall-back to the default."""
    if not LOCAL_CONFIG.is_file():
        return DEFAULT_DATA_ROOT, "init_batch.py 內建預設值"
    try:
        root = json.loads(LOCAL_CONFIG.read_text(encoding="utf-8-sig")).get("data_root")
    except (json.JSONDecodeError, AttributeError) as e:
        sys.exit(f"{LOCAL_CONFIG} 不是有效的 JSON 物件：{e}")
    if not root:
        sys.exit(f"{LOCAL_CONFIG} 沒有設定 data_root。")
    return root, LOCAL_CONFIG.name


def plan(dest: Path, force: bool):
    """Return (dirs to create, [(src, dst, action)]) without touching disk."""
    dirs = {dest / d for d in LAYOUT}
    dirs |= {dest / p.relative_to(ASSETS) for p in ASSETS.rglob("*") if p.is_dir()}
    files = []
    for src in sorted(p for p in ASSETS.rglob("*") if p.is_file()):
        if src.name in SKIP_NAMES or src.name.startswith("~$"):
            continue
        dst = dest / src.relative_to(ASSETS)
        if not dst.exists():
            action = "copy"
        elif force:
            action = "overwrite"
        else:
            action = "skip (exists)"
        files.append((src, dst, action))
    return sorted(dirs), files


def main():
    #--設定CLI
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")   # sys.exit messages
    ap = argparse.ArgumentParser(description="Create a grading batch folder from assets/.")
    ap.add_argument("--root", default=None,
                    help="資料根目錄，例如 D:\\claude_data（批次建在其下的 grading\\<domain>\\，不存在會自動建立）；"
                         "省略時讀 scripts/local_config.json 的 data_root，沒有設定檔則用內建預設值")
    ap.add_argument("--batch", required=True, help="梯次(yyyymm)，例如 202609")
    ap.add_argument("--course", default=None, help="課程代碼，預設讀 skill.toml 的 course_code")
    ap.add_argument("--dry-run", action="store_true", help="只列出會做的事，不建立或複製任何檔案")
    # ap.add_argument("--force", action="store_true", help="覆寫批次資料夾中已存在的同名檔案")
    args = ap.parse_args()
    #--確立data來源的根目錄
    root_str, root_source = (args.root, "--root") if args.root else default_root()
    root = Path(root_str)
    if root.exists() and not root.is_dir():
        sys.exit(f"資料根目錄的路徑是檔案，不是資料夾：{root}")
    domain_dir = root / GRADING_DIR / domain()
    dest = domain_dir / f"{args.course or course_code()}_{args.batch}"
    dirs, files = plan(dest, force=False)   # --force 已停用：已存在的檔案一律不覆寫
    #--建立目錄架構(若已存在就不重建)
    print(f"資料根目錄：{root}（來源：{root_source}）")
    print(f"{'[dry-run] ' if args.dry_run else ''}{ASSETS} -> {dest}")
    if not domain_dir.is_dir():
        print(f"  root  create        {domain_dir}")
        if not args.dry_run:
            domain_dir.mkdir(parents=True, exist_ok=True)
    for d in dirs:
        state = "exists" if d.is_dir() else "create"
        print(f"  dir   {state:<13} {d.relative_to(dest)}")
        if state == "create" and not args.dry_run:
            d.mkdir(parents=True, exist_ok=True)
    for src, dst, action in files:
        print(f"  file  {action:<13} {dst.relative_to(dest)}")
        if action != "skip (exists)" and not args.dry_run:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)

    copied = sum(a != "skip (exists)" for _, _, a in files)
    skipped = len(files) - copied
    verb = "would copy" if args.dry_run else "copied"
    print(f"{verb} {copied} file(s), skipped {skipped} existing file(s)")


if __name__ == "__main__":
    main()
