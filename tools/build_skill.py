#!/usr/bin/env python3
"""
build_skill.py — package one skill folder as a portable .skill (zip) for
Claude.ai upload or copying to another machine.

Left out of the package: .venv/, __pycache__/, *.pyc, Excel lock files (~$*),
scripts/local_config.json (per-machine paths), .gitignore, and anything in
assets/02_work/ except .gitkeep.

Usage (from the repo root):
    python tools/build_skill.py 04 --name excel_grading_v1
Output: dist/<name>.skill (default name = the SKILL.md name). --name renames
only the file; the folder inside keeps the SKILL.md name Claude shows.
"""
import re
import sys
import argparse
import zipfile
from pathlib import Path, PurePosixPath

REPO = Path(__file__).resolve().parents[1]
SKILLS_DIR = REPO / "skills"
OUT_DIR = REPO / "dist"
SKIP_DIRS = {".venv", "__pycache__", ".git"}
SKIP_FILES = {".gitignore", "local_config.json"}


def find_skill(arg: str) -> Path:
    """skills/ folder by number prefix (04) or full name (04-excel-grading)."""
    matches = [d for d in SKILLS_DIR.iterdir() if d.is_dir() and (d.name == arg or d.name.startswith(f"{arg}-"))]
    if len(matches) != 1:
        sys.exit(f"找不到唯一的 skill「{arg}」：{', '.join(d.name for d in matches) or '無符合的資料夾'}")
    if not (matches[0] / "SKILL.md").is_file():
        sys.exit(f"{matches[0].name} 缺 SKILL.md，無法打包")
    return matches[0]


def skill_name(skill_dir: Path) -> str:
    text = (skill_dir / "SKILL.md").read_text(encoding="utf-8")
    m = re.search(r"^---\s*\n(.*?)\n---", text, re.S)
    name = m and re.search(r"^name:\s*(\S+)\s*$", m.group(1), re.M)
    if not name:
        sys.exit(f"{skill_dir / 'SKILL.md'} 開頭沒有 name: 欄位")
    return name.group(1)


def skill_files(skill_dir: Path):
    for p in sorted(skill_dir.rglob("*")):
        if not p.is_file():
            continue
        rel = p.relative_to(skill_dir)
        if set(rel.parts[:-1]) & SKIP_DIRS or p.name in SKIP_FILES:
            continue
        if p.suffix == ".pyc" or p.name.startswith("~$"):
            continue
        if rel.parts[:2] == ("assets", "02_work") and p.name != ".gitkeep":
            continue
        yield p, rel


def main():
    # 1. 主控台輸出改用 UTF-8，避免 Windows（cp950）印中文路徑出錯
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    # 2. 讀取參數：skill（必填）、--name（選填）
    ap = argparse.ArgumentParser(description="Package one skill folder as a portable .skill zip.")
    ap.add_argument("skill", help="skills/ 下的資料夾編號或名稱，例如 04")
    ap.add_argument("--name", help="自訂 .skill 檔名（不含副檔名），例如 excel_grading_v1")
    args = ap.parse_args()

    # 3. 找到 skills/ 下唯一符合的資料夾，且必須有 SKILL.md
    skill_dir = find_skill(args.skill)

    # 4. 從 SKILL.md 開頭讀出 name，作為 zip 內的資料夾名稱
    name = skill_name(skill_dir)

    # 5. 決定輸出檔名並檢查合法字元，輸出到 dist/<檔名>.skill
    file_name = (args.name or name).removesuffix(".skill")
    if not file_name.strip() or set(file_name) & set('<>:"/\\|?*'):
        sys.exit(f"打包名稱「{file_name}」不可為空或含有 < > : \" / \\ | ? *")
    out = OUT_DIR / f"{file_name}.skill"
    out.parent.mkdir(parents=True, exist_ok=True)
    root = PurePosixPath(name)

    # 6. 逐一寫入 skill 檔案（排除 venv、快取、本機設定、02_work 作業檔）
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for p, rel in skill_files(skill_dir):
            zf.write(p, str(root / PurePosixPath(*rel.parts)))
            n += 1

    # 7. 印出輸出路徑與打包的檔案數
    print(f"{out}\n  skill 檔案 {n} 個")


if __name__ == "__main__":
    main()
