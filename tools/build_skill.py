#!/usr/bin/env python3
"""
build_skill.py — package one skill folder as a portable .skill (zip) for
Claude.ai upload or copying to another machine.

Left out of the package: .venv/, __pycache__/, *.pyc, Excel lock files (~$*),
scripts/local_config.json (per-machine paths), .gitignore, and anything in
assets/02_work/ except .gitkeep.

After packaging, the same files (plus the skill's .gitignore) are also
copied into another folder, e.g. a separate git repo that versions the
skill. The destination is, in order:
  1. --copy-to
  2. tools/local_config.json (per-machine, not in git), keyed by the skill folder:
       {"copy_to": {"04-excel-grading": "E:\\\\other\\\\excel_eval"}}
  3. the skill's skill.toml (the default, in git):
       [publish]
       copy_to = 'D:\\【git_rjsiao】\\git_skill\\s06_evaluation\\excel_eval'
Changed files are overwritten, nothing in the destination is deleted (files
that are only there are listed), and its parent folder must already exist:
for --copy-to a missing parent is an error, for the defaults (2, 3) the copy
is skipped with a note, so CI and other machines still just package.
--no-copy skips the copy.

Usage (from the repo root):
    python tools/build_skill.py 04 --name excel_grading_v1
    python tools/build_skill.py 04 --name excel_grading_v1 --copy-to D:\\path\\excel_eval
    python tools/build_skill.py 04 --no-copy
Output: dist/<name>.skill (default name = the SKILL.md name). --name renames
only the file; the folder inside keeps the SKILL.md name Claude shows.
"""
import re
import sys
import json
import shutil
import tomllib
import argparse
import zipfile
from pathlib import Path, PurePosixPath

REPO = Path(__file__).resolve().parents[1]
SKILLS_DIR = REPO / "skills"
OUT_DIR = REPO / "dist"
LOCAL_CONFIG = Path(__file__).resolve().parent / "local_config.json"
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


def skill_files(skill_dir: Path, keep_gitignore: bool = False):
    for p in sorted(skill_dir.rglob("*")):
        if not p.is_file():
            continue
        rel = p.relative_to(skill_dir)
        if set(rel.parts[:-1]) & SKIP_DIRS:
            continue
        if p.name in SKIP_FILES and not (keep_gitignore and p.name == ".gitignore"):
            continue
        if p.suffix == ".pyc" or p.name.startswith("~$"):
            continue
        if rel.parts[:2] == ("assets", "02_work") and p.name != ".gitkeep":
            continue
        yield p, rel


def copy_destination(skill_dir: Path, arg: str | None) -> tuple[Path | None, str]:
    """(destination, where it came from): --copy-to, else tools/local_config.json
    copy_to[<skill folder name>], else the skill's skill.toml [publish] copy_to."""
    if arg:
        return Path(arg), "--copy-to"
    if LOCAL_CONFIG.is_file():
        try:
            config = json.loads(LOCAL_CONFIG.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError as e:
            sys.exit(f"{LOCAL_CONFIG} 不是有效的 JSON：{e}")
        dest = (config.get("copy_to") or {}).get(skill_dir.name)
        if dest:
            return Path(dest), "tools/local_config.json"
    toml = skill_dir / "skill.toml"
    if toml.is_file():
        dest = tomllib.loads(toml.read_text(encoding="utf-8")).get("publish", {}).get("copy_to")
        if dest:
            return Path(dest), f"{skill_dir.name}/skill.toml"
    return None, ""


def copy_skill(skill_dir: Path, dest: Path) -> tuple[int, int, list[str]]:
    """Copy the packaged files (plus .gitignore) into dest. Returns
    (copied, unchanged, files that exist only in dest). Never deletes."""
    dest = dest.resolve()
    if not dest.parent.is_dir():
        sys.exit(f"複製目的地的上層資料夾不存在：{dest.parent}")
    if dest == REPO or REPO in dest.parents:
        sys.exit(f"複製目的地不可在本 repo 裡：{dest}")
    copied = unchanged = 0
    sources = set()
    for p, rel in skill_files(skill_dir, keep_gitignore=True):
        sources.add(rel)
        target = dest / rel
        if target.is_file() and target.read_bytes() == p.read_bytes():
            unchanged += 1
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, target)
        copied += 1
    extra = sorted(str(t.relative_to(dest)) for t in dest.rglob("*")
                   if t.is_file() and ".git" not in t.relative_to(dest).parts
                   and t.relative_to(dest) not in sources)
    return copied, unchanged, extra


def main():
    # 1. 主控台輸出改用 UTF-8，避免 Windows（cp950）印中文路徑出錯
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")   # sys.exit messages

    # 2. 讀取參數：skill（必填）、--name、--copy-to、--no-copy（選填）
    ap = argparse.ArgumentParser(description="Package one skill folder as a portable .skill zip.")
    ap.add_argument("skill", help="skills/ 下的資料夾編號或名稱，例如 04")
    ap.add_argument("--name", help="自訂 .skill 檔名（不含副檔名），例如 excel_grading_v1")
    ap.add_argument("--copy-to", help="打包後把 skill 檔案複製到這個資料夾；省略時依序讀 tools/local_config.json、"
                                      "skill.toml [publish] 的 copy_to")
    ap.add_argument("--no-copy", action="store_true", help="只打包，不複製")
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

    # 8. 有複製目的地時（--copy-to、tools/local_config.json、skill.toml [publish]），把同一批檔案（含 .gitignore）複製過去
    dest, source = (None, "") if args.no_copy else copy_destination(skill_dir, args.copy_to)
    # 預設目的地（不是 --copy-to）在這台機器上不存在時（例如 CI、別台電腦）只打包、不報錯
    if dest and source != "--copy-to" and not (dest.is_absolute() and dest.parent.is_dir()):
        print(f"略過複製：{source} 的目的地 {dest} 在這台機器上不存在")
        dest = None
    if dest:
        copied, unchanged, extra = copy_skill(skill_dir, dest)
        print(f"複製到 {dest.resolve()}（來源設定：{source}）\n  更新 {copied} 個、未變動 {unchanged} 個")
        if extra:
            print(f"  目的地多出 {len(extra)} 個來源沒有的檔案（未刪除）：{', '.join(extra)}")


if __name__ == "__main__":
    main()
