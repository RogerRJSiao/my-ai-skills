# kd-mcp — my-ai-skills 的 MCP Servers

把 `skills/` 裡需要「確定性執行」的部分（讀檔、計分、寫結果檔）包成 MCP tools，
讓 Claude Code / Claude Desktop 等 MCP 客戶端直接呼叫，不必每次讓模型重新推導
解析與計分邏輯。

一個套件裡有多個 server，依 **domain** 區分，各自部署到不同主機：

| Server | Domain | Tool 前綴 | 預設 port | 部署主機 | 目前的 skills |
|---|---|---|---|---|---|
| `kd-grading-internal` | `training_internal`（內訓） | `gi_` | 8101 | 主機 A | 04-excel-grading（05-word-grading、06-ppt-grading 規劃中） |
| `kd-grading-external` | `training_external`（外訓） | `ge_` | 8102 | 主機 B | （07-powerbi-grading 規劃中） |

## 安裝

```bash
cd mcp-server
python -m venv .venv
.venv/Scripts/python -m pip install -e ".[office]"   # macOS/Linux：.venv/bin/python
```

需求：Python ≥ 3.11、`mcp` 2.x。

Extras 對應評分 engine 需要的套件：

| Extra | 內容 |
|---|---|
| `office` | `openpyxl`（`xlsx_s06` engine 需要） |

如果沒裝某個 engine 需要的 extra，server 仍會啟動，但 `*_list_skills` 會把該 skill 標成 `engine_available: false`。

## 啟動

```bash
python -m kd_mcp.servers.grading_internal                     # stdio（給 Claude Code / Desktop）
python -m kd_mcp.servers.grading_internal --transport streamable-http --port 8101
#   → http://127.0.0.1:8101/mcp；放在 IIS 後面時加 --allowed-host <網站主機名稱>
kd-grading-internal --help                                    # 安裝後的 entry point，參數相同
```

在本 repo 開發時，根目錄的 `.mcp.json` 會讓 Claude Code 自動載入 `kd-grading-internal`。
它指向 `mcp-server/.venv/Scripts/python.exe`，這是 Windows 路徑；macOS/Linux 請改成 `.venv/bin/python`。


## Tools（grading toolkit，以 `gi_` 為例）

所有 server 共用同一組 tools，只有前綴不同。每個 tool 都要帶 `skill` 參數。
傳入不屬於該 server domain 的 skill 會被拒絕。

| Tool | 用途 | 會寫檔？ |
|---|---|---|
| `gi_list_skills` | 列出本 server 可評分的 skill、接受的格式、預設課程代碼／實作項目、engine 是否可用；`issues` 列出 skill.toml 的問題 | 否 |
| `gi_get_rubric` | 讀評閱標準；`diffs_from_scorer` 非空代表配分已異動 | 否 |
| `gi_list_submissions` | 列出資料夾內符合該 skill 格式的檔案，並解析「部門代號-工號-姓名」；`filename_valid=false` 需向使用者確認 | 否 |
| `gi_inspect_submission` | 萃取客觀事實（04：樞紐表、圖表、凍結窗格、框線、排序/篩選…） | 否 |
| `gi_grade_submission` | 計分：各小項得分＋評閱結果說明、`needs_review`、`warnings`、`draft_summary` | 否 |
| `gi_write_results` | 批改並寫入『單科評分結果』活頁簿（由 skill 的 `engine/results.py` 寫入）；帶 `batch` 且梯次資料夾存在時存到 `<課程代碼>_<梯次>/03_results/`，否則存到 output，可追加到既有結果檔；`meta_sources` 標示課程代碼／實作項目的來源 | 是 |

建議流程：

1. `gi_list_skills`、`gi_list_submissions`（04 請指向梯次的 `02_work/Excel/`，例如 `S06_202609/02_work/Excel`）。
2. 每個檔案呼叫 `gi_grade_submission`。
3. 處理 `needs_review`：人工確認後用 `overrides` 帶入。
4. 撰寫 150 字簡述。
5. `gi_write_results`。

建立梯次、複製到 02_work、三科彙總沒有對應的 tool，請在終端機執行 skill 的
`scripts/`（`init_batch.py`、`check_batch.py`、`copy_to_work.py`、`build_all_report.py`）。

### 04-excel-grading（engine `xlsx_s06`）計分重點

- **規則來源：** 標準是評閱標準 `S06_excel_scoring_criteria_v<最大版本>.xlsx`；
  engine 依照據此整理的 `skills/04-excel-grading/references/scoring_checks.md` 計分。
  三者不一致時以 xlsx 為準，修改 md 與 engine。包含這些「連坐」規則：
  - 沒有以【薪資】A1:J13 為來源的樞紐表（可放在其他活頁，排除【考題】範例）→ 1.x 全部 0 分。
  - 原始資料被改（車輛銷售 A2:A6 與 A2:E2、武將能力 A2:A7 與 A2:E2）→ 3.x／4.x 全部 0 分。
  - 圖表非動態或類型不對（3.x 只限長條圖／直條圖／直方圖，4.x 只限雷達圖）→ 3.x／4.x 全部 0 分。
- **不自動給分、改標 `needs_review` 的情況：**
  - 樞紐表來源解析不出是哪個活頁與範圍（1.x）。
  - 樞紐輸出區找不到數值格，無法判讀格式（1.4）。
  - 凍結窗格設在 B3／A3／B1／A1 以外的位置（2.1）。
  - 檔案沒有計算結果，又沒有個人平均由大到小的排序條件（2.4）。
  - 沒有提供題目檔，無法檢查原始資料（3.1、4.1）。
  - 圖表用「資料表(Data Table)」而不是資料標籤（3.4）。
  - 雷達圖 `radarStyle=marker` 但未指定標記符號（4.3）。
  - 【車輛銷售】只有直方圖等新式圖表（chartEx），程式無法判讀（3.x）。
- **`分數加總` 用 `SUMIFS(工號, 姓名)`：** 已知案例中有兩位學生共用工號 242000，只靠姓名或只靠工號都可能加錯人。
- **寫檔保護：** 不覆寫既有輸出（除非 `overwrite=true`）。同一學生已在結果檔時預設報錯，可用 `on_duplicate="replace"` 取代舊資料。

## 設定與安全

| 環境變數 | CLI 參數 | 說明 | 預設 |
|---|---|---|---|
| `KD_MCP_DATA_ROOT` | `--data-root` | 學生檔案的根目錄；tool 的相對路徑以它為基準 | 目前工作目錄 |
| `KD_MCP_OUTPUT_DIR` | `--output-dir` | 可寫入的目錄（另外各梯次的 `03_results/` 也可寫入） | `<data_root>/outputs` |
| `KD_MCP_SKILLS_DIR` | `--skills-dir` | skills 目錄（部署時指向複製過去的 skills） | repo 的 `skills/` |
| `KD_MCP_ALLOWED_DIRS` | — | 額外的唯讀目錄（Windows 用 `;` 分隔） | 無 |
| `KD_MCP_AUDIT_LOG` | — | 稽核 log 檔；設為空字串則關閉 | `<output_dir>/_audit/<server>.jsonl` |

- **讀檔：** 只允許 data_root、skills、output 和額外唯讀目錄。
- **寫檔：** 只允許 output，以及 data_root 底下各梯次的 `03_results/`（`00_input`、`01_template`、`02_work` 不可寫）。
- **路徑解析：** 檢查前會先解析 `..` 與 symlink。
- **稽核 log：** 每次 tool 呼叫記一行 JSON（時間、server、tool、user、狀態、耗時）。只記檔名，不記內容、簡述或 overrides。
- **HTTP 模式：** 只綁 127.0.0.1，並只接受白名單內的 Host 標頭。目前**沒有使用者驗證**，見 Roadmap。

## 架構

```
mcp-server/
├─ pyproject.toml                  # extras: office；每個 server 一個 entry point
├─ src/kd_mcp/
│  ├─ core/                        # ① 不依賴 MCP
│  │  ├─ config.py                 #   Settings（環境變數 / CLI）
│  │  ├─ security.py               #   PathPolicy：讀／寫白名單
│  │  ├─ logging.py                #   稽核 log（JSON Lines）
│  │  ├─ registry.py               #   掃描 skills/*/skill.toml，依 domain 篩選
│  │  └─ grading/
│  │     ├─ models.py              #   ItemScore/GradeResult、檔名解析、overrides、簡述草稿
│  │     └─ engines/               #   依 skill.toml 的 engine_path 載入 skill 的 engine/
│  │        ├─ base.py             #   Engine 介面：inspect / score / rubric / results
│  │        └─ __init__.py         #   依 engine_path 從 skill 資料夾載入
│  │                               #   （04-excel-grading 的 xlsx_s06 放在 skills/04-excel-grading/engine/）
│  ├─ runtime/                     # ② 啟動與傳輸
│  │  ├─ transport.py              #   CLI、stdio / streamable-http、Host 白名單
│  │  └─ tooling.py                #   tool 包裝：稽核 log＋可預期錯誤轉 ToolError
│  ├─ toolkits/                    # ③ 可重用的 tool 組
│  │  └─ grading.py                #   register(mcp, *, domain, prefix, data_root, output, ...)
│  └─ servers/                     # ④ 實際啟動的 server
│     ├─ grading_internal/         #   kd-grading-internal（內訓）→ 主機 A
│     └─ grading_external/         #   kd-grading-external（外訓）→ 主機 B
```

相依方向：`servers → toolkits → core`、`runtime → core`。`core/` 不 import `mcp`。

### skill.toml

每個要給 MCP 用的 skill，在資料夾根目錄放一個 `skill.toml`。沒有 `skill.toml` 的 skill 會被略過。

```toml
[skill]
id = "04-excel-grading"        # 必須等於資料夾名稱
title = "S06 Excel實作 評分"
domain = "training_internal"   # 決定由哪個 server 提供
format = "xlsx"                # 多種格式用 | 分隔，例如 "pbix|pbip"
course_code = "S06"            # 寫結果檔的預設值（選填）
practice_item = "Excel實作"     # 同上（選填）

[grading]
engine = "xlsx_s06"            # engine 名稱
engine_path = "engine"         # 必填：skill 資料夾內的 engine 套件
pass_score = 60

[assets]                       # 相對於 skill 資料夾，engine 依名稱取用
rubric = "assets/..."
question = "assets/..."
result_template = "assets/..."
```

### 新增一個評分 skill（例如 05-word-grading）

1. 在 `skills/05-word-grading/` 寫好 `SKILL.md` 和 assets，再加一份 `skill.toml`（`domain`、`format = "docx"`、`engine`、`engine_path`）。
2. 在 skill 資料夾新增 `engine/` 套件，並在 `skill.toml` 設 `engine_path = "engine"`：
   - 實作 `base.Engine`，在 `__init__.py` 提供 `ENGINE`；共用的資料結構從 `kd_mcp.core.grading.models` import。
   - `ENGINE.results` 是寫單科評分結果活頁簿的模組（`Meta`、`default_output_name`、`meta_from_existing`、`write_results`）；
     格式和 04 相同時，可以複製 `skills/04-excel-grading/engine/results.py`。
   - 需要的第三方套件加進 `pyproject.toml` 的 extras。
3. 檔名解析、overrides、簡述草稿，直接沿用 `core/grading/`。

只要 domain 已經有對應的 server，就**不用改任何 server 或 toolkit 程式**。

### 新增一個 server（新 domain）

複製 `servers/grading_internal/`，改 `NAME`、`DOMAIN`、`PREFIX`、`DEFAULT_PORT`，
並在 `pyproject.toml` 加 entry point。

## Roadmap（尚未實作）

| 模組 | 用途 | 何時做 |
|---|---|---|
| `runtime/auth.py` | HTTP token → 使用者 ID，寫入稽核 log 的 `user` 欄位 | IIS 對外開放前（目前只綁 127.0.0.1） |
| `core/masking.py` | log 與回傳內容的姓名／工號遮罩 | 外訓 server 上線前 |
| `core/jobs.py` + `toolkits/jobs.py` | 批次評分佇列：job_submit / job_status | 一次評分量超過 MCP 逾時時 |
| `core/skill_runner.py` | 以子行程執行 skill 自帶的 scripts | 出現只有 script、沒有 engine 的 skill 時 |
| 05–07 各自的 `engine/`（docx／pptx／pbix），extras `powerbi` / `http` | 05–07 評分 | 05–07 開發時 |
