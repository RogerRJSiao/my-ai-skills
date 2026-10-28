---
name: s06-excel-grading
description: Standard operating procedure for grading "S06" Excel實作 (pivot table + charts) student submissions against the bundled scoring criteria (S06_excel_scoring_criteria_v*.xlsx), question and answer files, and writing results into the 單科評分結果 workbook (活頁1 評分細項 + 活頁2 評分結果表). Also covers the batch-folder tools: setting up a 梯次 folder (init_batch.py), checking it before grading (check_batch.py), copying compliant submissions into 02_work (copy_to_work.py) and building the 三科 學員檔案與成績計算 workbook (build_all_report.py). Use this whenever the user uploads one or more student .xlsx submissions named 「部門代號-工號-姓名.xlsx」 and asks to grade/評分/批改 an "S06" or "Excel實作" assignment, wants scores added to an existing 評分結果 file, asks to 建立梯次/檢查梯次/複製到 02_work/三科彙總, or references 評閱標準 / S06_excel_scoring_criteria / S06_excel_answer / S06_excel_question. Trigger this even if the user just says "幫我評分這個檔案" in a conversation that has already established the S06 Excel grading context.
---

# S06 Excel實作 評分 SOP

批改「薪資樞紐分析 / 成績表 / 車輛銷售直條圖 / 武將能力雷達圖」四大題、共18個
小項的 Excel 實作考題，並把結果寫進標準格式的『單科評分結果』活頁簿。

## 這個 skill 涵蓋的固定資產（assets/）

- `01_template/S06_excel_scoring_criteria_v1.xlsx` — 評分標準表（大項/小項/評分標準/配分/
  額外加分配分）。**這是評分的唯一依據**，如果使用者這次提供了不同版本的評閱
  標準檔案，以使用者提供的為準，並比對是否有配分變動。
- `01_template/S06_excel_question_v1.xlsx` — 原始題目檔（未作答），用來比對「原始資料有沒有
  被更動」。
- `01_template/S06_excel_answer_v1.xlsx` — 標準解答檔，可作為「正確答案長什麼樣子」的參考，
  但**評分永遠以評閱標準表的文字敘述為準**，不是拿學生檔案跟答案檔案做逐格
  diff（答案檔案本身的呈現方式不一定是唯一正確解，例如 1.2 欄標籤兩種呈現
  方式都合格）。
- `03_results/temp_S06_單科_評分結果yyyymm_v1.xlsx` — 空白輸出模板，兩個活頁的表頭固定
  格式如下：
  - 活頁1「評分細項」：截止日/課程代碼/實作項目/部門代號/工號/姓名/
    大項/小項代號/得分/完成比例/執行者/評閱結果說明
  - 活頁2「評分結果表」：截止日/課程代碼/實作項目/部門代號/工號/姓名/副檔名/檔案數/
    分數加總/是否合格/執行者/評閱結果簡述(150字以內)

詳細的逐小項技術檢查方法、常見錯誤案例、已知檔案的評分結果對照表，都在
`references/scoring_checks.md`，開始評分前先讀過一遍。

## 梯次資料夾與工具（scripts/、engine/）

每一梯次的資料放在 `<資料根目錄>/grading/training_internal/S06_<yyyymm>/`，資料根目錄
預設 `D:\claude_data`（可用 `scripts/local_config.json` 的 `data_root` 或 `--root` 改）：

```
S06_<yyyymm>/
├─ 00_input/            manifest_source/（S06_三科_原始檔案<yyyymm>[_yyyymmdd_HHMMSS].csv）、
│                       roster/（S06_三科_學員名冊<yyyymm>.xlsx）、submissions/（學員作業）
├─ 01_template/         評閱標準、題目、答案（init_batch.py 從 assets 複製）
├─ 02_work/             Excel／Word／PPT 評閱工作區，評分前必須淨空
└─ 03_results/          單科評分結果、學員檔案與成績（temp_ 開頭的是樣板，不要刪）
```

工具都在本 skill 資料夾內，只需要 Python 3.11+ 和 openpyxl：

| 工具 | 用途 |
|---|---|
| `python scripts/init_batch.py --batch <yyyymm> [--dry-run]` | 建立梯次資料夾並複製樣板；已存在的檔案一律不覆寫 |
| `python scripts/check_batch.py --batch <yyyymm>` | 評分前檢查六項：清單、名冊有資料、submissions 有資料夾、01_template 三份檔案（顯示版本與修改日）、02_work 淨空、03_results 只有樣板。唯讀 |
| `python scripts/copy_to_work.py --batch <yyyymm> [--type all\|xlsx\|docx\|pptx] [--dry-run]` | 預設三種檔案都處理（`--type` 可只選一種），把合規的檔案複製到 02_work/Excel｜Word｜PPT：路徑是 `submissions/<單位-工號-姓名>/<單位-工號-姓名>.<副檔名>`，資料夾與檔名的單位-工號-姓名一致，且是 00_input/roster 學員名冊中的一列（單位不分大小寫）；其餘列出略過原因，02_work 已有的不合規檔案只提醒、不刪除。不解壓縮 zip／7z／rar，不覆寫 02_work 已有的檔案 |
| `python scripts/build_all_report.py --batch <yyyymm>` | 三科單科評分都完成後執行：複製 03_results 的 `temp_<課程>_三科_學員檔案與成績計算yyyymm_v1.xlsx` 成 `<課程>_三科_學員檔案與成績計算<梯次>_<yyyymmdd_HHMMSS>.xlsx`；活頁 manifest 依原始檔案清單填入每個檔案與歸檔／命名／重複異常（歸檔異常與 copy_to_work.py 同規則：資料夾的單位-工號-姓名須是名冊一列、檔名須與資料夾一致）；活頁 score 依學員名冊順序填入學員，並把各科 `<課程>_<Excel\|Word\|PPT>_評分結果<梯次>_<yyyymmdd_HHMMSS>.xlsx`「評分結果表」的分數加總填入 Excel原始／Word原始／PPT原始（缺檔的科目留空白；單科檔沒評到分的學員以 0 分計，視為歸檔異常）；歸檔異常、命名異常（含未繳交）的科目，原始與 Excel／Word／PPT 欄都填 0，其餘 Excel／Word／PPT 為原始分數（上限 100）；備註依科目寫出異常與封頂原因。總分、是否合格不填。不需要 MCP |
| `python engine/extract_submission.py <梯次>/02_work/Excel/<學生檔案> --question <題目檔>` | 輸出單一作答檔的客觀事實（JSON），學生檔案一律取自 02_work，見第2步 |

梯次的 01_template 有多個版本（_v1、_v2）時，一律取版本號最大的。

## 若環境有 kd-grading-internal 的 `gi_*` tools

在 my-ai-skills repo 用 Claude Code 時，`.mcp.json` 會載入 kd-grading-internal
（內訓評分 server）。此時第1~3步可以改用下列 tools，呼叫時一律帶
`skill="04-excel-grading"`：

- `gi_get_rubric`：核對配分。
- `gi_list_submissions`：解析檔名；`folder` 指向梯次的 `02_work/Excel`（例如 `S06_202609/02_work/Excel`）。
- `gi_grade_submission`：依 `references/scoring_checks.md` 計分，並回傳 `needs_review` 項目。
- `gi_write_results`：寫入單科評分結果活頁簿；帶 `batch`（例如 `202612`）時直接存到梯次的 `03_results/`，
  找不到梯次資料夾才存到 server 的 output 資料夾。

流程規則不變：
- 第0步仍要先問齊四個欄位。課程代碼和實作項目雖然會預設帶入本資料夾 `skill.toml`
  的值，仍要向使用者確認。
- `needs_review` 的項目要人工確認後，用 `overrides` 帶入。
- 評閱結果簡述仍由你依得分分布撰寫，再傳入 `summary`。

沒有這些 tools 時，照下面的流程手動執行。

## 流程

### 第0步：確認缺漏的「無法從檔名取得」的資訊

檔名格式「部門代號-工號-姓名.xlsx」只能解析出部門代號/工號/姓名三個欄位。
以下欄位**無法從任何檔案自動取得**，必須向使用者確認（如果對話中已經提供過，
直接沿用，不用重複問）：

- 截止日
- 課程代碼／課程名稱（例如 "S06"）
- 實作項目（例如 "Excel實作"）
- 執行者（評閱者姓名或代號）

如果這是一次全新的評分任務且以上任一項缺漏，先用一次提問把它們問齊，再開始
批改。如果是「幫既有結果表追加學生」的情境（活頁簿已存在），直接讀取既有
活頁2任一列的這四個欄位沿用即可，不用重問。

同時檢查每個上傳檔案的檔名是否符合「部門代號-工號-姓名.xlsx」格式：
- 檔名用 `-` 切開應該剛好是三段（副檔名前）。
- 如果檔名不符合格式（段數不對、看起來像是打錯字），停下來跟使用者確認
  正確的部門代號/工號/姓名，不要自己猜測或用檔名硬湊。

### 第1步：讀取評閱標準

用 `openpyxl`/`markitdown` 讀梯次的 `01_template/S06_excel_scoring_criteria_v<最大版本>.xlsx`
（沒有梯次資料夾時讀本 skill 的 `assets/01_template/`；或使用者本次提供的版本），取出大項、小項代號（如 "1.1"）、配分、額外加分
配分。18 個小項與滿分彙總已經整理在 `references/scoring_checks.md` 最上方
的表格，可以直接核對、不用每次重新解析。

### 第2步：逐一批改每個學生檔案

學生檔案**一律讀梯次的 `02_work/Excel/`**（由 `copy_to_work.py` 從 submissions 複製過來的合規檔案），
不要直接讀 `00_input/submissions/`；不在 02_work 的檔案不納入評分。對 `02_work/Excel/` 下的每個 .xlsx：

1. 用 `python engine/extract_submission.py <梯次>/02_work/Excel/<學生檔案> --question <梯次>/01_template/S06_excel_question_v<最大版本>.xlsx`
   跑一次（只需要 openpyxl），拿到結構化 JSON（pivot table 設定、圖表動態性與座標軸/圖例內容、
   凍結窗格、合併/跨欄置中、框線、排序/篩選、公式）。
2. 對照 `references/scoring_checks.md` 的規則，逐一判斷 18 個小項的得分，
   並用一兩句話寫出「評閱結果說明」（未完成的欠項，或已完成的加分項——
   對照原本的評分標準文字，講清楚具體原因，不要只寫「正確」兩個字）。
3. 特別注意這些「連坐」規則，先檢查、可以省掉後面逐項檢查的力氣：
   - 整個檔案沒有以【薪資】資料為來源的樞紐表（樞紐表可放在其他活頁，但要
     排除【考題】活頁本身的範例）→ 1.1~1.5 全部0分。
   - 車輛銷售／武將能力的 A欄或第2列原始資料被更動 → 對應大題(3.x 或 4.x)
     全部0分，不論圖表其他設定是否正確。
   - 圖表不是動態的，或類型不對（考題3只限長條圖／直條圖／直方圖，考題4只限
     雷達圖）→ 對應大題全部0分。
4. 完成比例 = 得分 / 該小項配分（含額外加分小項則除以額外加分上限），存成
   0%~100% 的百分比格式。

### 第3步：彙整活頁2並產出檔案

- 「分數加總」用公式（例如
  `=SUMIFS(評分細項!$I$2:$I$<末列>,評分細項!$E$2:$E$<末列>,E<列>,評分細項!$F$2:$F$<末列>,F<列>)`）
  依工號＋姓名加總該學生在活頁1所有列的得分，不要手動加完寫死數字——這樣之後如果活頁1
  有調整還能自動重算。
- 「是否合格」用公式判斷 `>=60` 顯示「合格」否則「不合格」。
- 「評閱結果簡述」150字以內：約100字寫優點、約50字寫待改善（如果沒有待改善
  項目可以省略後半段）。內容要具體對應到該學生實際的得分分布，不要寫套話。
- 輸出檔名慣例：`<課程代碼>_Excel_評分結果<梯次yyyymm>_<yyyymmdd_HHMMSS>.xlsx`
  （例如 `S06_Excel_評分結果202609_20261021_160327.xlsx`，或使用者指定的檔名）。

### 第4步：如果是「追加學生」到既有結果檔

如果使用者提供的是已經有資料的評分結果檔（而不是空白模板），用 openpyxl
讀取兩個活頁目前的 `max_row`，從下一列開始 append，**不要**覆蓋既有列。
append 完後，記得把活頁2所有列（含原本就有的）的 SUMIFS 公式範圍
（`$2:$<末列>` 那幾段）更新成涵蓋新增後的活頁1末列，否則舊學生的加總會
漏掉新資料（活頁1本來就沒新增列的話則不受影響，但統一重寫一次比較保險）。

### 第5步：驗證與交付

- 用 xlsx skill 的 `recalc.py` 重新計算公式，確認 `status: success` 且
  `total_errors: 0`。
- 用 openpyxl 以 `data_only=True` 讀回結果，跟自己算的分數/合格判定核對一次
  再交付，避免公式範圍寫錯導致加總對不上。
- 有梯次資料夾時存到梯次的 `03_results/`（`build_all_report.py` 從這裡讀）；在 claude.ai
  上則存到 `/mnt/user-data/outputs/`，用 `present_files` 交付。並在對話中用
  條列式簡短說明每位學生的總分、合格與否、以及扣分主因（不用整份 12欄 x
  18列的明細貼在對話裡，那些都在檔案裡了）。

## 何時要人工介入、不要自己硬判

- 3.4「運算列表資料標籤」是否為圖表下方的「資料表(Data Table)」功能，
  `engine/extract_submission.py` 會回報 `has_data_table`；遇到 `has_data_labels: false`
  但 `has_data_table: true` 時，打開 chart XML（`xl/charts/chartN.xml`）人工確認一次。
- 如果評閱標準表本身跟這次使用者提供的版本配分不同（例如日後調整），以
  使用者這次提供的檔案為準，並在回覆中提醒使用者配分已更新。
- 檔名解析不出部門代號/工號/姓名，或一個檔案裡找不到預期的四個工作表
  （薪資/成績/車輛銷售/武將能力），先跟使用者確認檔案是否正確，不要憑空
  帶入預設值。工作表名稱允許前後綴（例如「1.薪資」「成績_作答」），名稱完全
  相同的優先，否則取唯一一個含該文字的工作表；實際對應見
  `extract_submission.py` 輸出的 `sheet_names`。有多個工作表都含同一文字時
  （`ambiguous_sheets`）不會自動挑選，要跟使用者確認評哪一個。
