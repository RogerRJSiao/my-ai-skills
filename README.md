# 🤖 My AI Skills Hub

> 集中管理所有 AI 自動化技能小專案的 GitHub 倉庫。
> 每個 Skill 專案皆為獨立資料夾，可單獨複製指令交給 AI 一鍵執行。

---

## 📂 Skill 清單

| # | 資料夾 | 說明 | 觸發關鍵字（範例） | 狀態 |
|---|--------|------|------------------|------|
| 01 | [01-credit-card-rewards-xlsx](./skills/01-credit-card-rewards-xlsx/) | 信用卡優惠 PDF → Excel 整理工具，萃取商家回饋資訊並輸出統一格式 `.xlsx` | 信用卡優惠、回饋整理 | ✅ 可用 |
| 02 | [02-pbi-version-management](./skills/02-pbi-version-management/) | Power BI (`.pbix`) 版號判斷、檔名規則與 CHANGELOG 維護，含常見疏漏檢查清單 | pbix 版控、更新 changelog | ✅ 可用 |
| 03 | [03-improvement-proposal-draft](./skills/03-improvement-proposal-draft/) | 「改善提案表」撰寫與檢查，將技術素材轉為非技術主管看得懂的四段式提案 | 改善提案、現狀缺失、改善內容 | ✅ 可用 |
| 04 | [04-excel-grading](./skills/04-excel-grading/) | 「S06 Excel實作」評分 SOP，批改學員作答檔（4 大題 18 小項）並寫入評分結果活頁簿 | S06、Excel實作、評分/批改 | ✅ 可用 |

---

## 🚀 快速開始

1. 進入任一 Skill 資料夾，閱讀該資料夾的 `README.md`
2. 依需求選擇使用方式：
   - **Claude.ai**：將資料夾打包成 `.skill` 上傳（Settings → Features → Skills）
   - **Claude Code**：將資料夾複製到 `~/.claude/skills/` 或專案的 `.claude/skills/`
   - **其他 AI**：把 `SKILL.md` 全文貼給 AI，並附上指令「請依照此 Skill 執行」
3. 若該 Skill 含 `requirements.txt`（01、04），先執行 `pip install -r requirements.txt`
4. 提供素材（PDF、pbix 截圖、技術資料、學生 `.xlsx` 等），AI 即依 Skill 流程產出結果

---

## 📂 各類檔案在 AI 提示詞中的角色對照

| 檔案 / 資料夾 | 在 AI 提示詞（Prompt）中扮演的角色 | 範例 |
|--------------|----------------------------------|------|
| `SKILL.md` | **System Prompt（系統設定）**。定義 AI 的專家身份、觸發條件、核心工作流（Pipeline）與思考邏輯。 | 所有 Skill |
| `references/*.md` | **Knowledge Base / Data Constraints（知識庫與資料約束）**。欄位規則、同義詞對照、檢查細則等，限制 AI 不可瞎編。 | `column_rules.md`、`merchant_normalize.md`、`changelog_template.md`、`scoring_checks.md` |
| `assets/` | **Templates / Fixed Inputs（輸出範本與固定資產）**。輸出格式範本、評分標準、標準解答等。 | `template_spec.md`、`評閱標準.xlsx`、`評分結果.xlsx` |
| `scripts/` | **Tools（可執行工具）**。讓 AI 呼叫的腳本，負責產出檔案或萃取客觀事實。 | `build_xlsx.py`、`inspect_submission.py` |
| `examples/` | **Few-Shot（範例）**。真實案例，作為 AI 歸納寫法與規則的依據。 | 03 的歷年改善提案 |
| `docs/` | **給人讀的操作手冊**，不直接餵給 AI。 | 03 的 `改善提案撰寫操作建議.md` |

---

## 📋 專案規範

- 每個 Skill 必須包含：`README.md`、`SKILL.md`
- 選配：`references/`、`scripts/`、`assets/`、`sample-data/`、`examples/`、`docs/`
- 輸出的 `.xlsx` 不納入版控（`outputs/` 已加入 `.gitignore`）
- `assets/*.xlsx` 為輸出範本，**納入版控**（`.gitignore` 已豁免）
- `sample-data/*.pdf` 不納入版控（避免上傳銀行原始資料）
- 有 Python 腳本的 Skill，依賴統一寫在該 Skill 的 `requirements.txt`
- 資料夾命名：`{兩位數編號}-{skill-name}`，`SKILL.md` 一律放在 Skill 資料夾根目錄

---

## 🗂️ 版控規則說明

```
納入版控   ✅  skills/**/assets/*.xlsx        （輸出範本、評分標準等固定資產）
納入版控   ✅  所有 .md、.py、.txt、.png
納入版控   ✅  skills/**/sample-data/bank_*.pdf   （公開優惠頁，無個資）
納入版控   ✅  skills/**/references/*.pdf         （AI 對話紀錄等參考文件）
排除版控   ❌  outputs/*.xlsx                （本地產出）
排除版控   ❌  sample-data/*.pdf             （原始資料，bank_ 開頭除外）
```
