# 💳 01 — 信用卡優惠 PDF → Excel 整理工具

本 Skill 能讀取一或多份銀行信用卡優惠頁面截圖或 PDF，
自動萃取所有商家回饋資訊，輸出為格式統一、可篩選的 `.xlsx` 報表。

> **💡 通用概念說明**
> 本 Skill 的核心概念不限於信用卡回饋，同樣適用於一般企業導入 AI Agent 的場景：
> 將非結構化的 PDF 文件（如報表、型錄、合約摘要、政策文件）交給 AI 自動讀取、
> 萃取關鍵欄位，並整理成標準化的 Excel 表格，取代人工逐頁複製貼上的作業流程。
> 只需定義好欄位規則（如本專案的 `column_rules.md`）與輸出格式範本，
> 即可快速複用此模式，應用於不同部門或不同類型的文件整理需求。

---

## 📁 資料夾結構

```
01-credit-card-rewards-xlsx/
├── README.md                          ← 本文件（給人類閱讀）
├── SKILL.md                           ← 核心：給 AI 讀取的自動化 Pipeline 技能檔
├── requirements.txt                   ← Python 環境依賴套件
├── references/
│   ├── column_rules.md                ← 10 欄位定義與填寫規則
│   ├── merchant_normalize.md          ← 商家名稱標準化 & 同義對照表
│   ├── organize_skills_for_credit_cards_reward_by_Claude_Code_AI.pdf  ← AI 對話紀錄
│   └── organize_skills_for_credit_cards_reward_by_Google_Gemini_AI.pdf ← AI 對話紀錄
├── scripts/
│   └── build_xlsx.py                  ← 可直接執行的輸出腳本
├── assets/
│   ├── template_spec.md               ← 色碼/字型/欄寬完整規格說明
│   └── 信用卡回饋方案整理_sample.xlsx  ← ✅ 實際輸出範例（550筆，4張卡）
└── sample-data/
    ├── README.md                      ← PDF 放置說明
    ├── bank_台新Richart_rewards.pdf   ← ✅ 納入版控（公開優惠頁）
    ├── bank_聯邦Line點_rewards.pdf
    ├── bank_國泰CUBE_rewards.pdf
    └── bank_中信AllMe_rewards.pdf
```

---

## ▶️ 使用模式

**實作範例：**
- 輸出結果：[信用卡回饋方案整理_sample.xlsx](./assets/信用卡回饋方案整理_sample.xlsx)
- 來源 PDF（共 4 份）：
  [bank_台新Richart_rewards.pdf](./sample-data/bank_台新Richart_rewards.pdf) ／
  [bank_聯邦Line點_rewards.pdf](./sample-data/bank_聯邦Line點_rewards.pdf) ／
  [bank_國泰CUBE_rewards.pdf](./sample-data/bank_國泰CUBE_rewards.pdf) ／
  [bank_中信AllMe_rewards.pdf](./sample-data/bank_中信AllMe_rewards.pdf)

---

### 模式一：手動貼貼 — 適合臨時處理 1-2 份 PDF

如果你只是手動下載了 1-2 家銀行的新 PDF，想立刻讓 AI 轉成表格，可以直接在任何 AI 聊天介面操作，不需安裝任何工具。

**操作步驟：**

1. 打開你常用的 AI 聊天介面（Claude、ChatGPT 等皆可）
2. 將以下三個檔案的文字內容全部複製貼給 AI（或直接拖入讓它讀取）：
   - [`SKILL.md`](./SKILL.md)（告訴 AI 它是誰、要做什麼任務）
   - [`references/column_rules.md`](./references/column_rules.md)（告訴 AI 欄位怎麼填）
   - [`references/merchant_normalize.md`](./references/merchant_normalize.md)（告訴 AI 商店名稱怎麼對齊）
3. 同時上傳你想處理的銀行回饋 PDF（或複製網頁上的優惠文字貼給它）
4. 最後加上這段 Prompt：

```
請擔任 SKILL.md 中定義的角色，嚴格遵循 column_rules.md 與 merchant_normalize.md
的規範，分析我提供給你的這份銀行 PDF 資料，並直接為我輸出一張 Markdown 格式的
資料表（包含那 10 個欄位）。
```

**產出結果：** AI 會直接在對話框輸出 Markdown 表格，全選複製後貼到 Excel 即完成。

---

### 模式二：自動化 Pipeline — 適合批次處理多份 PDF

> **前置需求：** 需安裝 Python，並執行 `pip install -r requirements.txt`

1. 將銀行優惠 PDF 放入 `sample-data/`
2. 將 `SKILL.md` 全文貼給 AI，附上以下指令：

```
請依照此 SKILL.md 執行，來源 PDF 在 sample-data/ 資料夾，
輸出格式請參考 assets/信用卡回饋方案整理_sample.xlsx，
輸出檔案命名為：信用卡回饋方案整理_YYYY-MM-DD.xlsx
```

3. AI 將自動完成所有步驟並提供下載連結

---

## 📦 環境需求

```bash
pip install -r requirements.txt
```

---

## 📤 輸出欄位（共 10 欄）

| 欄 | 欄位名稱 | 說明 |
|----|----------|------|
| A | 信用卡名稱 | 發卡銀行與卡片名稱 |
| B | 開始日 | 優惠開始日 (yyyy-mm-dd) |
| C | 結束日 | 優惠結束日 (yyyy-mm-dd) |
| D | 商家名稱 | 合作商家或通路描述 |
| E | 回饋%數 | 數字加%，如 `3%` |
| F | 方案切換 | 需切換方案填方案名；需綁定支付工具填工具名（如 `LINE Pay`）；一般持卡即享填「通用」 |
| G | 相同商家 | 跨卡重複商家標記 `Exist`（淡黃底） |
| H | 最高回饋 | 同商家中最高回饋率標記 `Max`（淡綠底） |
| I | 優惠備註 | 限制條件、方案說明等 |
| J | 查詢日期 | 資料查詢當天日期 (yyyy-mm-dd) |

---

## 🗂️ 版控說明

| 路徑 | 是否納入版控 |
|------|------------|
| `assets/信用卡回饋方案整理_sample.xlsx` | ✅ 是（輸出範本） |
| `sample-data/bank_*.pdf` | ✅ 是（公開優惠頁，無個資） |
| `references/*.pdf` | ✅ 是（AI 對話紀錄等參考文件） |
| `sample-data/*.pdf`（非 bank_ 開頭） | ❌ 否（原始資料） |
| `outputs/*.xlsx` | ❌ 否（本地產出） |
