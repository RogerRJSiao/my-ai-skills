# 🤖 My AI Skills Hub

> 集中管理所有 AI 自動化技能小專案的 GitHub 倉庫。
> 每個 Skill 專案皆為獨立資料夾，可單獨複製指令交給 AI 一鍵執行。

---

## 📂 Skill 清單

| # | 資料夾 | 說明 | 狀態 |
|---|--------|------|------|
| 01 | [01-credit-card-rewards-xlsx](./skills/01-credit-card-rewards-xlsx/) | 信用卡優惠 PDF → Excel 整理工具 | ✅ 可用 |
| 02 | [02-future-skill-template](./skills/02-future-skill-template/) | 未來擴充用模板 | 🚧 佔位 |

---

## 🚀 快速開始

1. 進入任一 Skill 資料夾，閱讀該資料夾的 `README.md`
2. 將 PDF 原始資料放入 `sample-data/`
3. 把 `SKILL.md` 全文貼給 AI，並附上指令「請依照此 Skill 執行」
4. AI 將自動讀取 PDF、整理資料、輸出 `.xlsx`

---

## 📋 專案規範

- 每個 Skill 必須包含：`README.md`、`SKILL.md`
- 選配：`references/`、`scripts/`、`assets/`、`sample-data/`
- 輸出的 `.xlsx` 不納入版控（`outputs/` 已加入 `.gitignore`）
- `assets/*.xlsx` 為輸出範本，**納入版控**（`.gitignore` 已豁免）
- `sample-data/*.pdf` 不納入版控（避免上傳銀行原始資料）
- Python 依賴統一寫在各 Skill 的 `requirements.txt`

---

## 🗂️ 版控規則說明

```
納入版控   ✅  skills/**/assets/*.xlsx   （輸出範本）
納入版控   ✅  所有 .md、.py、.txt
排除版控   ❌  outputs/*.xlsx            （本地產出）
排除版控   ❌  sample-data/*.pdf         （原始資料）
```

---

## 📄 授權

MIT License — 詳見 [LICENSE](./LICENSE)
