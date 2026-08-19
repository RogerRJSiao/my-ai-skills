# improvement-proposal-draft

「改善提案表」撰寫輔助 skill，適用於多種主題的改善提案（數位轉型、系統優化、教育訓練教材更新、報表效能改善等），不限單一部門或技術領域。

## 這是什麼

一份可安裝進 Claude 的 skill，內建「改善提案表」的四段式架構、現狀缺失三層公式、改善內容的「專業詞彙+白話科普」寫法、台灣用詞檢查清單、送出前排版檢查清單。規則歸納自歷年實際提案案例（詳見 `examples/`）。

## 如何使用

1. 從 `SKILL.md` 打包成 `.skill` 檔，或將整個 repo clone 下來，把此資料夾當作 skill 目錄使用。
2. 在 Claude 中安裝（Save skill）。
3. 之後只要提到「改善提案」「現狀缺失」「改善內容」等相關需求，或上傳程式碼/截圖/PDF等技術素材並希望轉換成提案格式，會自動觸發此 skill。

## 目錄結構

```
improvement-proposal-draft/
├── README.md                          # 本說明文件
├── SKILL.md                           # skill核心邏輯（給Claude讀）
│
├── docs/
│   └── 改善提案撰寫操作建議.md          # 給人讀的操作手冊（教育訓練/紙本參考用）
│
└── examples/
    ├── proposals/                     # 歷年提案範例（歸納skill規則的依據）
    │   ├── 01_數位轉型自動化.txt
    │   ├── 02_成品系統統一編號改善.txt
    │   ├── 03_S12教育訓練教材更新.txt
    │   ├── 04_PBI開箱日報名整合.txt
    │   ├── 05_S07六大核心工具更新.txt
    │   └── 06_PBI財務會計模板檔與版本管理改善.txt   # 最終定稿範例
    │
    ├── materials/                     # 改善前的佐證素材(依提案主題分子資料夾)
    │   └── pbi-financial-template/
    │       ├── v3.txt
    │       ├── v3.1_full.txt
    │       ├── v3.1_finance.txt
    │       ├── CHANGELOG_Finance_Accounting.md
    │       └── timeout-error-screenshot.png
    │
    ├── deliverables/                  # 改善後的實作成果(同樣依提案主題分子資料夾)
    │   └── pbi-financial-template/
    │       └── README.md              # 目前僅有SharePoint連結，尚無實體檔案
    │
    └── official-form/                 # 正式系統的申請表單截圖(通用參考，非特定提案)
        ├── 申請表單畫面.png            # 完整欄位配置(提案名稱/現狀缺失/改善內容/效益分析等)
        └── 分類下拉選單.png            # 完整9種分類選項
```

## 設計原則

- **`SKILL.md` 放根目錄**：Claude 的 skill 系統以資料夾內的 `SKILL.md` 作為進入點，路徑越短，日後 clone 後越好直接掛載使用。
- **`examples/proposals/`**：保留完整的歷史提案文字，作為 skill 規則的可追溯依據，方便日後調整規則時回頭對照真實案例。
- **`examples/materials/`**：依「提案主題」分子資料夾（而非依檔案類型分類），存放**改善前**用來佐證問題點的素材（程式碼、SQL、截圖、CHANGELOG等）。同一提案的素材可能同時包含多種格式，用主題分類比用類型分類更容易維護與查找。
- **`examples/deliverables/`**：與`materials/`對應但方向相反，存放**改善後**的實作成果（最終產出檔案或連結）。目前多數提案的成果只存在公司SharePoint，尚未有實體檔案入repo時，先用`README.md`記錄連結即可；日後若要放實體大型檔案（如pbix），建議評估是否需要git LFS，避免直接塞進一般git repo造成repo肥大。
- **`examples/official-form/`**：正式系統的表單截圖，屬於通用參考（不綁定特定提案主題），用來確認欄位配置與分類下拉選單的最新選項，若公司表單日後改版，只需更新這裡的截圖即可。
- **`docs/` 與 `SKILL.md` 分離**：前者給人讀（教育訓練、紙本參考），後者給 Claude 讀（機器指令），維護頻率與對象不同，分開可避免互相干擾。

## 使用限制

- 不會自動判斷提案的「分類標籤」絕對正確，僅提供建議，最終仍需依公司審核制度確認。
- 不會驗證 SharePoint、GitLab 等外部連結是否仍然有效。
- 不會捏造未提供素材中的技術細節或數據，資訊不足時會主動詢問。
