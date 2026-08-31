# 04 — excel-grading

「S06 Excel實作」考題評分 SOP skill，批改「薪資樞紐分析 / 成績表 / 車輛銷售
直條圖 / 武將能力雷達圖」四大題、共18個小項的學生 .xlsx 作答檔，並把結果寫進
標準格式的『評分結果』活頁簿。

來源：從 `s06-excel-grading_v0.1.skill`（Claude skill 打包檔）解壓匯入。

## 這是什麼

一份可安裝進 Claude 的 skill，內建評閱標準比對邏輯、樞紐表/圖表/凍結窗格等
技術檢查方法、常見錯誤案例，以及「評分結果」活頁簿（明細 + 結果表）的固定
輸出格式。

## 如何使用

1. 從 `SKILL.md` 打包成 `.skill` 檔，或將整個 repo clone 下來，把此資料夾當作
   skill 目錄使用。
2. 在 Claude 中安裝（Save skill）。
3. 上傳一份或多份「部門代號-工號-姓名.xlsx」學生作答檔，並提到評分/批改
   「S06」或「Excel實作」，會自動觸發此 skill。

## 目錄結構

```
04-excel-grading/
├── README.md                                      # 本說明文件
├── SKILL.md                                       # skill核心邏輯（給Claude讀）
├── requirements.txt                               # Python 依賴（openpyxl）
│
├── assets/                                        # 評分固定資產（評分依據＋輸出模板）
│   ├── S06_Excel實作檔_評閱標準_v0_1.xlsx           # 評分標準表（唯一評分依據）
│   ├── S06_excel_question.xlsx                    # 原始題目檔（未作答，用來比對是否被更動）
│   ├── S06_excel_answer.xlsx                      # 標準解答檔（參考用，非逐格diff依據）
│   └── S06_Excel實作檔_評分結果_v0_1.xlsx           # 空白輸出模板（活頁1明細＋活頁2結果表）
│
├── references/
│   └── scoring_checks.md                          # 逐小項技術檢查方法、常見錯誤案例
│
└── scripts/
    └── inspect_submission.py                      # 讀取學生檔案，輸出結構化 JSON 供比對評分
```

## 設計原則

- **評分依據以「評閱標準」文字敘述為準**：不是拿學生檔案跟標準解答檔逐格
  diff，因為同一評分標準底下可能有多種合格的呈現方式。
- **`inspect_submission.py` 只萃取客觀事實，不計分**：樞紐表設定、圖表動態性、
  凍結窗格、合併儲存格、框線、排序/篩選、公式等，由模型依 `scoring_checks.md`
  的規則自行判斷得分，避免每次評分重新推導 XML 解析邏輯。
- **`references/scoring_checks.md` 與 `SKILL.md` 分離**：前者是逐小項的技術
  檢查細則與案例庫（內容會隨遇到的新案例增修），後者是評分流程與規則的
  進入點，分開維護避免流程文件被案例細節塞爆。

## 使用限制

- 不會自動判斷評閱標準的「配分」是否為最新版本，若使用者提供不同版本的
  評閱標準檔案，以使用者提供的為準。
- 不會捏造未提供素材中的技術細節或分數，資訊不足（檔名解析不出來、缺少
  必要工作表等）時會主動詢問，不會憑空帶入預設值。
- `inspect_submission.py` 目前偵測不到圖表的「資料表(Data Table)」節點
  （`plotArea/dTable`），遇到疑似資料標籤但腳本回報 `has_data_labels: false`
  時，需要人工開 chart XML 確認。
