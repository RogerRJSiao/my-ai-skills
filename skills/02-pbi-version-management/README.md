# pbi-version-management

一個 Claude Agent Skill，用來記錄與維護 Power BI (`.pbix`) 檔案的版本命名規則與 CHANGELOG。

適用情境：多個部門/欄位模式（如 full / finance）並行維護同一份底層邏輯的 pbix 專案，需要一致的版號判斷標準、命名規則，以及每次更新時的紀錄檢查。

---

## 這個 Skill 能幫你做什麼

- 判斷本次改動該跳**主版號**還是**次版號**（依資料模型/參數機制 vs SQL/欄位/bug 修正區分）
- 套用統一的命名規則：`{報表名稱}v{主版號}.{次版號}_{欄位模式}.pbix`
- 讀取 pbix 檔案本身 + 儲存位置截圖（如 SharePoint、知識中心檔案清單），交叉核對檔名、上線日期是否一致
- 產出/更新 CHANGELOG.md，並自動跑過一套「常見疏漏檢查清單」（欄位定義前後矛盾、備註漏寫改動項目、欄位數暴增暴減沒說明原因、欄/列用字混淆等）

---

## 部署建議

### 方式一：Claude.ai 網頁版 / App（個人帳號，Pro / Max / Team / Enterprise）

1. 到本 repo 下載 `pbi-version-management.skill`（或直接下載整個資料夾，內含 `SKILL.md` 與 `references/`）
2. 登入 [claude.ai](https://claude.ai)，進入 **Settings → Features → Skills**
3. 確認已開啟 **Code execution**（上傳自訂 skill 的前提條件）
4. 點擊 **Upload skill**，選擇下載好的 `.skill` 檔上傳
5. 上傳後即可在對話中自動觸發，不需要每次手動喚醒

> 注意：個人帳號上傳的 skill 預設**只有自己看得到**。若同事也要使用，必須各自下載本 repo 的檔案並自行上傳。

### 方式二：Team / Enterprise 方案 —— 直接分享給同事

- 若你的帳號在 Team 或 Enterprise 方案下，上傳後可以直接將 skill **分享給指定同事**，對方會在自己技能清單的「已分享的技能」看到並可啟用（但不可編輯內容，你更新後對方會自動同步最新版）
- 若你是組織擁有者，也可以把 skill **provision（佈建）給整個組織**，全體成員自動取得，不需個別上傳

### 方式三：Claude Code（本機開發環境）

Claude Code 的 skill 遵循開放的 Agent Skills 標準，不需透過網頁上傳：

```bash
# 個人使用（僅自己的機器）
git clone <本repo網址>
cp -r pbi-version-management ~/.claude/skills/

# 專案共用（隨 repo 一起版控，team 成員 clone 專案即自動取得）
cp -r pbi-version-management <你的專案路徑>/.claude/skills/
```

放好後 Claude Code 會自動偵測並在相關任務時載入，不需額外設定。

### 方式四：Claude API / 自建應用整合

若你們有自建的內部工具串接 Claude API，可透過 Skills API 上傳這份 skill，讓組織內所有呼叫 API 的服務共用（不受個人帳號限制）。細節請參考 [Claude Platform Docs - Agent Skills](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview)。

---

## 更新這個 Skill

若之後發現新的檢查項目（例如又踩到一種容易漏寫的備註情境），建議：

1. 直接修改本 repo 的 `SKILL.md`，在「撰寫備註時的檢查清單」段落新增一條
2. commit + push，讓改動有版本紀錄可追溯（這跟本 skill 教 pbix 做的事情是同一套精神）
3. 若已上傳到 Claude.ai，記得重新打包上傳覆蓋舊版；若是 Claude Code 的 `.claude/skills/` 方式，直接 `git pull` 更新即可生效

---

## 提示詞範例

以下範例展示這個 skill 預期會被觸發的情境，可以直接照樣造句：

### 1. 新版本要上線，請 Claude 判斷版號

> 「這次改動是把 SQL 查詢加上索引鍵，並把財務用的欄位從 47 欄縮減到 27 欄，其他都沒變，請問這次要跳 v3.2 還是 v4？」

### 2. 上傳截圖 + pbix 檔案，請 Claude 核對版本紀錄

> 「附件是我們知識中心資料夾的截圖，還有這次要上線的 pbix 檔案，請幫我核對檔名、修改時間跟我們既有的 CHANGELOG.md 是否一致，有沒有需要補充或修正的地方。」

### 3. 直接請 Claude 更新 CHANGELOG

> 「幫我把這次的改動加進 CHANGELOG.md：full 版本這次優化了兩個子查詢的索引使用方式，finance 版本沒有變動但版號要跟著同步升級，上線日是這週五。」

### 4. 詢問版控判斷邏輯本身

> 「我們新增了一個『地區別』的篩選參數，這個算大改還是小改？如果之後還要再加『幣別』參數呢？」

### 5. 新專案要套用同一套規則

> 「我要幫另一個部門的 pbix 也建立一份 CHANGELOG，欄位模式分別是『業務』跟『財務』兩種，麻煩套用我們既有的命名規則幫我產出模板。」

### 6. 檢查既有 CHANGELOG 是否有疏漏

> 「幫我重新檢查一次這份 CHANGELOG.md，看看有沒有欄位定義前後不一致、或是備註漏寫的地方。」（並上傳既有的 CHANGELOG.md 檔案）

---

## 檔案結構

```
pbi-version-management/
├── SKILL.md                          # 主要規則：命名規範、版號判斷標準、檢查清單、工作流程
└── references/
    └── changelog_template.md         # CHANGELOG.md 空白模板，可直接複製套用到新專案
```

---

## 授權 / 使用範圍

本 skill 為內部維護規範，依專案需求自由修改調整（例如報表名稱前綴、欄位模式定義），不綁定特定資料庫或 BI 工具版本。
