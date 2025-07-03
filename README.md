# 智慧排班系統 (Intelligent Shift Scheduling System)

這是一個結合大型語言模型（LLM）與 Python 演算法的混合式智慧排班系統。它旨在自動化和簡化複雜的員工排班流程，從解析人類的自然語言請假請求，到產生公平且最佳化的班表，最後再匯出易於閱讀的報告與 Excel 檔案。

## 核心功能

- **自然語言理解**: 使用 Google Gemini 模型解析非結構化的文字請假請求，並轉換為結構化資料。
- **公平性演算法**: 內建的 Python 演算法會根據預設規則（如每日人力需求、週末班、大夜班）和員工個人限制，以公平為原則進行排班。
- **互動式微調**: 在產生初版班表後，提供一個命令列介面，讓管理者可以進行臨時的單人調班，系統會以「最小幅度影響」的原則尋找替代人選。
- **AI 摘要報告**: 再次利用 Gemini 模型，將最終的排班數據轉換為人性化的管理摘要，並自動草擬給員工的通知訊息。
- **Excel 匯出**: 將總班表及每位員��的個人班表匯出為格式精美的 Excel 檔案，方便分發與歸檔。

## 系統工作流程

1.  **解析請求**: 程式啟動，讀取在 `main.py` 中硬編碼的自然語言請假請求。
2.  **產生初版班表**: 呼叫 Gemini API 將請求轉換為結構化資料，並執行 Python 演算法產生一份基於公平性原則的初版班表 (`shift_schedule_initial.xlsx`)。
3.  **手動調整**: 程式會詢問使用者是否需要手動調班。使用者可以輸入要請假的員工和日期，系統會自動尋找最適合的代班人選並更新班表。此步驟可重複進行。
4.  **產生最終報告**: 所有調整完成後，程式會將最終的班表數據傳送給 Gemini，生成一份包含公平性分析和員工通知範例的摘要報告。
5.  **匯出最終班表**: 將最終確認的班表匯出為 `shift_schedule_final.xlsx`。

## 環境建置與執行

### 1. 前置需求

- **Python**: 建議使用 Python 3.8 或更高版本。
- **uv**: 本專案使用 `uv` 作為虛擬環境與套件管理工具。
- **Gemini API 金鑰**: 你必須擁有一個 Google Gemini API 金鑰。

### 2. 安裝步驟

首先，建立並��用虛擬環境，然後安裝所需的依賴套件：

```bash
# 從 pyproject.toml 安裝依賴套件
uv pip install -r requirements.txt
```

接著，設定您的 API 金鑰。建議使用環境變數來管理金鑰，避免洩漏。

**Windows (Command Prompt):**
```cmd
set GEMINI_API_KEY="YOUR_API_KEY"
```

**Windows (PowerShell):**
```powershell
$env:GEMINI_API_KEY="YOUR_API_KEY"
```

**Linux / macOS:**
```bash
export GEMINI_API_KEY="YOUR_API_KEY"
```

### 3. 執行程式

完成設定後，透過 `uv` 執行主程式：

```bash
uv run python main.py
```

程式將會開始執行，並在命令列中顯示進度，最終產生輸出檔案。

## 程式設定

主要的設定都集中在 `main.py` 檔案的開頭，您可以直接修改以符合您的需求：

- `EMPLOYEES`: 員工清單。
- `SHIFT_TYPES`: 所有的班別類型。
- `SHIFTS_PER_DAY_REQUIREMENTS`: 設定平日與週末各班別所需的人力。
- `EMPLOYEE_CONSTRAINTS`: 設定個別員工的特殊班別限制。目前程式會隨機指派限制給部分員工，您可以將其修改為固定設定。
- `raw_requests`: 在 `if __name__ == "__main__":` 區���中，您可以修改這個字串來模擬不同的自然語言請假請求。

## 輸出檔案

程式執行後會產生以下檔案：

- `employee_constraints.txt`: 當次執行時，隨機產生的員工班別限制條件，方便追蹤與除錯。
- `shift_schedule_initial.xlsx`: 根據初始條件產生的第一版班表。
- `shift_schedule_final.xlsx`: 經過所有手動調整後，最終確認的班表。
