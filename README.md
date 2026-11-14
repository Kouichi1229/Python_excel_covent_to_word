# Excel 統計資料自動轉換 Word 報表工具

本 Repository 保存一組以 Python 將 Excel 統計工作表套用 Word 範本、自動產生格式化報表的歷史工具。程式主要面向大學院級與系所層級的學術統計資料，包含論文數、論文成長率、被引用次數、平均被引用次數及 H 指數等報表。

> 本文件是依 Repository 內的程式、範本、輸入檔與操作檔整理的技術說明。這是歷史專案回顧，並未在目前環境重新執行或驗證輸出結果。Repository 中的範例 Excel／Word 檔可能包含機構資料，使用前應確認其公開與再散布權限。

## 專案功能

目前版本包含三組 Excel → Word 報表流程：

1. **院級統計報表**：讀取院級 Excel 工作簿，依七種工作表類型套用 Word 範本，產生七份院級統計報表。
2. **系所統計報表**：讀取系所統計 Excel 工作簿，產生五種學院／系所論文統計報表。
3. **系所論文數與被引次數報表**：分別讀取論文數及被引次數統計工作簿，逐一處理 16 個學院／單位工作表，輸出以單位名稱命名的 Word 文件。

## 技術架構

| 類別 | 技術／工具 | 程式中的用途 |
|---|---|---|
| 程式語言 | Python | 流程控制、資料整理及報表產生 |
| Excel 讀取 | openpyxl | 載入工作簿、選取工作表及讀取儲存格 |
| Word 範本 | docxtpl | 將資料 context 填入 DOCX 範本並輸出文件 |
| Word 文件 | python-docx | 程式匯入相關文件操作模組 |
| 圖像／Office 整合 | Pillow、pywin32、ImageGrab | 程式匯入相關模組；目前已核對的主要報表流程未見明確使用這些功能 |
| 執行環境 | Windows 批次檔 | 透過 .bat 啟動 Python 程式 |

## 程式架構

| 目錄 | 主要程式 | 職責 |
|---|---|---|
| `台大院級word表格/` | `Collage_Main_app.py` | 依序呼叫七種院級報表產生函式 |
| 同上 | `Collage_Paste_action.py` | 讀取院級 Excel 工作表，組合資料 context，填入七種 Word 範本 |
| `台大系所word表格/` | `Department_Main_app.py` | 依序呼叫五種系所報表產生函式 |
| 同上 | `Department_Paste_action.py` | 讀取系所 Excel 工作表，填入五種 Word 範本 |
| `台大系論文及被引次數/系所論文數統計/` | `Paper_counts_main.py`、`paste_action.py` | 逐一讀取 16 個單位的論文數工作表並產生 DOCX |
| `台大系論文及被引次數/系所被引用次數/` | `Rn_counts_main.py`、`paste_action.py` | 逐一讀取 16 個單位的被引次數工作表並產生 DOCX |

## 主要處理流程

1. 以 openpyxl 載入指定 Excel 檔案。
2. 依固定工作表名稱選取資料。
3. 從工作表第 2 列讀取表頭，從第 3 列起逐列整理欄位值（部分報表函式採不同起始列）。
4. 將資料組成 Python list／dictionary，作為 docxtpl 的 template context。
5. 載入對應 DOCX 範本，執行 `render(context)`。
6. 儲存為 Word 文件；系所論文數與被引次數流程會以學院／單位名稱組成輸出檔名。

## 執行方式與限制

- 部分目錄附有 Windows `.bat` 啟動檔；院級工具另附套件安裝批次檔。
- 程式以相對路徑讀取 Excel 與 `template\` 目錄，執行時需要在對應專案目錄啟動，並確保輸入檔名及範本路徑存在。
- 院級目錄的 `requirements.txt` 目前只列出 `openpyxl` 與 `docxtpl`；其他子專案的依賴未集中管理。
- 目前程式使用固定輸入檔名與工作表名稱，並非可任意上傳 Excel 的通用轉檔服務。
- 目前未見完整的輸入格式驗證、錯誤回報、輸出目錄管理或自動化測試。
- 程式碼包含多個未使用或疑似歷史遺留的 import；本次整理不擅自移除，以免改變舊環境相容性。

## 技術回顧

- [PROJECT_RETROSPECTIVE.md](PROJECT_RETROSPECTIVE.md)：整理實際實作能力、模組流程、履歷描述及可確認的技術限制。

## 專案定位

**Python 文件自動化、Excel 資料讀取、Word 範本渲染及批次報表產生。**
