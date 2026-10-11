# 建置、設定與環境

## 目錄
- 環境需求
- 設定檔 manual.config.json
- 三支腳本
- 建置怎麼運作
- 迭代：建置、檢查、量測
- 驗收 PDF
- 常見問題

## 環境需求

| 項目 | 用途 | 沒有時 |
|---|---|---|
| Python 3.9 以上 | 執行腳本 | 無法執行 |
| python-docx、Pillow | 產生 docx、處理圖片 | `pip install python-docx pillow` |
| PyMuPDF | `measure_blank.py` 量空白 | `pip install pymupdf` |
| Windows + Microsoft Word + pywin32 | 實測目錄頁碼、匯出 PDF（內嵌字型） | 仍可產生 docx，但目錄頁碼是占位數字、不產生 PDF；腳本會明確提示 |
| 微軟正黑體、Arial | 預設字型 | Windows 與 Office 內建 |

## 設定檔 manual.config.json

放在 `docs/manual/manual.config.json`，從 `templates/manual.config.json` 複製。缺少的鍵用預設值。

| 鍵 | 說明 |
|---|---|
| `title` | 書名；輸出檔名就是 `<title>.docx` 與 `<title>.pdf` |
| `subtitle` | 封面副標；留空就不印（封面位置不變） |
| `org` | 單位，印在封面 |
| `author` | docx 的作者與最後修改者（PDF 也帶這個作者） |
| `versionDate` | 版本日期，印在封面 |
| `url` | 系統網址，印在封面；沒有就留空 |
| `accentColor` | 章名的顏色（6 位十六進位，不含 `#`） |
| `chaptersDir`、`figuresDir`、`outDir` | 章節、圖、輸出資料夾，相對於 `docs/manual/` |
| `bodyFontEastAsia`、`bodyFontLatin` | 內文字型，預設微軟正黑體與 Arial |
| `terms` | 用語表：`{"錯誤用語": "正確用語"}`；值為 `null` 表示禁用此詞 |
| `figurePlan` | 圖表計畫檔路徑，空就略過相關檢查 |
| `photoCredits` | 照片來源清單檔路徑，空就略過相關檢查 |
| `annotations` | 標註框檔案的 glob（例如 `docs/ui-inventory/states/*.boxes.json`），用來檢查圖上標籤；空就略過 |

相對路徑都相對於設定檔所在資料夾；值為 `null` 用預設值；未知的鍵只提醒、不報錯。

書名、單位、作者、日期、網址這幾項向使用者確認後再填，不要自己編。

## 偏好檔

| 檔案 | 誰讀 | 放什麼 |
|---|---|---|
| `manual.config.json` | 腳本 | 書名、單位、作者、日期、網址、字型、章名色、用語表、各種路徑 |
| `manual.style.json` | 腳本 | 排版數值：字級、縮排、對齊、行距、段距、圖外框、頁邊界 |
| `style.md` | Claude | 文體、敘事、內容取捨、用語、頁面標準；可寫「遵守某個 skill」 |

偏好有三種給法，隨使用者方便：口頭說（Claude 記進對應檔）、直接改檔、或在 `style.md` 指定要遵守的 skill 或文件。

## 三支腳本

都在專案根目錄執行；`--config` 預設是 `docs/manual/manual.config.json`。

```
python <skill>/scripts/build_manual.py   [--config PATH] [--style PATH] [--out DIR] [--chapters DIR] [--no-pdf]
python <skill>/scripts/check_manual.py   [--config PATH] [--fix]
python <skill>/scripts/measure_blank.py  [PDF] [--config PATH] [--threshold PCT] [--gap-cm CM] [--json] [--chapters DIR]
```

- `--style PATH`：排版偏好檔，預設是設定檔同資料夾的 `manual.style.json`；沒有這個檔就用內建預設。鍵見 [format.md](format.md)。
- `--out DIR`：輸出到別的資料夾（試做、比較時用），不動 `docs/manual/` 裡的正式檔。
- `--no-pdf`：只產生 docx，略過 Word；沒有 Word 的環境，或只想快速看文字時用。
- `--threshold`：空白比例門檻，預設 10%。
- `--gap-cm`：頁面內部的空白段超過幾公分才算（例如圖被擠到下一頁，上方留下的空白）。

## 建置怎麼運作

1. 讀 `chapters/NN-*.md`，依語法轉成 docx（所有段落是 Normal 樣式，格式見 [format.md](format.md)）。
2. 第一階段：目錄頁碼先用占位數字，產生 docx。
3. 用 Word 開啟，量測每章起始頁碼。
4. 第二階段：把實測頁碼寫回目錄，重新產生 docx。
5. 用 Word 匯出 PDF（內嵌字型）。

檔案被 Word 或 Acrobat 開著而無法覆寫時，腳本只回報並停止。請使用者關閉後再跑，不要強制結束程式。

## 迭代：建置、檢查、量測

```
build_manual.py  ->  check_manual.py  ->  measure_blank.py
       ^                                        |
       +---------- 補說明文字 <-----------------+
```

- `check_manual.py`：章檔編號連續、每檔一個章標題且章號與檔號一致、節編號、表格欄數、圖說與圖號連續、圖有被引用、圖表計畫一致、照片來源格式、用語、標點與排版規則。逐項修到沒有錯誤。
- `measure_blank.py`：封面與目錄不算，最後一頁顯示但不計入。輸出每頁空白比例、頁尾空白、頁內大空白、建議補的中文字數（約 1 cm 空白對應 60 字，標題、圖說、圖需要的空間更多，只當參考）。
- 補內容後頁面會重新分配，一定要重建重量。

## 驗收 PDF

- 用 PyMuPDF 或其他工具把 PDF 抽幾頁轉成圖片實際看：封面、目錄、含圖的頁、含表格的頁、最後一頁。
- 逐條核對目錄頁碼與 PDF 的實際頁碼。
- 確認圖與圖說同頁、圖說沒有單獨落在頁首。
- 圖片外框是淡灰色，不是黑色。

## 常見問題

- **Windows 終端機裡中文變亂碼**：這是終端機編碼，不是檔案壞了。執行腳本前加 `PYTHONIOENCODING=utf-8`，或把輸出導向檔案再用編輯器讀；JSON 與章節檔都是 UTF-8。
- **PDF 的字型不對**：確認本機有微軟正黑體與 Arial；PDF 由 Word 匯出，才會內嵌字型。
- **特殊符號（例如「⋮」）變成別的字型**：Word 會替換成別的字型；先小量試印確認，必要時換成其他符號。
- **多張圖跨頁時，圖說和圖分開**：圖片段落已設與下段同頁；仍分開時，檢查圖片是否過高（上限 21.5 cm）。
- **每章換頁留下大片空白**：不要用分頁；章節接續排。
- **改了章節內容但 PDF 沒變**：PDF 是 `build_manual.py` 最後一步匯出的；只重建 docx（`--no-pdf`）不會更新 PDF。
