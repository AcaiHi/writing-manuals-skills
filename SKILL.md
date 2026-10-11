---
name: writing-manuals
description: >
  把系統的截圖與使用故事寫成「看起來像純文字檔」的 Word 使用說明書或結案手冊，並產生 docx 與 pdf：
  章節用 Markdown 撰寫，工具負責純文字外觀的排版、「圖 N.M. 名稱」圖說、照片來源、固定目錄與頁碼、
  每頁填滿的檢查。當使用者說「寫使用手冊」「寫使用說明書」「結案手冊」「依截圖寫操作說明」
  「Word 手冊」「手冊排版」「手冊目錄頁碼」「手冊有幾頁空白」，或要為已完成的系統產出可交付的
  docx / pdf 手冊時使用。需要 Windows、Microsoft Word 與 pywin32 才能產生 PDF 與目錄頁碼。
---

# writing-manuals

目標：把系統的使用故事與截圖，寫成一份可交付的 Word 使用說明書。文件刻意做成**純文字外觀**：
不用 Word 的標題樣式、大綱、框線、底色，版面只靠字型、字級、粗體、顏色與文字符號。

```
docs/manual/
├─ manual.config.json        書名、單位、作者、日期、網址、用語表
├─ manual.style.json         排版偏好（字級、縮排、行距、對齊、外框顏色…）
├─ style.md                  寫作與內容偏好（文體、取捨、用語），可指定規範
├─ chapters/NN-<章名>.md     每章一個檔，連續段落
├─ figures/                  手冊用圖（標註圖或原始截圖）
├─ figure-plan.md            圖表計畫：每章幾張圖、哪張圖放哪裡
└─ <書名>.docx / .pdf        產出
```

腳本都在這個 skill 的 `scripts/`，**一律在專案根目錄執行**（預設設定檔路徑是相對於目前資料夾的 `docs/manual/manual.config.json`；`<skill>` 代表本 skill 的資料夾）：

```
python <skill>/scripts/build_manual.py    # 產生 docx 與 pdf
python <skill>/scripts/check_manual.py    # 檢查章節、圖號、用語、格式
python <skill>/scripts/measure_blank.py   # 量每頁空白，列出要補內容的頁
```

套件：python-docx、Pillow、PyMuPDF（`pip install python-docx pillow pymupdf`）；產生 PDF 與目錄頁碼另需 Windows、Word 與 pywin32。環境細節見 [reference/pipeline.md](reference/pipeline.md)。

## 撰寫：複製這份 todo 清單並逐項勾選

```
手冊進度：
- [ ] 1. 材料：讀使用故事、截圖清單、必要時讀原始碼確認系統行為
- [ ] 2. 設定：建立 manual.config.json、manual.style.json、style.md，向使用者確認書名、單位、作者、日期、網址與偏好
- [ ] 3. 計畫：規劃章節與圖表計畫（figure-plan.md），等使用者確認
- [ ] 4. 撰寫：每章一個 md 檔，圖與圖說就位
- [ ] 5. 建置：build_manual.py
- [ ] 6. 檢查：check_manual.py，修到沒有錯誤；新手試讀一輪，依報告修改
- [ ] 6b. 同步：圖有重拍或改標籤時，更新 _image-changes.md 並同步內文
- [ ] 7. 填滿：measure_blank.py，補說明文字，重建、重量，直到頁面填滿
- [ ] 8. 驗收：打開 PDF 抽看頁面，核對目錄頁碼，回報
```

**1. 材料** — 章節順序沿用使用故事。每張圖要寫什麼，以畫面上看到的與原始碼的實際行為為準，不憑印象。
截圖還沒拍或不完整時，先把截圖補齊再寫（見 `capturing-journeys` skill）。

**2. 設定** — 複製 `templates/` 裡的 `manual.config.json`、`manual.style.json`、`style.md` 到 `docs/manual/`，填入內容；欄位說明見 [reference/pipeline.md](reference/pipeline.md)。
問使用者有沒有自己的偏好，特別是**寫作風格要依什麼來源**：使用者自己的 prompt、某個 skill、一份範例文件或舊手冊、一份寫作指南，或單純口頭描述，任何來源都行。排版數值改 `manual.style.json`，寫作風格與內容取捨記進 `style.md`（記下來源與要點，下次直接沿用）。專案已有這些檔就沿用，不重建。

**3. 計畫** — 列出每章的圖號數與每張圖的檔案、圖名、放在哪一節，給使用者確認。規則見 [reference/figures.md](reference/figures.md)。
使用者已預先確認，或明說非互動、不要停下來問時，視為已確認，把計畫寫進回報；**非互動只是不停下來問，不是省略步驟**：新手試讀與圖文同步一樣要做，需要使用者決定的事（例如圖上有人名）列在最後的回報裡。

**4. 撰寫** — **先讀 `docs/manual/style.md`，照它寫**；它指定了寫作風格的來源（某個 skill、使用者的 prompt、範例文件…）就先載入或讀進來再寫，使用者當下給的指示優先。章節語法見 [reference/chapter-syntax.md](reference/chapter-syntax.md)，通用的寫作方法見 [reference/writing.md](reference/writing.md)。從 `templates/chapter.md` 起手。章檔編號從 `00-前言.md` 起連續，沒有前言也要有 00 號檔。

**4b. 分工** — 章節多時分給子代理平行寫（每個 1 到 2 章），章節少就自己寫；派工內容、新手試讀、圖文同步的做法見 [reference/workflow.md](reference/workflow.md)。

**5. 建置** — `build_manual.py`。產出放在 `docs/manual/`；舊檔被 Word 或 Acrobat 開著無法覆寫時，只回報、請使用者關閉，**不要強制結束任何程式**。

**6. 檢查** — `check_manual.py` 列出的錯誤逐項修，再跑一次。可機械修正的（圖片路徑、括號）用 `--fix`。

**7. 填滿** — 每頁都要盡量填滿。`measure_blank.py` 會列出空白超過門檻的頁、所屬章節與建議補的字數；補的是**說明文字**，不是放大圖片、插入分頁或空白行。補完重建再量，直到除封面、目錄與最後一頁外，每頁空白不超過約 10%。

**8. 驗收** — 把 PDF 抽幾頁轉成圖片實際看（封面、目錄、一個含圖的頁、一個含表格的頁、最後一頁），並逐條核對目錄頁碼與 PDF 的實際頁碼一致。

## 偏好與預設

排版與寫作偏好都放在專案裡的檔案，改檔即可，不必改腳本：

- `manual.style.json`：字級、首行縮排、對齊、行距、段距、圖外框顏色與粗細、頁邊界、圖高上限、照片來源字級與顏色。
- `style.md`：文體、敘事、內容取捨（要不要放某類畫面、要不要出現人名等）、用語、頁面填滿的標準。

工具內建的結構（固定，不是偏好）：所有段落是 Normal 樣式、不用 Word 標題樣式與自動目錄；圖說「圖 N.M. 名稱」置於圖下方；固定文字目錄加實測頁碼。這些的理由與細節見 [reference/format.md](reference/format.md)；所有寫死的預設與連動關係列在 [reference/defaults.md](reference/defaults.md)，要改腳本前先讀。

## 更多

- 流程與分工（平行撰寫、新手試讀、圖文同步、填頁）：[reference/workflow.md](reference/workflow.md)
- 版面規格與可調參數：[reference/format.md](reference/format.md)
- 內建預設與修改腳本前的注意事項（要改腳本時先讀）：[reference/defaults.md](reference/defaults.md)
- 章節語法：[reference/chapter-syntax.md](reference/chapter-syntax.md)
- 寫作原則：[reference/writing.md](reference/writing.md)
- 圖與圖說：[reference/figures.md](reference/figures.md)
- 建置、設定與環境：[reference/pipeline.md](reference/pipeline.md)
