# 內建預設與修改腳本前的注意事項

這份列出**寫死在 `scripts/build_manual.py` 裡的預設**，以及改它們之前要知道的事。
要改腳本的人（或 Claude）先讀這份，避免改了 A 忘了連動的 B。位置以函式與常數名稱為準，行號會變。

## 目錄
- 怎麼判斷該改哪裡
- 可用樣式檔改的預設
- 寫死的預設（要改只能改腳本）
- 連動關係：改了 A 要同步改 B
- 兩階段建置與 Word
- 改完怎麼驗證

## 怎麼判斷該改哪裡

| 想改的東西 | 改哪裡 |
|---|---|
| 書名、單位、作者、網址、字型、章名色、用語、路徑 | `manual.config.json` |
| 字級、縮排、對齊、行距、段距、圖外框、頁邊界、表格內距與自動調整 | `manual.style.json` |
| 文體、內容取捨、用語、頁面標準 | `style.md` |
| 下面「寫死的預設」那些 | 修改 `scripts/build_manual.py` |

## 可用樣式檔改的預設

鍵、預設值與意義見 [format.md](format.md) 的「可調的排版偏好」。腳本內對應 `DEFAULT_STYLE`、`apply_style()`。

## 寫死的預設（要改只能改腳本）

| 預設 | 內容 | 位置 | 為什麼這樣 |
|---|---|---|---|
| 允許標點符號溢出邊界 | **關閉**（`overflowPunct=0`） | `make_ppr_base()`，套在 docDefaults 與 Normal | 使用者明確要求；標點不可超出邊界 |
| 文件格線時貼齊格線 | **關閉**（`snapToGrid=0`），頁面設定不含 `docGrid` | `make_ppr_base()`、`setup_styles()` | 使用者明確要求；避免行距被格線改動 |
| 只用 Normal 樣式 | 所有段落都是 Normal，不用標題樣式、大綱階層、段落框線、底色 | 全檔：`plain_paragraph()`、`add_heading()` | 「看起來像純文字檔」的核心設計 |
| 字型不設在 run 上 | 字型只設在 docDefaults 與 Normal，run 繼承 | `make_rfonts()`、`setup_styles()` | 使用者在 Word 新增文字時字型一致 |
| 語言標記 | 東亞語言 zh-TW，佈景字型語言 zh-TW | `setup_styles()` | 範本預設是 ja-JP，會讓中文換成日文字型 |
| 左右對齊的最後一行 | 手動換行不拉開（`doNotExpandShiftReturn`） | `setup_styles()` | 避免表格內換行行被拉開 |
| 圖片段落與圖說 | 置中、無首行縮排；圖片段落「與下段同頁」 | `add_image()`、`add_caption()`、`no_indent()` | 圖說不單獨落在下一頁頂端 |
| 圖片外框 | 畫在圖片本身（`a:ln`），不是段落框線；圖縮小一個外框寬 | `add_picture_outline()`、`add_image()` | 外框不超出文字寬度 |
| 圖說格式 | 「圖 N.M. 名稱」，在圖正下方 | `CAPTION_RE`、`add_caption()` | 使用者指定；不用【】、不用子圖 |
| 照片來源 | 圖說下一行、置中、無縮排、不跨頁、網址是可點連結且只用直接格式 | `add_citation()`、`add_hyperlink()` | 使用者指定；圖、圖說、來源同頁 |
| 章標題 | 18 pt 粗體、章名色，後接一條「─」線（字數見樣式檔 `ruleChars`） | `add_heading()`（level 1）、`RULE_SIZE`、`RULE_N` | 見下方「連動關係」 |
| 節標題 | 「▍」加粗體，字級 `h2Size` | `add_heading()`（level 2） | 純文字裝飾 |
| 框標題 | 「── 本章重點 ──」「── 常見問題 ──」，12 pt，尾端「─」字數見樣式檔 `boxRuleChars` | `BOX_TITLES`、`add_heading()` | 純文字裝飾 |
| 小節標題 | 「■ 」加粗體 | `add_heading()`（level 3） | 純文字裝飾 |
| 章節不換頁 | 章標題前只留一個空白段落；只有封面到目錄、目錄到正文分頁 | `add_heading()`、`build_docx()` | 使用者要求，避免每章尾端留白 |
| 問與答 | 凸排：左縮與凸出各 2 字元，無首行縮排 | `hanging_indent()` | 使用者指定 |
| 目錄 | 固定文字、單一靠右點狀前導線定位點、頁碼兩階段回填 | `build_docx()`、`measure_pages()` | 不用 Word 自動目錄（會套「目錄 1」樣式） |
| 表格 | 無表格樣式；0.5 pt #808080 單線框；寬度等於文字寬度；首列粗體並跨頁重複；列不跨頁；表後留一個空白段 | `add_table()`、`TABLE_BORDER_SZ`、`TABLE_BORDER_RGB` | 純文字外觀 |
| 文件屬性 | 作者與最後修改者填設定檔 `author` | `build_docx()` | 避免留下 python-docx |
| 色彩 | 章名色來自 `accentColor`，節名色是其 0.75 倍；其餘深字、灰線、次要字、注意紅是固定色碼 | `apply_config()`、`darker()`、`INK`、`GREY`、`MUTED`、`ALERT` | 預設色票 |
| 封面、目錄標題 | 目錄標題「目　錄」 | `TOC_TITLE` | — |
| 紙張 | A4 直式 | `TEXT_W` 計算處 | — |

## 連動關係：改了 A 要同步改 B

- **分隔線字數**：`ruleChars`（73）與 `boxRuleChars`（46）在樣式檔，依 9 pt 與預設字型量到剛好等於文字寬度。實測換成標楷體加 Times New Roman 後，73 與 46 仍貼齊邊界（468 與 465，文字寬 471 pt）；改字級、頁邊界或換成寬度不同的字型後，要輸出 PDF 量一次再調這兩個數字。
- **字級與縮排**：`firstLineIndentChars` 同時決定首行縮排與問答凸排；twips 後備值用 `bodySize` 乘以 20 計算，改 `bodySize` 時這個後備值會自動跟著算。
- **表格寬度**：欄寬在建置時依內容與 `TEXT_W`、儲存格內距計算；改頁邊界或內距後要重建才會重算。表格靠左邊緣的位置與左內距有關（Word 2010 相容模式下，左框線在表格縮排左邊一個左內距的位置）。
- **圖高上限與外框**：圖會縮小一個外框寬，改外框粗細時文字寬度內的圖寬隨之改變。
- **標點與格線**：兩者要同時設在 docDefaults 與 Normal；只改一處，Word 開啟後有些段落會繼承到範本預設。
- **字型與頁數**：換字型會改變每行字數，頁數與目錄頁碼都會變，必須重新建置（兩階段）。

## 兩階段建置與 Word

目錄頁碼要靠 Word 實測，所以建置是兩階段：先用占位頁碼產生 docx，用 Word 量每章起始頁，再把頁碼寫回目錄重新產生。
- 改動會影響頁數的任何預設（字級、字型、行距、邊界、圖高、補字），都要走完兩階段，單看 `--no-pdf` 的 docx 目錄頁碼不準。
- 沒有 Word 或 pywin32 時只能產出 docx 與占位頁碼。

## 改完怎麼驗證

1. `python <skill>/scripts/build_manual.py --out <暫存資料夾>`，不要直接覆蓋正式檔。
2. 用 python-docx 檢查：全部段落是 Normal 樣式、docDefaults 與 Normal 都有 `overflowPunct=0` 與 `snapToGrid=0`、頁面沒有 `docGrid`。
3. 用 Word 開啟，逐段確認「允許標點符號溢出邊界」與「貼齊格線」都是關閉的。
4. 抽看 PDF 的封面、目錄、含圖頁、含表格頁；核對目錄頁碼。
5. 跑 `check_manual.py` 與 `measure_blank.py`。
