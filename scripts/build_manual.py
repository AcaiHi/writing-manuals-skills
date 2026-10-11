"""由章節 Markdown 產生 Word 使用說明書（DOCX）與 PDF。

要修改這支腳本前，先讀 reference/defaults.md：所有寫死的預設（例如關閉「允許標點符號溢出邊界」與
「貼齊格線」）、位置與連動關係都列在那裡。

用途
    讀取 <chaptersDir>/NN-*.md（NN 為兩位數章檔編號），產生「純文字外觀」的說明書：
    封面、固定文字目錄（含點狀前導線與頁碼）、各章內文。PDF 與目錄頁碼由 Microsoft Word
    量測與匯出（兩階段：第一階段以占位頁碼 00 建檔交 Word 量測各章起始頁，第二階段寫入實際頁碼）。

用法（在專案根目錄執行）
    python build_manual.py [--config PATH] [--style PATH] [--out DIR] [--chapters DIR] [--no-pdf]
      --config PATH   設定檔，預設 docs/manual/manual.config.json
      --out DIR       輸出資料夾（覆寫設定檔 outDir），測試時請指向暫存資料夾
      --chapters DIR  章節來源資料夾（覆寫設定檔 chaptersDir）
      --no-pdf        不啟動 Word：只產生 DOCX，目錄頁碼維持占位符 00，不產生 PDF
    輸出檔名 = 設定檔 title + ".docx" / ".pdf"。
    設定檔內的相對路徑（chaptersDir、figuresDir、outDir）皆相對於設定檔所在資料夾。

需要的套件
    python-docx、Pillow（必要）；pywin32 + Microsoft Word（選用，量測目錄頁碼並匯出 PDF，
    字型由 Word 的 PDF 匯出嵌入）。沒有 pywin32 或 Word 時仍會產生 DOCX，並明確提示
    「未產生 PDF 與目錄頁碼」，退出碼為 0。

退出碼
    0 成功（含無 Word 時只產生 DOCX）；1 設定檔錯誤、找不到章節、圖片不存在、
    圖片重複使用、表格欄數超過表頭、照片來源行位置錯誤；2 輸出檔被其他程式（Word、Acrobat）
    開啟而無法覆寫（只回報，不會關閉任何程式，也不會改動任何檔案）。

格式規則（使用者指定的規格，不可改動語意）
DOCX 刻意為「純文字」：每個段落都是 Normal 樣式，不帶任何段落格式（無大綱階層、段距、網底、
框線、色彩、字元樣式），以下列例外為限。唯一的格式是文字層級（run）的字級、粗體與色彩
（主色章名、「▍」節標記、「── 本章重點 ──」框線、灰色「─」字元分隔線、「注意：」「問：」
「答：」引導詞；全為實際字元，不用框線）。字型不設在 run 上，而由 Normal 樣式 / docDefaults
繼承（東亞文字 = bodyFontEastAsia，西文與數字 = bodyFontLatin）。段前段後 = 0、單行行距、
取消「貼齊格線」同樣只設一次於 docDefaults + Normal（節區不帶 w:docGrid）；取消「允許標點
符號溢出邊界」（w:overflowPunct=0）、左右對齊（w:jc="both"，每段繼承；左右對齊不會拉伸
段落的最後一行）、首行縮排 2 字元（w:ind w:firstLineChars=200，並以 w:firstLine = 2 x 內文字級
（twips）作為後備）只設在 Normal：每個內文段落（敘述、注意框）都繼承縮排。
每張圖片帶 1 pt 灰色（#A6A6A6）外框（圖片本身 pic:spPr 中的 a:ln，屬繪圖層級，不是段落框線）；
圖片縮小 1 pt 讓外框仍落在文字寬度內。
各章不換頁：章標題直接接在前文之後，前面以一個空白 Normal 段落隔開（不使用段距）。全書只有
封面 -> 目錄、目錄 -> 內文兩處分頁（明確的分頁字元，後者位於第一章標題段落中）。

「不帶段落格式」的例外：
(0h) 問答段落（章節 md 中以「問：」「答：」開頭的行）使用凸排：w:ind leftChars=200
     hangingChars=200（twips 後備 left=880 / hanging=440，與 Word 寫入值相同），不設 firstLine，
     換行後的文字對齊於 2 字元標籤之後。僅限這些段落。
(0)  不可縮排的段落帶明確覆寫 w:ind firstLine=0 firstLineChars=0：圖片段落、「圖 N.M.」圖說
     （兩者置中，縮排會使其偏離中心）、章節標題、字元分隔線、封面各行、目錄各行、空白 / 分頁
     段落與頁尾頁碼段落（置中，w:jc center）。
(a)  圖片段落與圖說段落（「圖 N.M. 名稱」）一律置中（w:jc center）；圖片段落帶「與下段同頁」，
     圖說不會單獨落到下一頁頂端。
(b)  固定文字目錄：每行為「標題<TAB>頁碼」，在右側文字邊界設單一靠右、點狀前導線的定位點
     （僅這些段落有 w:tabs）。頁碼由 Word 量測（第一階段占位 00），第二階段寫入。
(c)  照片來源行（章節 md 中緊接在「圖 N.M.」圖說之後的「> 照片來源：名稱，作者，授權，網址」，
     多筆以「；」分隔）：置中、4 pt（run 字級與段落標記字級）、灰色 #A6A6A6（run 層級 w:color）、
     無首行縮排、段落中不分頁；圖說帶「與下段同頁」，使圖片 + 圖說 + 來源同頁。每個網址為可點選的
     w:hyperlink，其 run 只有直接格式（字級、灰色、底線），不使用 rStyle / Hyperlink 樣式。
     只標註圖中實際可見的 CC BY 照片（CC0 / 公有領域不需標註）；check_manual.py 可依設定檔
     photoCredits 檢查每筆是否與照片來源清單一致。
(d)  文件屬性：作者 / 最後修改者 = author（匯出的 PDF 也帶此資訊），標題 = title。
(e)  簡單表格（章節檔中的 GitHub 式 Markdown 表格）：Word 表格，只用直接格式 -- 無表格樣式
     （無 w:tblStyle、w:tblLook），所有外框與內線為 0.5 pt #808080 單線，寬度 = 文字寬度，
     固定欄寬依儲存格內容計算，無網底。表頭列（第一列）粗體、跨頁時每頁重複（w:tblHeader），
     並與第一個內容列同頁（其儲存格段落帶 keepNext）；列不跨頁分割（w:cantSplit）。儲存格段落為
     Normal 文字並覆寫首行縮排為 0（同圖說）；字級 TABLE_SIZE。表格後接一個空白 Normal 段落。
     `###` 行為第三層標題（粗體 run，無縮排）。

常數意義
    MARGIN               上下左右頁邊界 2.2 cm（A4 直式 21.0 x 29.7 cm）
    TEXT_W               文字寬度 = 頁寬 - 左右邊界（EMU）
    MAX_IMG_H            圖片最大高度 21.5 cm（超過則依比例縮小）
    OUTLINE_W            圖片外框粗細，12700 EMU = 1 pt
    OUTLINE_RGB          圖片外框顏色 #A6A6A6（Word「白色，背景1，較深35%」）
    TOC_TITLE            目錄頁標題「目　錄」
    H1/H2/H3_SIZE        章 / 節 / 小節標題字級 18 / 14 / 12 pt
    BODY_SIZE            內文字級 11 pt（也決定首行縮排與凸排的 twips 後備值）
    CAPTION_SIZE         圖說字級 10 pt
    CITE_SIZE            照片來源字級 4 pt
    TABLE_SIZE           表格儲存格字級 10 pt
    TABLE_BORDER_SZ/RGB  表格線 w:sz=4（1/8 pt 為單位，即 0.5 pt）、顏色 #808080
    TABLE_CELL_MAR       儲存格內距（twips）上、左、下、右，預設全 0（樣式檔 tableCellMarginTwips）
    TABLE_AUTOFIT        表格是否隨內容自動調整；預設 False = 固定欄寬（樣式檔 tableAutofit）
    CITE_RGB             照片來源文字與連結顏色 #A6A6A6
    INDENT_CHARS         內文首行縮排字元數 2
    TEAL                 主色（章名、節標記、圖號），取自設定檔 accentColor，預設 028090
    TEAL_D               主色較深版（節名、小節名）= 主色各分量 x 0.75（預設即 01606C）
    INK / GREY / MUTED / ALERT  深色文字 16323A / 灰色分隔線 A6A6A6 / 次要文字 5B747C / 注意紅 C2372F
    RULE_SIZE, RULE_N    分隔線字級 9 pt、73 個「─」= 文字寬度（改字級時需重新量測）
    BOX_TITLES           以「── 名稱 ────」框線呈現的二階標題：本章重點、常見問題（12 pt，46 個「─」）
    CAPTION_RE           圖說「圖 N.M.」
    PHOTO_RE             圖說下方的照片來源行「> 照片來源：」
    URL_RE               網址（僅 ASCII，故「；」「，」會結束網址）
    TABLE_SEP_RE         表格分隔列 |---|:--:|
    DEFAULTS             設定檔各鍵的預設值
"""
import argparse
import json
import re
import sys
import tempfile
import unicodedata
from pathlib import Path

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Emu, Pt, RGBColor
    from docx.text.run import Run
    from PIL import Image
except ImportError as e:  # pragma: no cover
    sys.exit(f"錯誤：缺少必要套件（{e.name}）。請執行：pip install python-docx Pillow")

DEFAULT_CONFIG = Path("docs") / "manual" / "manual.config.json"
DEFAULTS = {
    "title": "使用說明書",
    "subtitle": "",
    "org": "",
    "author": "",
    "versionDate": "",
    "url": "",
    "accentColor": "028090",
    "chaptersDir": "chapters",
    "figuresDir": "figures",
    "outDir": ".",
    "bodyFontEastAsia": "微軟正黑體",
    "bodyFontLatin": "Arial",
    "terms": {},
    "figurePlan": "",
    "photoCredits": "",
    "annotations": "",
}

DEFAULT_STYLE = {
    "bodySize": 11, "h1Size": 18, "h2Size": 14, "h3Size": 12,
    "captionSize": 10, "tableSize": 10, "citeSize": 4, "citeColor": "A6A6A6",
    "firstLineIndentChars": 2, "align": "both",
    "lineSpacing": 1.0, "paragraphSpacingPt": 0,
    "pictureOutlineColor": "A6A6A6", "pictureOutlinePt": 1,
    "marginCm": 2.2, "maxImageHeightCm": 21.5,
    "tableCellMarginTwips": [0, 0, 0, 0],  # 儲存格內距：上、左、下、右（twips，1 pt = 20）
    "ruleChars": 73,  # 章標題下方「─」線的字數：依 9 pt 與字型量到剛好等於文字寬度；換字型、字級或邊界後要重調
    "boxRuleChars": 46,  # 「── 本章重點 ──」這類框標題尾端「─」的字數；同上
    "tableAutofit": False,  # False = 固定欄寬，大小不隨文字變動；True = 讓 Word 依內容自動調整
}


def load_style(path):
    """Read the optional style JSON (formatting preferences); missing file or keys take DEFAULT_STYLE."""
    style = dict(DEFAULT_STYLE)
    if path is None or not Path(path).is_file():
        return style
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        fail(f"樣式檔 {path} 不是有效的 JSON：{e}")
    if not isinstance(data, dict):
        fail(f"樣式檔 {path} 的最外層必須是物件")
    unknown = sorted(k for k in data if k not in DEFAULT_STYLE and not k.startswith("_"))
    if unknown:
        print(f"提醒：樣式檔有未使用的鍵：{', '.join(unknown)}")
    style.update({k: v for k, v in data.items() if k in DEFAULT_STYLE and v is not None})
    for k in ("citeColor", "pictureOutlineColor"):
        style[k] = str(style[k]).lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", style[k]):
            fail(f"樣式檔 {k} 必須是 6 位十六進位色碼，目前為 {style[k]}")
    m = style["tableCellMarginTwips"]
    if not (isinstance(m, list) and len(m) == 4 and all(isinstance(v, (int, float)) and v >= 0 for v in m)):
        fail("樣式檔 tableCellMarginTwips 必須是 4 個非負數字的陣列：[上, 左, 下, 右]（twips）")
    if style["align"] not in ("both", "left", "center", "right"):
        fail("樣式檔 align 必須是 both / left / center / right 其中之一")
    return style


def apply_style(style):
    global H1_SIZE, H2_SIZE, H3_SIZE, BODY_SIZE, CAPTION_SIZE, CITE_SIZE, TABLE_SIZE, CITE_RGB
    global INDENT_CHARS, ALIGN, LINE_TWIPS, SPACE_TWIPS, OUTLINE_RGB, OUTLINE_W, MARGIN, TEXT_W, MAX_IMG_H
    global TABLE_CELL_MAR, TABLE_AUTOFIT, RULE_N, BOX_RULE_N, RULE
    H1_SIZE, H2_SIZE, H3_SIZE = style["h1Size"], style["h2Size"], style["h3Size"]
    BODY_SIZE, CAPTION_SIZE, CITE_SIZE, TABLE_SIZE = style["bodySize"], style["captionSize"], style["citeSize"], style["tableSize"]
    CITE_RGB, OUTLINE_RGB = style["citeColor"], style["pictureOutlineColor"]
    INDENT_CHARS, ALIGN = style["firstLineIndentChars"], style["align"]
    LINE_TWIPS = int(round(240 * float(style["lineSpacing"])))
    SPACE_TWIPS = int(round(20 * float(style["paragraphSpacingPt"])))
    OUTLINE_W = int(round(12700 * float(style["pictureOutlinePt"])))
    MARGIN = Cm(style["marginCm"])
    TEXT_W = Emu(Cm(21.0) - 2 * MARGIN)
    MAX_IMG_H = Cm(style["maxImageHeightCm"])
    top, left, bottom, right = style["tableCellMarginTwips"]
    TABLE_CELL_MAR = {"top": int(top), "left": int(left), "bottom": int(bottom), "right": int(right)}
    TABLE_AUTOFIT = bool(style["tableAutofit"])
    RULE_N, BOX_RULE_N = int(style["ruleChars"]), int(style["boxRuleChars"])
    RULE = "─" * RULE_N


ALIGN, LINE_TWIPS, SPACE_TWIPS, TABLE_AUTOFIT = "both", 240, 0, False
CHAPTERS = []
OUT_DOCX = OUT_PDF = None
TITLE = SUBTITLE = ORG = AUTHOR = VERSION_DATE = SYSTEM_URL = ""
EA_FONT, LATIN_FONT = "微軟正黑體", "Arial"  # East Asian text / Latin + digits (from the config)
MARGIN = Cm(2.2)
TEXT_W = Emu(Cm(21.0) - 2 * MARGIN)
MAX_IMG_H = Cm(21.5)
OUTLINE_W = 12700  # picture outline: exactly 1 pt (12700 EMU)
OUTLINE_RGB = "A6A6A6"  # neutral grey = Word 「白色，背景1，較深35%」
TOC_TITLE = "目　錄"
H1_SIZE, H2_SIZE, H3_SIZE, BODY_SIZE, CAPTION_SIZE, CITE_SIZE = 18, 14, 12, 11, 10, 4
TABLE_SIZE = 10  # table cell text (run-level size)
TABLE_BORDER_SZ, TABLE_BORDER_RGB = 4, "808080"  # w:sz is in 1/8 pt: 4 = 0.5 pt
TABLE_CELL_MAR = {"top": 0, "left": 0, "bottom": 0, "right": 0}  # twips (set from the style file)
CITE_RGB = "A6A6A6"  # photo-credit text + link colour (run-level w:color)
INDENT_CHARS = 2  # body-text first-line indent, in characters
# palette -- run-level w:color only; TEAL / TEAL_D come from the config's accentColor
TEAL, TEAL_D, INK, GREY, MUTED, ALERT = "028090", "01606C", "16323A", "A6A6A6", "5B747C", "C2372F"
RULE_SIZE, RULE_N, BOX_RULE_N = 9, 73, 46  # 73 / 46 x 「─」 = text width (from the style file ruleChars / boxRuleChars; re-measure if the font or size changes)
RULE = "─" * RULE_N
BOX_TITLES = ("本章重點", "常見問題")  # 「── 名稱 ────」 rule headings (12 pt, 46 x 「─」)
CAPTION_RE = re.compile(r"圖 \d+\.\d+\.")  # figure caption 「圖 N.M. 名稱」
PHOTO_RE = re.compile(r">\s*照片來源：")  # photo-credit line under a caption
URL_RE = re.compile(r"https?://[!-~]+")  # ASCII only, so 「；」「，」 end a URL
TABLE_SEP_RE = re.compile(r"^(?=.*\|)\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$")  # |---|:--:| row


def fail(msg, code=1):
    print(f"錯誤：{msg}", file=sys.stderr)
    sys.exit(code)


def load_config(path):
    """Read the JSON config; missing keys take DEFAULTS.  Returns (cfg, base_dir)."""
    path = Path(path)
    if not path.is_file():
        fail(f"找不到設定檔 {path}（請在專案根目錄執行，預設設定檔是 docs/manual/manual.config.json；沒有的話由 templates/manual.config.json 複製一份，或以 --config 指定）")
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        fail(f"設定檔 {path} 不是有效的 JSON：{e}")
    if not isinstance(data, dict):
        fail(f"設定檔 {path} 的最外層必須是物件")
    unknown = sorted(set(data) - set(DEFAULTS))
    if unknown:
        print(f"提醒：設定檔有未使用的鍵：{', '.join(unknown)}")
    cfg = dict(DEFAULTS)
    cfg.update({k: v for k, v in data.items() if v is not None})  # null = use the default
    color = str(cfg["accentColor"]).lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", color):
        fail(f"設定檔 accentColor 必須是 6 位十六進位色碼（例如 028090），目前為 {cfg['accentColor']!r}")
    cfg["accentColor"] = color
    if not str(cfg["title"]).strip() or re.search(r'[\\/:*?"<>|]', str(cfg["title"])):
        fail("設定檔 title 不可為空，也不可含檔名禁用字元 \\ / : * ? \" < > |（輸出檔名 = title）")
    return cfg, path.resolve().parent


def darker(rgb, f=0.75):
    return "".join(f"{int(int(rgb[i:i + 2], 16) * f):02X}" for i in (0, 2, 4))


def apply_config(cfg):
    global TITLE, SUBTITLE, ORG, AUTHOR, VERSION_DATE, SYSTEM_URL, EA_FONT, LATIN_FONT, TEAL, TEAL_D
    TITLE, SUBTITLE, ORG = str(cfg["title"]), str(cfg["subtitle"] or ""), str(cfg["org"] or "")
    AUTHOR, VERSION_DATE, SYSTEM_URL = str(cfg["author"] or ""), str(cfg["versionDate"] or ""), str(cfg["url"] or "")
    EA_FONT, LATIN_FONT = str(cfg["bodyFontEastAsia"]), str(cfg["bodyFontLatin"])
    TEAL = cfg["accentColor"]
    TEAL_D = darker(TEAL)


def _el(tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


def make_rfonts():
    """w:rFonts for docDefaults / Normal only -- runs never carry fonts."""
    return _el("w:rFonts", **{"w:ascii": LATIN_FONT, "w:hAnsi": LATIN_FONT,
                              "w:eastAsia": EA_FONT, "w:cs": LATIN_FONT})


def make_ppr_base(indent=False):
    """hanging punctuation off (overflowPunct=0, Word's 「允許標點符號溢出邊界」) +
    snap-to-grid off + spacing 0 / 0 / single + [first-line indent of 2 characters, only
    for the Normal style] + justified (left-right aligned), children in schema order
    (overflowPunct < snapToGrid < spacing < ind < jc).
    Justification never stretches a paragraph's last / only line, so headings,
    captions, cover lines, TOC lines and image-only paragraphs keep their look."""
    els = [
        _el("w:overflowPunct", **{"w:val": "0"}),
        _el("w:snapToGrid", **{"w:val": "0"}),
        _el("w:spacing", **{"w:before": str(SPACE_TWIPS), "w:after": str(SPACE_TWIPS), "w:line": str(LINE_TWIPS), "w:lineRule": "auto"}),
    ]
    if indent:
        # firstLineChars is in 1/100 character; firstLine (twips = 1/20 pt) is the fallback
        # for consumers that ignore the character unit: 2 x body font size.
        els.append(_el("w:ind", **{"w:firstLineChars": str(INDENT_CHARS * 100),
                                   "w:firstLine": str(INDENT_CHARS * BODY_SIZE * 20)}))
    els.append(_el("w:jc", **{"w:val": ALIGN}))
    return els


def no_indent(paragraph):
    """Explicit override of the Normal style's first-line indent (headings, captions,
    images, cover / TOC lines, blank and page-break paragraphs, footer)."""
    ind = paragraph._p.get_or_add_pPr().get_or_add_ind()
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:firstLineChars"), "0")
    return paragraph


def hanging_indent(paragraph):
    """FAQ paragraph (「問：」 / 「答：」): hanging indent (凸排), so wrapped lines align after the
    2-character label.  w:ind leftChars=200 hangingChars=200, no firstLine (this overrides the
    first-line indent inherited from Normal).  In Word's character-unit model the label starts
    at the usual 2-character indent and wrapped lines at 4 characters, so the twips fallback is
    what Word itself writes for this setting: left = 4 x body size (880), hanging = 2 x (440)."""
    ind = paragraph._p.get_or_add_pPr().get_or_add_ind()
    unit = BODY_SIZE * 20  # one body-size character in twips (220 at 11 pt)
    ind.set(qn("w:leftChars"), str(INDENT_CHARS * 100))
    ind.set(qn("w:left"), str(2 * INDENT_CHARS * unit))
    ind.set(qn("w:hangingChars"), str(INDENT_CHARS * 100))
    ind.set(qn("w:hanging"), str(INDENT_CHARS * unit))
    return paragraph


def plain_paragraph(doc):
    """A new paragraph that must not be indented."""
    return no_indent(doc.add_paragraph())


def set_run(run, size=None, bold=None, rgb=None):
    """Run-level size / bold / colour; the font is inherited from Normal."""
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if rgb:
        run.font.color.rgb = RGBColor.from_string(rgb)


def add_text(paragraph, text, size=None, bold=None, rgb=None):
    size = BODY_SIZE if size is None else size
    r = paragraph.add_run(text)
    set_run(r, size, bold, rgb)
    return r


def add_page_break(paragraph):
    """Explicit page-break character as the first run of the paragraph."""
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def add_field(paragraph, instr, size=None):
    run = paragraph.add_run()
    set_run(run, size)
    for tag, extra in (("begin", None), ("instr", instr), ("separate", None), ("end", None)):
        if tag == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = extra
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        run._r.append(el)
    return run


def add_inline(paragraph, text, size=None):
    size = BODY_SIZE if size is None else size
    """Plain text; `code` spans lose their backticks and are otherwise plain."""
    for part in re.split(r"`([^`]*)`", text):
        if part:
            add_text(paragraph, part, size)


def setup_styles(doc):
    """Only the built-in Normal style is used.  Font, paragraph spacing,
    snap-to-grid and justification (jc=both) are set on BOTH docDefaults and Normal; paragraphs and runs
    inherit them (this also overrides the python-docx template's
    after=200 / line=276 defaults).  The 2-character first-line indent is on Normal only."""
    dd = doc.styles.element.find(qn("w:docDefaults"))

    # docDefaults / rPrDefault: fonts (theme refs dropped), zh-TW East Asian language
    rpr_default = dd.find(qn("w:rPrDefault")).find(qn("w:rPr"))
    old = rpr_default.find(qn("w:rFonts"))
    if old is not None:
        rpr_default.remove(old)
    rpr_default.insert(0, make_rfonts())
    lang = rpr_default.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr_default.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-TW")

    # docDefaults / pPrDefault: no hanging punctuation, snap-to-grid off, spacing 0 / 0 / single, justified
    ppr_default = dd.find(qn("w:pPrDefault")).find(qn("w:pPr"))
    for child in list(ppr_default):
        ppr_default.remove(child)
    for el in make_ppr_base():
        ppr_default.append(el)

    # Normal style: same font + paragraph settings
    normal = doc.styles["Normal"]
    rpr = normal.element.get_or_add_rPr()
    rpr.insert(0, make_rfonts())
    normal.font.size = Pt(BODY_SIZE)
    ppr = normal.element.get_or_add_pPr()
    for child in list(ppr):
        ppr.remove(child)
    for el in make_ppr_base(indent=True):
        ppr.append(el)

    # a justified line ending in a manual line break (<br> in a table cell) is not stretched
    compat = doc.settings.element.find(qn("w:compat"))
    if compat is not None and compat.find(qn("w:doNotExpandShiftReturn")) is None:
        compat.insert(0, OxmlElement("w:doNotExpandShiftReturn"))  # before w:useFELayout (schema order)

    # theme-font language hint: Traditional Chinese rather than the template's ja-JP
    tfl = doc.settings.element.find(qn("w:themeFontLang"))
    if tfl is not None:
        tfl.set(qn("w:eastAsia"), "zh-TW")


def last_is_blank(doc):
    """True if the last body block is an empty paragraph (no text, break or drawing)."""
    blocks = [c for c in doc.element.body if c.tag != qn("w:sectPr")]
    if not blocks or blocks[-1].tag != qn("w:p"):
        return False
    last = blocks[-1]
    return not "".join(last.itertext()).strip() and not last.findall(".//" + qn("w:br"))         and not last.findall(".//" + qn("w:drawing"))


def add_heading(doc, text, level, toc):
    """Chapter (level 1) / section (2) / subsection (3) title: plain Normal paragraph, bold
    larger run.  Chapters do not start a new page: one empty Normal paragraph precedes the
    heading as spacing (skipped if the text before it already ends in an empty paragraph).
    The first chapter instead carries the TOC -> body page break."""
    if level == 1 and toc and not last_is_blank(doc):
        plain_paragraph(doc)
    p = plain_paragraph(doc)
    if level == 1:
        if not toc:
            add_page_break(p)  # the body starts on its own page, after the TOC
        toc.append(text)
    p.paragraph_format.keep_with_next = True  # a heading never sits alone at a page bottom
    if level == 1:
        add_text(p, text, H1_SIZE, True, TEAL)
        add_text(plain_paragraph(doc), RULE, RULE_SIZE, False, GREY)
    elif level == 2 and text in BOX_TITLES:
        add_text(p, "── ", H3_SIZE, False, GREY)
        add_text(p, text, H3_SIZE, True, TEAL)
        add_text(p, " " + "─" * BOX_RULE_N, H3_SIZE, False, GREY)
    elif level == 2:
        add_text(p, "▍", H2_SIZE, True, TEAL)
        add_text(p, text, H2_SIZE, True, TEAL_D)
    else:
        add_text(p, "■ ", 9, False, TEAL)
        add_text(p, text, H3_SIZE, True, TEAL_D)
    return p


def add_caption(doc, text, keep_next=False):
    p = plain_paragraph(doc)
    if CAPTION_RE.match(text):
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if keep_next:
        p.paragraph_format.keep_with_next = True  # caption stays with its photo-credit line
    m = CAPTION_RE.match(text)
    if m:
        add_text(p, m.group(0), CAPTION_SIZE, True, TEAL)
        text = text[m.end():]
    add_text(p, text.replace("`", ""), CAPTION_SIZE, True, INK)
    return p


def add_hyperlink(doc, paragraph, url, size, rgb=None):
    rgb = CITE_RGB if rgb is None else rgb
    """External w:hyperlink; its run carries direct formatting only (size + grey colour +
    underline): no rStyle / Hyperlink character style."""
    rid = doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    h.set(qn("w:history"), "1")
    r = OxmlElement("w:r")
    h.append(r)
    paragraph._p.append(h)
    run = Run(r, paragraph)
    run.text = url
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(rgb)
    run.font.underline = True


def add_citation(doc, text):
    """Photo-credit line under a figure caption: centred, 4 pt, grey, not indented, URLs as hyperlinks."""
    p = plain_paragraph(doc)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_together = True
    body = PHOTO_RE.sub("照片來源：", text, count=1)
    pos = 0
    for m in URL_RE.finditer(body):
        if m.start() > pos:
            add_text(p, body[pos:m.start()], CITE_SIZE, rgb=CITE_RGB)
        add_hyperlink(doc, p, m.group(0), CITE_SIZE)
        pos = m.end()
    if pos < len(body):
        add_text(p, body[pos:], CITE_SIZE, rgb=CITE_RGB)
    # the paragraph mark counts towards the last line's height: keep it 4 pt too
    mark = OxmlElement("w:rPr")
    for tag in ("w:sz", "w:szCs"):
        mark.append(_el(tag, **{"w:val": str(CITE_SIZE * 2)}))
    p._p.get_or_add_pPr().append(mark)
    return p


def add_picture_outline(inline):
    """1 pt grey outline on the picture itself (pic:spPr/a:ln), not a paragraph border.
    The line is centred on the picture edge, so half of it (0.5 pt) lies outside the
    picture box: wp:effectExtent reserves that, keeping the outline inside the margins."""
    sppr = inline.xpath(".//pic:spPr")[0]
    for old in sppr.findall(qn("a:ln")):
        sppr.remove(old)
    ln = OxmlElement("a:ln")
    ln.set("w", str(OUTLINE_W))
    fill = OxmlElement("a:solidFill")
    clr = OxmlElement("a:srgbClr")
    clr.set("val", OUTLINE_RGB)
    fill.append(clr)
    ln.append(fill)
    sppr.append(ln)  # after xfrm / prstGeom, as the schema requires
    ee = inline.find(qn("wp:effectExtent"))
    if ee is None:
        ee = OxmlElement("wp:effectExtent")
        inline.find(qn("wp:extent")).addnext(ee)
    half = str(OUTLINE_W // 2)
    for side in "ltrb":
        ee.set(side, half)


def add_image(doc, path):
    with Image.open(path) as im:
        w, h = im.size
    # picture box + the outline's outer half (0.5 pt per side) must fit text width / max height
    avail_w, avail_h = TEXT_W - OUTLINE_W, MAX_IMG_H - OUTLINE_W
    width = avail_w
    if width * h / w > avail_h:
        width = int(avail_h * w / h)
    p = plain_paragraph(doc)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True  # keep the picture on the same page as its caption
    shape = p.add_run().add_picture(str(path), width=width)
    add_picture_outline(shape._inline)


def split_row(line):
    """Cells of one markdown table row (outer pipes optional, `\\|` = literal pipe)."""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    return [c.replace("\\|", "|").strip() for c in re.split(r"(?<!\\)\|", s)]


def parse_blocks(text):
    """Non-blank lines of a chapter file; each markdown table (header row + separator row +
    body rows) becomes ONE item: a list of rows (lists of cell strings)."""
    raw = text.splitlines()
    items, i = [], 0
    while i < len(raw):
        s = raw[i].strip()
        if not s:
            i += 1
        elif s.startswith("|") and i + 1 < len(raw) and TABLE_SEP_RE.match(raw[i + 1].strip()):
            rows, i = [split_row(s)], i + 2
            while i < len(raw) and raw[i].strip().startswith("|"):
                rows.append(split_row(raw[i]))
                i += 1
            items.append(rows)
        else:
            items.append(s)
            i += 1
    return items


def char_width(c):
    """Display width of one character in em (Arial for ASCII, full-width for CJK / symbols)."""
    if unicodedata.east_asian_width(c) in "WFA" and ord(c) > 0x7F:
        return 1.0
    if c.isupper():
        return 0.74
    if c.isdigit() or c == "_":
        return 0.58
    if c == " ":
        return 0.3
    return 0.55 if c.isalpha() else 0.62


def text_width(s):
    return sum(char_width(c) for c in s)


def column_widths(rows, total):
    """Column widths (twips, summing to `total`) from the cell contents.  Every column gets at
    least its longest unbreakable chunk (a Latin word or ~3 CJK characters, capped at the
    column's natural width).  If every cell fits on one line the columns are widened
    proportionally to their natural width; otherwise short columns keep their natural width
    (never wrap) and the long ones share the remaining space equally ("water filling")."""
    em, pad = TABLE_SIZE * 20, TABLE_CELL_MAR["left"] + TABLE_CELL_MAR["right"]
    ncol = len(rows[0])
    nat, mn = [0.0] * ncol, [0.0] * ncol
    for r in rows:
        for c, cell in enumerate(r):
            for line in cell.split("<br>"):
                nat[c] = max(nat[c], text_width(line))
                words = re.findall(r"[!-~]+|[^!-~\s]", line)
                mn[c] = max(mn[c], max((text_width(w) for w in words), default=0.0))
    nat = [n * em * 1.04 + pad for n in nat]  # 4 % safety: Word never breaks inside a Latin word
    mn = [min(max(m * 1.04, 3.0) * em + pad, n) for m, n in zip(mn, nat)]
    if sum(mn) >= total:                       # cannot even fit the minimums: share by minimum
        w = [m / sum(mn) * total for m in mn]
    elif sum(nat) <= total:                    # everything fits on one line: widen proportionally
        w = [n + (total - sum(nat)) * n / sum(nat) for n in nat]
    else:                                      # cap the wide columns at a common width t
        lo, hi = 0.0, max(nat)
        for _ in range(60):
            t = (lo + hi) / 2
            if sum(max(m, min(n, t)) for m, n in zip(mn, nat)) > total:
                hi = t
            else:
                lo = t
        w = [max(m, min(n, lo)) for m, n in zip(mn, nat)]
        capped = [i for i, n in enumerate(nat) if n > lo and mn[i] <= lo]
        for i in capped:                       # hand out the rounding slack
            w[i] += (total - sum(w)) / len(capped)
    w = [int(x) for x in w]
    w[-1] += total - sum(w)
    return w


def add_cell_text(paragraph, text, bold):
    """Cell text: `code` loses its backticks, <br> is a line break; run-level size / bold only."""
    for n, line in enumerate(text.split("<br>")):
        if n:
            paragraph.add_run().add_break()
        for part in re.split(r"`([^`]*)`", line):
            if part:
                add_text(paragraph, part, TABLE_SIZE, True if bold else None)


def add_table(doc, rows, where=""):
    """Simple table with direct formatting only (see module docstring, exception (e))."""
    ncol = len(rows[0])
    for r in rows:
        if len(r) > ncol:
            fail(f"{where}：表格列有 {len(r)} 格，表頭只有 {ncol} 格：{r[0][:30]}")
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    total = int(TEXT_W / 635)  # EMU -> twips
    widths = column_widths(rows, total)
    table = doc.add_table(rows=len(rows), cols=ncol)
    table.style = None  # no table style
    tbl = table._tbl
    # tblPr in schema order: tblW, tblInd, tblBorders, tblLayout, tblCellMar (no tblStyle / tblLook / shd).
    # The document is in Word 2010 compatibility mode (compatibilityMode=14), where a table's
    # left border sits one left-cell-margin to the left of tblInd: tblInd = that margin puts
    # the border exactly on the left text margin, so width = text width fits between the margins.
    tpr = tbl.tblPr
    for child in list(tpr):
        tpr.remove(child)
    tpr.append(_el("w:tblW", **{"w:w": str(total), "w:type": "dxa"}))
    tpr.append(_el("w:tblInd", **{"w:w": str(TABLE_CELL_MAR["left"]), "w:type": "dxa"}))
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(f"w:{side}", **{"w:val": "single", "w:sz": str(TABLE_BORDER_SZ),
                                           "w:space": "0", "w:color": TABLE_BORDER_RGB}))
    tpr.append(borders)
    tpr.append(_el("w:tblLayout", **{"w:type": "autofit" if TABLE_AUTOFIT else "fixed"}))
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        mar.append(_el(f"w:{side}", **{"w:w": str(TABLE_CELL_MAR[side]), "w:type": "dxa"}))
    tpr.append(mar)
    for col, w in zip(tbl.tblGrid.findall(qn("w:gridCol")), widths):
        col.set(qn("w:w"), str(w))
    for ri, (row, cells) in enumerate(zip(table.rows, rows)):
        trpr = row._tr.get_or_add_trPr()
        trpr.append(_el("w:cantSplit"))
        if ri == 0:
            trpr.append(_el("w:tblHeader"))  # repeat the header row on every page
        for cell, text, w in zip(row.cells, cells, widths):
            cell.width = Emu(w * 635)
            p = no_indent(cell.paragraphs[0])
            if ri == 0 and len(rows) > 1:
                p.paragraph_format.keep_with_next = True  # header never stranded at a page bottom
            add_cell_text(p, text, bold=(ri == 0))
    plain_paragraph(doc)  # one empty paragraph as space after the table
    return table


def render_chapter(doc, md_path, seen, toc):
    lines = parse_blocks(md_path.read_text(encoding="utf-8"))
    pending = []  # image paths awaiting their caption

    def flush(caption=None):
        nonlocal pending
        for img in pending:
            add_image(doc, img)
        if caption:
            add_caption(doc, caption)
        pending = []

    def text_at(j):
        """Line j as text ('' if out of range or a table)."""
        return lines[j].strip() if 0 <= j < len(lines) and isinstance(lines[j], str) else ""

    for i, item in enumerate(lines):
        if isinstance(item, list):
            flush()
            add_table(doc, item, md_path.name)
            continue
        line = item.strip()
        m = re.match(r"!\[[^\]]*\]\(([^)]+)\)", line)
        if m:
            img = (md_path.parent / m.group(1)).resolve()
            if img in seen:
                fail(f"圖片重複使用 {img}")
            seen.add(img)
            pending.append(img)
            continue
        if (line.startswith("【") and line.endswith("】")) or CAPTION_RE.match(line):
            nxt = text_at(i + 1)
            if pending:
                for img in pending:
                    add_image(doc, img)
                add_caption(doc, line, keep_next=bool(PHOTO_RE.match(nxt)))
                pending = []
            else:
                flush(line)
            continue
        if PHOTO_RE.match(line):
            if not CAPTION_RE.match(text_at(i - 1) if i else ""):
                fail(f"{md_path.name}：照片來源行必須緊接在「圖 N.M.」圖說之後：{line[:40]}")
            add_citation(doc, line)
            continue
        flush()
        if line.startswith("# "):
            add_heading(doc, line[2:], 1, toc)
        elif line.startswith("## "):
            add_heading(doc, line[3:], 2, toc)
        elif line.startswith("### "):
            add_heading(doc, line[4:], 3, toc)
        elif line.startswith(">"):
            # callout: a normal paragraph; only the lead-in 「注意：」 is bold
            p = doc.add_paragraph()
            body = line.lstrip("> ").strip()
            mm = re.match(r"(注意[：:])(.*)", body)
            if mm:
                add_text(p, mm.group(1), BODY_SIZE, True, ALERT)
                body = mm.group(2)
            add_inline(p, body)
        else:
            p = doc.add_paragraph()
            mq = re.match(r"(問|答)：(.*)", line)
            if mq:
                hanging_indent(p)  # FAQ lead-in: hanging indent (凸排)
                add_text(p, mq.group(1) + "：", BODY_SIZE, True, TEAL if mq.group(1) == "問" else MUTED)
                line = mq.group(2)
            add_inline(p, line)
    flush()


def build_body(doc):
    """Chapters. Returns (toc_titles, image_count)."""
    toc, seen = [], set()
    for ch in CHAPTERS:
        render_chapter(doc, ch, seen, toc)
    return toc, len(seen)



def build_docx(path, toc_pages=None):
    """Build the plain DOCX at `path`.

    `toc_pages` maps heading text -> page number; None means the measuring pass
    (every TOC page number is the placeholder 00).
    """
    titles = toc_titles()
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = MARGIN
    # no document grid at all (the template ships <w:docGrid w:linePitch="360"/>)
    grid = sec._sectPr.find(qn("w:docGrid"))
    if grid is not None:
        sec._sectPr.remove(grid)
    setup_styles(doc)
    doc.core_properties.author = AUTHOR
    doc.core_properties.last_modified_by = AUTHOR
    doc.core_properties.title = TITLE

    # cover: text and character rules only (run size / bold / colour; blank paragraphs push it down).
    # Empty config values leave their line empty, so the cover layout never shifts.
    for _ in range(9):
        plain_paragraph(doc)
    add_text(plain_paragraph(doc), RULE, RULE_SIZE, False, TEAL)
    plain_paragraph(doc)
    add_text(plain_paragraph(doc), TITLE, 28, True, TEAL)
    p = plain_paragraph(doc)
    if SUBTITLE:
        add_text(p, SUBTITLE, 14, False, INK)
    plain_paragraph(doc)
    add_text(plain_paragraph(doc), RULE, RULE_SIZE, False, GREY)
    for _ in range(15):
        plain_paragraph(doc)
    for text, size, bold, rgb in ((ORG, 16, True, INK), (AUTHOR, 11, False, MUTED)):
        p = plain_paragraph(doc)
        if text:
            add_text(p, text, size, bold, rgb)
    p = plain_paragraph(doc)
    if VERSION_DATE:
        add_text(p, f"版本日期　{VERSION_DATE}", 11, False, MUTED)
    p = plain_paragraph(doc)
    if SYSTEM_URL:
        add_hyperlink(doc, p, SYSTEM_URL, 10, MUTED)  # system address: 10 pt grey link, underline only

    # static table of contents: 「title<TAB>page」, one paragraph per chapter, each with
    # ONE right-aligned dot-leader tab stop at the right text margin
    p = plain_paragraph(doc)
    add_page_break(p)
    add_text(p, TOC_TITLE, H1_SIZE, True, TEAL)
    add_text(plain_paragraph(doc), RULE, RULE_SIZE, False, GREY)
    tab_pos = Emu(sec.page_width - sec.left_margin - sec.right_margin)
    for t in titles:
        page = "00" if toc_pages is None else str(toc_pages[t])
        p = plain_paragraph(doc)
        p.paragraph_format.tab_stops.add_tab_stop(tab_pos, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        add_text(p, f"{t}\t{page}")

    toc, n_images = build_body(doc)
    assert toc == titles, (toc, titles)

    # footer: page number field only; the cover (first page) has no footer
    sec.different_first_page_header_footer = True
    sec.first_page_footer.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    for ft in (sec.first_page_footer, sec.footer):
        fp = ft.paragraphs[0]
        fp.style = doc.styles["Normal"]
        no_indent(fp)  # the page number stays where it was (no first-line indent)
    sec.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER  # centred page number only
    add_field(sec.footer.paragraphs[0], "PAGE", 9)

    doc.save(path)
    return n_images


_titles_cache = None


def toc_titles():
    """Chapter headings in document order (level-1 headings)."""
    global _titles_cache
    if _titles_cache is None:
        t = []
        for ch in CHAPTERS:
            for raw in ch.read_text(encoding="utf-8").splitlines():
                if raw.startswith("# "):
                    t.append(raw[2:].strip())
        _titles_cache = t
    return _titles_cache


def measure_pages(word_doc, titles):
    """Page number (as printed in the footer) of each heading's paragraph, read from Word."""
    word_doc.Repaginate()
    wanted, pages = set(titles), {}
    for para in word_doc.Paragraphs:
        txt = para.Range.Text.replace("\x0c", "").strip("\r\x07 ")
        if txt in wanted and txt not in pages:
            if para.Range.Information(12):  # wdWithInTable: a cell that merely equals a heading
                continue
            r = para.Range
            r.SetRange(r.End - 1, r.End - 1)
            pages[txt] = int(r.Information(1))  # wdActiveEndAdjustedPageNumber
    missing = wanted - set(pages)
    if missing:
        fail(f"Word 中找不到這些章標題：{missing}")
    return pages


def with_word(fn):
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")  # private instance; never touches the user's Word
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        return fn(word)
    finally:
        word.Quit()


def word_available():
    """(True, '') if pywin32 is importable and Word can be started, else (False, reason)."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "未安裝 pywin32（pip install pywin32）"
    try:
        with_word(lambda word: None)
    except Exception as e:  # COM error: Word not installed / not registered
        return False, f"無法啟動 Microsoft Word（{e}）"
    return True, ""


def preflight():
    """Validate every chapter before anything is written: images exist and are used once."""
    if not CHAPTERS:
        return
    seen = {}
    for ch in CHAPTERS:
        for n, raw in enumerate(ch.read_text(encoding="utf-8").splitlines(), 1):
            m = re.match(r"!\[[^\]]*\]\(([^)]+)\)", raw.strip())
            if not m:
                continue
            img = (ch.parent / m.group(1)).resolve()
            where = f"{ch.name}:{n}"
            if not img.is_file():
                fail(f"{where}：圖片不存在 {img}")
            if img in seen:
                fail(f"{where}：圖片 {img.name} 重複使用（第一次在 {seen[img]}），每張圖只能出現一次")
            seen[img] = where


def locked(f):
    """True if `f` exists and cannot be opened for writing (open in Word / Acrobat ...)."""
    if not f.exists():
        return False
    try:
        open(f, "r+b").close()
    except PermissionError:
        return True
    return False


def main():
    global OUT_DOCX, OUT_PDF, CHAPTERS
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="由章節 Markdown 產生使用說明書 DOCX + PDF")
    ap.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="設定檔（預設 docs/manual/manual.config.json）")
    ap.add_argument("--out", type=Path, help="輸出資料夾（覆寫設定檔 outDir）")
    ap.add_argument("--chapters", type=Path, help="章節資料夾（覆寫設定檔 chaptersDir）")
    ap.add_argument("--style", type=Path, help="樣式檔（排版偏好；預設為設定檔同資料夾的 manual.style.json，沒有就用內建預設）")
    ap.add_argument("--no-pdf", action="store_true", help="不啟動 Word：只產生 DOCX，目錄頁碼為占位符 00")
    args = ap.parse_args()

    cfg, base = load_config(args.config)
    apply_config(cfg)
    apply_style(load_style(args.style if args.style else base / "manual.style.json"))
    chap_dir = args.chapters.resolve() if args.chapters else base / cfg["chaptersDir"]
    if not chap_dir.is_dir():
        fail(f"找不到章節資料夾 {chap_dir}")
    CHAPTERS = sorted(chap_dir.glob("[0-9][0-9]-*.md"))
    if not CHAPTERS:
        fail(f"章節資料夾 {chap_dir} 中沒有 NN-*.md 章節檔（NN 為兩位數，例如 01-開始使用.md）")
    if not toc_titles():
        fail("章節檔中沒有任何「# 」章標題")
    preflight()

    # 一律用絕對路徑：Word 是另一個程序，相對路徑會被解析到它自己的工作目錄而找不到檔案
    out_dir = (args.out.resolve() if args.out else base / cfg["outDir"]).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    OUT_DOCX, OUT_PDF = out_dir / f"{TITLE}.docx", out_dir / f"{TITLE}.pdf"
    busy = [f for f in (OUT_DOCX, OUT_PDF) if locked(f)]
    if busy:
        fail("下列檔案正被其他程式開啟（Word、Acrobat 等）而無法覆寫，請先關閉後再執行；"
             "本次未改動任何檔案：\n  " + "\n  ".join(str(f) for f in busy), code=2)

    titles = toc_titles()
    use_word, reason = (False, "指定了 --no-pdf") if args.no_pdf else word_available()
    if not use_word:
        n_images = build_docx(OUT_DOCX, None)
        print(f"docx: {OUT_DOCX}  ({n_images} 張圖)")
        print(f"注意：未產生 PDF 與目錄頁碼（{reason}）。目錄頁碼目前為占位符 00；"
              f"需在有 Microsoft Word 與 pywin32 的 Windows 上重新執行才會回填頁碼並匯出 PDF。")
        return

    # pass 1: placeholder page numbers -> measure where each chapter really starts
    tmp = Path(tempfile.mkdtemp()) / "manual-pass1.docx"
    build_docx(tmp, None)

    def pass1(word):
        d = word.Documents.Open(str(tmp), ReadOnly=True)
        try:
            return measure_pages(d, titles)
        finally:
            d.Close(False)

    pages = with_word(pass1)
    print("各章起始頁:", pages)

    # pass 2: final document with real page numbers (right tab stop -> same layout)
    n_images = build_docx(OUT_DOCX, pages)
    print(f"docx: {OUT_DOCX}  ({n_images} 張圖)")

    def pass2(word):
        # read-only and never saved: Word would rewrite styles.xml and drop the explicit
        # spacing / snapToGrid defaults, so the DOCX stays exactly as python-docx wrote it.
        d = word.Documents.Open(str(OUT_DOCX), ReadOnly=True)
        try:
            again = measure_pages(d, titles)
            if again != pages:
                fail(f"兩階段之間頁碼變動：{pages} vs {again}")
            # 17 = wdExportFormatPDF; fonts are embedded by Word's PDF exporter
            d.ExportAsFixedFormat(str(OUT_PDF), 17, False, 0, 0, 1, 1, 0, True, True, 1, True, True, False)
            print(f"pdf:  {OUT_PDF}  ({d.ComputeStatistics(2)} 頁)")
        finally:
            d.Close(False)

    with_word(pass2)


if __name__ == "__main__":
    main()
