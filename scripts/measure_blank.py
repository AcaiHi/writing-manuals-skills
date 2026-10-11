"""量測已產生的說明書 PDF 每一頁的空白量（唯讀，不改動任何檔案）。

用途
    排除封面與目錄頁，逐頁把內文頁轉為灰階影像，在文字區（頁高扣掉上下邊界）內，
    整列（幾乎）全白的像素列視為空白。每頁回報：
      blank %   （頁尾連續空白 + 較大的內部空白）/ 文字區高度
      trail cm  文字區底部的連續空白（下一頁開頭的內容放不下而留下的空間）
      gaps cm   高度 >= --gap-cm 的內部空白（例如圖片被推到下一頁時上方留下的空間）
      chars     空白 cm x 60：大約可再填入多少個全形中文字
    摘要列出空白比例超過門檻（預設 10%）的頁、所屬章與建議補充的字數。
    最後一頁會顯示但不計入（其頁尾空白無法避免）。

    「chars」的估計：在 A4、內文 11 pt、單行行距（每行約 0.67 cm、每行約 40 個全形字）下，
    1 cm 空白約 1.5 行，約 60 個中文字。標題、圖說、圖片每字所需空間較大，此數字僅供
    判斷大約要補多少說明文字。

用法（在專案根目錄執行）
    python measure_blank.py [PDF] [--config PATH] [--threshold PCT] [--gap-cm CM] [--json] [--chapters DIR]
      PDF            預設為 <outDir>/<title>.pdf（由設定檔推得）
      --config PATH  設定檔，預設 docs/manual/manual.config.json；只指定 PDF 與 --chapters 時可不需要
      --threshold    空白比例超過此值（%）的頁會列出，預設 10
      --gap-cm       內部空白至少此高度（cm）才計入，預設 1.2
      --dpi          轉圖解析度，預設 50
      --chapters DIR 章節資料夾（用來把頁對應到章），預設為設定檔 chaptersDir
      --json         以 JSON 輸出

需要的套件
    PyMuPDF（pip install pymupdf）。

退出碼
    0 成功；1 缺少 PyMuPDF、找不到 PDF 或設定檔、PDF 無法開啟。

常數意義
    MARGIN_CM      上下頁邊界 2.2 cm（與 build_manual.py 的 MARGIN 相同）
    CHARS_PER_CM   1 cm 空白約等於 60 個全形字（11 pt、單行行距）
    WHITE          灰階值 >= 235（0 至 255）的像素視為白色
    LEADER_RE      目錄的點狀前導線（連續 5 個以上的點），用來辨識目錄頁
"""
import argparse
import json
import re
import sys
from pathlib import Path

DEFAULT_CONFIG = Path("docs") / "manual" / "manual.config.json"
MARGIN_CM = 2.2          # top / bottom margin of the manual (build_manual.py MARGIN)
CHARS_PER_CM = 60        # 1 cm of blank = about 60 full-width characters (11 pt, single spacing)
WHITE = 235              # a pixel >= this (0-255 grey) counts as white
LEADER_RE = re.compile(r"[.．…]{5,}")  # TOC dot leaders

pymupdf = None


def fail(msg):
    print(f"錯誤：{msg}", file=sys.stderr)
    sys.exit(1)


def chapter_titles(chapters_dir):
    """Level-1 headings of the chapter files (to attribute pages to chapters)."""
    titles = []
    if not chapters_dir or not chapters_dir.is_dir():
        return titles
    for f in sorted(chapters_dir.glob("[0-9][0-9]-*.md")):
        for line in f.read_text(encoding="utf-8").splitlines():
            if line.startswith("# "):
                titles.append(line[2:].strip())
    return titles


def squash(s):
    return re.sub(r"\s+", "", s)


def first_body_page(doc):
    """1-based number of the first page after the cover and the TOC.  The TOC pages are the
    consecutive pages after the cover whose text carries dot-leader lines."""
    page = 2
    while page <= len(doc):
        leaders = sum(1 for ln in doc[page - 1].get_text().splitlines() if LEADER_RE.search(ln))
        if leaders < 3:
            break
        page += 1
    return page


def blank_runs(rows_blank):
    """[(start, length)] of consecutive True runs."""
    runs, start = [], None
    for i, b in enumerate(rows_blank + [False]):
        if b and start is None:
            start = i
        elif not b and start is not None:
            runs.append((start, i - start))
            start = None
    return runs


def measure_page(page, dpi, gap_cm):
    pt_per_cm = 72 / 2.54
    scale = dpi / 72
    top = int(MARGIN_CM * pt_per_cm * scale)
    bottom = int((page.rect.height - MARGIN_CM * pt_per_cm) * scale)
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, alpha=False)
    data, stride = pix.samples, pix.stride
    rows = [min(data[y * stride:(y + 1) * stride]) >= WHITE for y in range(top, min(bottom, pix.height))]
    cm_per_px = 2.54 / dpi
    height_cm = len(rows) * cm_per_px
    runs = blank_runs(rows)
    trail = next((n for s, n in runs if s + n == len(rows)), 0)
    gaps = sum(n for s, n in runs if s + n != len(rows) and n * cm_per_px >= gap_cm)
    if all(rows):  # completely empty page: everything is "trailing"
        gaps = 0
    trail_cm, gaps_cm = trail * cm_per_px, gaps * cm_per_px
    blank_cm = trail_cm + gaps_cm
    return {"trail_cm": round(trail_cm, 2), "gaps_cm": round(gaps_cm, 2), "blank_cm": round(blank_cm, 2),
            "blank_pct": round(100 * blank_cm / height_cm, 1), "chars": int(blank_cm * CHARS_PER_CM)}


def resolve_inputs(args):
    """PDF path and chapter dir from the arguments, falling back to the config."""
    pdf, chapters = args.pdf, args.chapters
    if pdf and chapters:
        return pdf, chapters
    cfg_path = args.config
    if not cfg_path.is_file():
        if pdf:  # chapters unknown: pages are simply not attributed to chapters
            return pdf, None
        fail(f"找不到設定檔 {cfg_path}；請在專案根目錄執行，或指定 PDF 路徑，或以 --config 指定設定檔")
    try:
        cfg = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        fail(f"設定檔 {cfg_path} 不是有效的 JSON：{e}")
    base = cfg_path.resolve().parent
    if not pdf:
        title = cfg.get("title") or "使用說明書"
        pdf = base / (cfg.get("outDir") or ".") / f"{title}.pdf"
    if not chapters:
        chapters = base / (cfg.get("chaptersDir") or "chapters")
    return pdf, chapters


def main():
    global pymupdf
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="說明書 PDF 每頁空白量報告")
    ap.add_argument("pdf", nargs="?", type=Path, help="PDF 路徑（預設由設定檔推得）")
    ap.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="設定檔（預設 docs/manual/manual.config.json）")
    ap.add_argument("--threshold", type=float, default=10.0, help="空白比例門檻 %%（預設 10）")
    ap.add_argument("--gap-cm", type=float, default=1.2, help="內部空白至少此高度才計入（預設 1.2）")
    ap.add_argument("--dpi", type=int, default=50, help="轉圖解析度（預設 50）")
    ap.add_argument("--chapters", type=Path, help="章節資料夾（預設為設定檔 chaptersDir）")
    ap.add_argument("--json", action="store_true", help="以 JSON 輸出")
    args = ap.parse_args()
    try:
        import pymupdf as _m
    except ImportError:
        try:
            import fitz as _m  # older PyMuPDF releases
        except ImportError:
            fail("缺少 PyMuPDF，請執行：pip install pymupdf")
    pymupdf = _m

    pdf, chapters = resolve_inputs(args)
    if not pdf.is_file():
        fail(f"找不到 PDF {pdf}（請先以 build_manual.py 產生，或直接指定 PDF 路徑）")
    try:
        doc = pymupdf.open(pdf)
    except Exception as e:
        fail(f"無法開啟 PDF {pdf}：{e}")
    titles = {squash(t): t for t in chapter_titles(chapters)}
    start = first_body_page(doc)
    chapter = ""  # chapter in effect at the end of each page: the last heading seen so far
    pages = []
    for n in range(start, len(doc) + 1):
        page = doc[n - 1]
        heads = [titles[squash(ln)] for ln in page.get_text().splitlines() if squash(ln) in titles]
        if heads:
            chapter = heads[-1]
        m = measure_page(page, args.dpi, args.gap_cm)
        m.update(page=n, chapter=chapter, starts=heads, last=(n == len(doc)))
        pages.append(m)

    counted = [m for m in pages if not m["last"]]
    over = [m for m in counted if m["blank_pct"] > args.threshold]
    summary = {
        "pdf": str(pdf), "pages": len(doc), "first_body_page": start, "threshold_pct": args.threshold,
        "gap_cm": args.gap_cm, "chars_per_cm": CHARS_PER_CM,
        "pages_over": [m["page"] for m in over], "n_over": len(over), "n_measured": len(counted),
        "worst": max(counted, key=lambda m: m["blank_pct"])["page"] if counted else None,
        "total_blank_cm": round(sum(m["blank_cm"] for m in counted), 1),
    }
    if args.json:
        print(json.dumps({"summary": summary, "pages": pages}, ensure_ascii=False, indent=2))
        return

    print(f"{pdf.name}: {len(doc)} pages; cover and TOC excluded -> measuring pages {start}-{len(doc)}")
    print(f"(1 cm of blank ~ {CHARS_PER_CM} CJK chars; gaps = internal blank runs >= {args.gap_cm} cm)")
    print(f"{'page':>4} {'blank%':>7} {'trail cm':>9} {'gaps cm':>8} {'chars':>6}  chapter")
    for m in pages:
        flag = "  <-- over" if m in over else ("  (last page, not counted)" if m["last"] else "")
        print(f"{m['page']:>4} {m['blank_pct']:>7.1f} {m['trail_cm']:>9.1f} {m['gaps_cm']:>8.1f} {m['chars']:>6}  "
              f"{m['chapter']}{'  [starts: ' + ', '.join(m['starts']) + ']' if m['starts'] else ''}{flag}")
    print()
    print(f"SUMMARY: {len(over)} of {len(counted)} pages over {args.threshold:g}% blank"
          + (f" -> pages {', '.join(str(m['page']) for m in over)}" if over else ""))
    for m in sorted(over, key=lambda m: -m["blank_pct"])[:15]:
        print(f"  p{m['page']:<3} {m['blank_pct']:>5.1f}%  blank {m['blank_cm']:.1f} cm = about {m['chars']} chars  ({m['chapter']})")
    print(f"total blank on counted pages: {summary['total_blank_cm']} cm")


if __name__ == "__main__":
    main()
