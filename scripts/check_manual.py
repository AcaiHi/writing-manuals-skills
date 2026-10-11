"""使用說明書章節來源檔（<chaptersDir>/NN-*.md）的格式與一致性檢查。

用途
    在 build_manual.py 之前執行，機械式檢查章節 Markdown：
    - 章檔編號 00、01、02 … 依序連續、無缺號、無重號
    - 每檔恰有一個「# 」章標題；章標題若寫「第 N 章」，N 須與檔號一致
    - 節編號 N.M（##）/ N.M.K（###）：層級與編號位數相符、N 等於章號、依序遞增
    - 簡單表格：表頭列 + |---| 分隔列 + 內容列，各列欄數一致，儲存格內不放圖片
    - 圖片連結存在；圖片正下方緊接「圖 N.M. 名稱」圖說；圖號章號正確、章內連續、不重複；
      圖名 20 字以內；不使用子圖編號；每張圖只用一次；每張圖都在敘述中被引用（圖 N.M、
      圖 N.M、N.K 列舉、圖 N.M 至 N.K 範圍）；引用的圖號必須存在
    - 照片來源行「> 照片來源：名稱，作者，授權，網址；…」緊接在圖說之下、每筆四欄
    - 排版：全形括號（「」與 `code` 內除外）、中文與英數字之間須有空格、中文旁不可用半形 , ; : ! ?、
      舊式圖號（【圖…】、圖 N-M）
    - 用語：設定檔 terms {錯誤用語: 正確用語}，正確用語為 null 表示禁用
    - 選用（設定檔有指定才檢查）：
      figurePlan   圖表計畫檔：圖號與計畫一致、計畫「放入」的圖都有用到、「只拍不放」的圖未被使用、
                   用到的圖都在計畫中
      photoCredits 照片來源清單：每筆來源的作者 / 授權 / 網址與清單一致、只標註 CC BY、
                   清單中每張 CC BY 照片至少被標註一次
      annotations  標註框定義檔（glob，例如 ../ui-inventory/states/*.boxes.json）：圖上的 (a)(b)…
                   標籤都要在引用該圖的段落中提到

用法（在專案根目錄執行）
    python check_manual.py [--config PATH] [--fix]
      --config PATH  設定檔，預設 docs/manual/manual.config.json
      --fix          同時套用安全的機械修正：圖片連結改為 figuresDir 的相對路徑、
                     全形括號改為半形（並補前後空格）。會直接改寫章節檔。
    設定檔內的相對路徑皆相對於設定檔所在資料夾。

需要的套件
    只用 Python 標準函式庫（3.8 以上）。

退出碼
    0 沒有問題；1 有問題（逐項列出）或設定檔 / 章節資料夾錯誤。

常數意義
    CJK            中文字元範圍（U+4E00 至 U+9FFF）
    IMG_RE         圖片行 ![](路徑)
    CAP_RE         圖說行「圖 N.M.名稱」（名稱前須有一個空格）
    PHOTO_RE       照片來源行「> 照片來源：…」
    REF_RE         敘述中的圖引用「圖 N.M」，可帶 (a) 等標註標籤
    REF_TAIL_RE    引用後的列舉 / 範圍：「、N.K」「至 N.K」「至圖 N.K」
    TABLE_SEP_RE   表格分隔列 |---|:--:|
    HEADING_RE     標題行（# / ## / ###）
    SECNUM_RE      標題開頭的節編號 N.M 或 N.M.K
    SUBFIG_RE      子圖編號「圖 N.M (a)」（不允許）
    OLD_FMT_RE     舊式圖號「【圖」「圖 N-M」（不允許）
    CAPTION_MAX    圖名最長字數 20
    PLAN_KINDS     圖表計畫檔第 6 欄可用值：「放入」「只拍不放」
    DEFAULTS       設定檔中本腳本會用到的鍵與預設值
"""
import argparse
import json
import os
import re
import sys
from pathlib import Path

DEFAULT_CONFIG = Path("docs") / "manual" / "manual.config.json"
DEFAULTS = {"chaptersDir": "chapters", "figuresDir": "figures", "terms": {},
            "figurePlan": "", "photoCredits": "", "annotations": ""}

CJK = r"一-鿿"
IMG_RE = re.compile(r"^!\[[^\]]*\]\(([^)]+)\)\s*$")
CAP_RE = re.compile(r"^圖 (\d+)\.(\d+)\.(.*)$")
PHOTO_RE = re.compile(r"^> *照片來源：(.*)$")  # photo-credit line: > 照片來源：名稱，作者，授權，網址；…
REF_RE = re.compile(r"圖 ?(\d+)\.(\d+)((?:\([a-z]\))*)")
# additional numbers in a list/range after a reference: 圖 4.7、4.8 / 圖 4.7 至 4.9
REF_TAIL_RE = re.compile(r"(?:、|\s*至\s*(?:圖 ?)?)(\d+)\.(\d+)(?!\d)")
TABLE_SEP_RE = re.compile(r"^(?=.*\|)\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$")  # |---|:--:| separator row
HEADING_RE = re.compile(r"^(#{1,3}) (.*)$")
SECNUM_RE = re.compile(r"^(\d+)\.(\d+)(?:\.(\d+))?(?=\s|$)")  # heading number 2.1 / 10.1.3
SUBFIG_RE = re.compile(r"圖 ?\d+\.\d+\s*\([a-z]\)")
OLD_FMT_RE = re.compile(r"【圖|圖 ?\d+-\d+")  # retired caption/reference format
CAPTION_MAX = 20
PLAN_KINDS = ("放入", "只拍不放")

issues = []


def issue(kind, where, msg):
    issues.append((kind, where, msg))


def fail(msg):
    print(f"錯誤：{msg}", file=sys.stderr)
    sys.exit(1)


def load_config(path):
    path = Path(path)
    if not path.is_file():
        fail(f"找不到設定檔 {path}（請在專案根目錄執行，預設設定檔是 docs/manual/manual.config.json；沒有的話由 templates/manual.config.json 複製一份，或以 --config 指定）")
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        fail(f"設定檔 {path} 不是有效的 JSON：{e}")
    cfg = dict(DEFAULTS)
    cfg.update({k: v for k, v in data.items() if v is not None})
    if not isinstance(cfg["terms"], dict):
        fail("設定檔 terms 必須是物件 {錯誤用語: 正確用語或 null}")
    return cfg, path.resolve().parent


def split_row(line):
    r"""Cells of one markdown table row (outer pipes optional, `\|` = literal pipe)."""
    t = line.strip()
    if t.startswith("|"):
        t = t[1:]
    if t.endswith("|") and not t.endswith("\\|"):
        t = t[:-1]
    return [c.replace("\\|", "|").strip() for c in re.split(r"(?<!\\)\|", t)]


def check_tables(name, lines):
    """Validate GitHub-style tables; returns the number of tables found."""
    n, i = 0, 0
    while i < len(lines):
        if not lines[i].strip().startswith("|"):
            i += 1
            continue
        where = f"{name}:{i + 1}"
        if i + 1 >= len(lines) or not TABLE_SEP_RE.match(lines[i + 1].strip()):
            issue("table", where, "以 | 開頭的行不屬於任何表格（需要表頭列 + |---| 分隔列）")
            i += 1
            continue
        header = split_row(lines[i])
        ncol = len(header)
        if any(not c for c in header):
            issue("table", where, "表頭有空白儲存格")
        if len(split_row(lines[i + 1])) != ncol:
            issue("table", f"{name}:{i + 2}", f"分隔列有 {len(split_row(lines[i + 1]))} 格，表頭有 {ncol} 格")
        i += 2
        body = 0
        while i < len(lines) and lines[i].strip().startswith("|"):
            w = f"{name}:{i + 1}"
            cells = split_row(lines[i])
            if TABLE_SEP_RE.match(lines[i].strip()):
                issue("table", w, "表格中出現第二個分隔列")
            else:
                body += 1
                if len(cells) != ncol:
                    issue("table", w, f"此列有 {len(cells)} 格，表頭有 {ncol} 格（儲存格內的 | 要寫成 \\|）")
                if any("![" in c for c in cells):
                    issue("table", w, "表格儲存格內不可放圖片")
            i += 1
        if not body:
            issue("table", where, "表格沒有內容列")
        n += 1
    return n


def check_chapter_files(files):
    """File numbering 00, 01, 02 ... without gaps or duplicates."""
    nums = [int(f.name[:2]) for f in files]
    for k in sorted({x for x in nums if nums.count(x) > 1}):
        issue("order", "chapters", f"章檔編號 {k:02d} 被多個檔案使用")
    missing = [f"{k:02d}" for k in range(max(nums) + 1) if k not in nums] if nums else []
    if missing:
        issue("order", "chapters", f"章檔編號缺號：{', '.join(missing)}")


def plan(path):
    """Figure plan: markdown table rows with 8 cells whose 6th cell is 放入 / 只拍不放 ->
    {figure id (4th cell): (圖號 (2nd cell), 子項 (3rd cell), kind)}."""
    rows = {}
    if not path:
        return rows
    for line in path.read_text(encoding="utf-8").splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) == 8 and cells[5] in PLAN_KINDS:
            rows[cells[3]] = (cells[1], cells[2], cells[5])
    return rows


def credits(path):
    """Photo credit list: markdown table | File | URL | Author | License | ->
    {display name (file stem before '_'): (author, license, url)}."""
    rows = {}
    if not path:
        return rows
    for line in path.read_text(encoding="utf-8").splitlines():
        c = [x.strip() for x in line.strip().strip("|").split("|")]
        if len(c) == 4 and c[0] != "File" and not set(c[0]) <= set("-"):
            rows[Path(c[0]).stem.split("_")[0]] = (c[2], c[3], c[1])
    return rows


def boxes(pattern):
    """{figure id: {letter: {"label": ...}}} from the annotation JSON files."""
    out = {}
    if not pattern:
        return out
    for f in sorted(pattern.parent.glob(pattern.name)):
        out.update(json.loads(f.read_text(encoding="utf-8")))
    return out


def outside_quotes(text):
    """Yield (index, char, in_quote) skipping 「」 and `code` regions."""
    depth, code = 0, False
    for i, c in enumerate(text):
        if c == "`":
            code = not code
        elif not code and c == "「":
            depth += 1
        elif not code and c == "」":
            depth = max(0, depth - 1)
        yield i, c, depth > 0 or code


def fix_parens(line):
    """Full-width parens outside 「」 / `code` -> half-width, with a space on the outer side
    when the neighbour is CJK / alnum / a quote bracket."""
    if line.startswith("#") or IMG_RE.match(line):
        return line
    chars = list(line)
    out = []
    flags = {i: q for i, _, q in outside_quotes(line)}
    for i, c in enumerate(chars):
        if c in "（）" and not flags[i]:
            if c == "（":
                prev = out[-1] if out else ""
                out.append((" " if re.match(f"[{CJK}A-Za-z0-9」]", prev or "x") and prev else "") + "(")
            else:
                nxt = chars[i + 1] if i + 1 < len(chars) else ""
                out.append(")" + (" " if re.match(f"[{CJK}A-Za-z0-9「]", nxt or "x") and nxt else ""))
        else:
            out.append(c)
    return "".join(out)


def optional_path(base, value, key):
    """Resolve an optional config path; '' -> None; a missing file is an error."""
    if not value:
        return None
    p = (base / value).resolve() if not Path(value).is_absolute() else Path(value)
    if key != "annotations" and not p.is_file():
        fail(f"設定檔 {key} 指向的檔案不存在：{p}")
    return p


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="檢查使用說明書章節檔的格式與一致性")
    ap.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="設定檔（預設 docs/manual/manual.config.json）")
    ap.add_argument("--fix", action="store_true", help="套用安全的機械修正（會改寫章節檔）")
    args = ap.parse_args()
    fix = args.fix

    cfg, base = load_config(args.config)
    ch_dir = (base / cfg["chaptersDir"]).resolve()
    fig_dir = (base / cfg["figuresDir"]).resolve()
    if not ch_dir.is_dir():
        fail(f"找不到章節資料夾 {ch_dir}")
    fig_rel = Path(os.path.relpath(fig_dir, ch_dir)).as_posix()  # e.g. ../figures
    plan_path = optional_path(base, cfg["figurePlan"], "figurePlan")
    cred_path = optional_path(base, cfg["photoCredits"], "photoCredits")
    box_pat = optional_path(base, cfg["annotations"], "annotations")
    terms = cfg["terms"]

    p = plan(plan_path)
    bx = boxes(box_pat)
    cr = credits(cred_path)
    cited = {}     # photo name -> [figure keys that credit it]
    last_fig = None
    used = {}
    fig_defs = {}  # "N.M" -> {"subs": [ids], "caption": str, "chapter": n}
    refs = []      # (chapter, para_text, N.M, subs)
    files = sorted(ch_dir.glob("[0-9][0-9]-*.md"))
    if not files:
        fail(f"章節資料夾 {ch_dir} 中沒有 NN-*.md 章節檔")
    check_chapter_files(files)
    n_tables = n_h1 = 0
    for path in files:
        chap = int(path.name[:2])
        lines = path.read_text(encoding="utf-8").splitlines()
        n_tables += check_tables(path.name, lines)
        h1 = [(ln, m.group(2)) for ln, l in enumerate(lines, 1) if (m := HEADING_RE.match(l)) and m.group(1) == "#"]
        n_h1 += len(h1)
        if len(h1) != 1:
            issue("heading", path.name, f"每個章節檔應恰有一個「# 」章標題，實際 {len(h1)} 個")
        elif (cm := re.match(r"第 (\d+) 章", h1[0][1])) and int(cm.group(1)) != chap:
            issue("heading", f"{path.name}:{h1[0][0]}", f"標題寫第 {cm.group(1)} 章，但檔號是第 {chap} 章")
        sec = sub = 0  # running section / subsection numbers
        new = []
        pending = None   # (fid, where) of the last image awaiting its caption
        expect = 0       # running figure count in this chapter
        prev_nonblank = None
        for ln, line in enumerate(lines, 1):
            where = f"{path.name}:{ln}"
            hm = HEADING_RE.match(line)
            if hm and len(hm.group(1)) >= 2 and (nm := SECNUM_RE.match(hm.group(2))):
                a, b, c = int(nm.group(1)), int(nm.group(2)), nm.group(3)
                if (len(hm.group(1)) == 2) != (c is None):
                    issue("heading", where, f"標題層級 {len(hm.group(1))} 與編號 {nm.group(0)} 不符（## 用 N.M，### 用 N.M.K）")
                elif a != chap:
                    issue("heading", where, f"節編號 {nm.group(0)} 出現在第 {chap} 章")
                elif c is None:
                    if b != sec + 1:
                        issue("heading", where, f"節 {nm.group(0)}：應為 {chap}.{sec + 1}")
                    sec, sub = b, 0
                else:
                    if b != sec:
                        issue("heading", where, f"小節 {nm.group(0)} 不在第 {chap}.{sec} 節之下")
                    elif int(c) != sub + 1:
                        issue("heading", where, f"小節 {nm.group(0)}：應為 {chap}.{b}.{sub + 1}")
                    sub = int(c)
            if OLD_FMT_RE.search(line):
                issue("format", where, f"舊式圖號（請用 圖 N.M / 圖 N.M. 名稱）：{OLD_FMT_RE.search(line).group(0)}")
            if SUBFIG_RE.search(line):
                issue("subfig", where, f"不可使用子圖編號：{SUBFIG_RE.search(line).group(0)}")
            m = IMG_RE.match(line)
            if m:
                link = m.group(1)
                fid = Path(link).stem
                want = f"{fig_rel}/{fid}.png"
                if (fig_dir / f"{fid}.png").exists() and link != want:
                    issue("link", where, f"{fid}：{link} -> {want}")
                    if fix:
                        line = f"![]({want})"; link = want
                if not (ch_dir / link).resolve().exists():
                    issue("link", where, f"圖片檔不存在 {link}")
                used.setdefault(fid, []).append(where)
                if pending:
                    issue("caption", pending[1], f"圖片 {pending[0]} 之後沒有自己的圖說")
                pending = (fid, where)
            c = CAP_RE.match(line)
            if c:
                n, k, rest = int(c.group(1)), int(c.group(2)), c.group(3)
                key = f"{n}.{k}"
                if n != chap:
                    issue("number", where, f"圖說 圖 {key} 出現在第 {chap} 章")
                if key in fig_defs:
                    issue("number", where, f"圖說 圖 {key} 重複")
                if not pending or not (prev_nonblank is not None and IMG_RE.match(prev_nonblank)):
                    issue("caption", where, f"圖 {key}：圖說不在圖片正下方")
                expect += 1
                if k != expect:
                    issue("number", where, f"圖 {key}：依順序應為 圖 {chap}.{expect}")
                fid = pending[0] if pending else None
                last_fig = key
                fig_defs[key] = {"subs": [fid] if fid else [], "caption": rest, "chapter": chap, "where": where}
                name = rest.strip()
                if not rest.startswith(" ") or not name:
                    issue("caption", where, f"圖 {key}：圖說缺少名稱（格式「圖 N.M. 名稱」）")
                if len(name) > CAPTION_MAX:
                    issue("caption", where, f"圖 {key}：圖名超過 {CAPTION_MAX} 字（{len(name)}）")
                if re.search(r"\([a-z]\)", name):
                    issue("subfig", where, f"圖 {key}：圖說列出子圖")
                if fid:
                    pl = p.get(fid)
                    if pl and pl[0] != f"圖 {key}":
                        issue("plan", where, f"{fid}：計畫為 {pl[0]}，手冊為 {key}")
                pending = None
            elif pending and line.strip() and not m:
                issue("caption", where, f"圖片 {pending[0]} 之後沒有圖說")
                pending = None
            ph = PHOTO_RE.match(line)
            if ph:
                if not (prev_nonblank is not None and CAP_RE.match(prev_nonblank)):
                    issue("credit", where, "照片來源行不在圖說正下方")
                seen_names = set()
                for entry in ph.group(1).split("；"):
                    parts = [x.strip() for x in entry.split("，")]
                    if len(parts) != 4:
                        issue("credit", where, f"每筆需為「名稱，作者，授權，網址」：{entry[:40]}")
                        continue
                    name, author, lic, url = parts
                    if cred_path:
                        row = cr.get(name)
                        if not row:
                            issue("credit", where, f"「{name}」不在照片來源清單中")
                        elif (author, lic, url) != row:
                            issue("credit", where, f"{name}：作者 / 授權 / 網址與清單不同 {row}")
                        elif not lic.startswith("CC BY"):
                            issue("credit", where, f"{name}：{lic} 不需標註來源，請刪除")
                    if not url.startswith(("http://", "https://")):
                        issue("credit", where, f"{name}：網址須以 http:// 或 https:// 開頭")
                    if name in seen_names:
                        issue("credit", where, f"{name} 在同一行標註兩次")
                    seen_names.add(name)
                    cited.setdefault(name, []).append(last_fig)
            if line.strip():
                prev_nonblank = line
            if not line.startswith("#") and not m and not c and not ph and line.strip():
                for r in REF_RE.finditer(line):
                    refs.append((chap, line, f"{r.group(1)}.{r.group(2)}", [], where))
                    pos = r.end()
                    while (tm := REF_TAIL_RE.match(line, pos)):
                        a, b = int(r.group(2)), int(tm.group(2))
                        if "至" in tm.group(0) and tm.group(1) == r.group(1):
                            for x in range(a + 1, b + 1):
                                refs.append((chap, line, f"{r.group(1)}.{x}", [], where))
                        else:
                            refs.append((chap, line, f"{tm.group(1)}.{tm.group(2)}", [], where))
                        pos = tm.end()
                # typography
                for i, ch, q in outside_quotes(line):
                    if q:
                        continue
                    if ch in "（）" and not fix:
                        issue("typo", where, f"全形括號（第 {i} 字）：…{line[max(0, i - 8):i + 8]}…")
                        break
                for bad, good in terms.items():
                    if bad in line:
                        issue("term", where, f"「{bad}」" + (f" -> 「{good}」" if good else "（不可使用）"))
                stripped = re.sub(r"「[^」]*」|`[^`]*`", "「」", line)
                for mm in re.finditer(f"[{CJK}][A-Za-z0-9]|[A-Za-z0-9][{CJK}]", stripped):
                    issue("spacing", where, f"中文與英數字相鄰未空格「{mm.group(0)}」")
                for mm in re.finditer(f"[{CJK}][,;:!?]|[,;:!?][{CJK}]", stripped):
                    issue("typo", where, f"中文旁使用半形標點「{mm.group(0)}」")
            if fix:
                line = fix_parens(line)
            new.append(line)
        if fix:
            path.write_text("\n".join(new) + "\n", encoding="utf-8")

    # every CC BY photo is credited at least once (CC0 / public domain need no credit)
    if cred_path:
        for name, (author, lic, url) in cr.items():
            if lic.startswith("CC BY") and name not in cited:
                issue("credit", cred_path.name, f"{name}（{lic}）從未在手冊中標註")
    # numbering continuity
    by_ch = {}
    for key, d in fig_defs.items():
        by_ch.setdefault(d["chapter"], []).append(int(key.split(".")[1]))
    for ch, nums in by_ch.items():
        if sorted(nums) != list(range(1, len(nums) + 1)):
            issue("number", f"ch{ch}", f"圖號不連續：{sorted(nums)}")
    # usage
    for fid, ws in used.items():
        if len(ws) > 1:
            issue("usage", ", ".join(ws), f"{fid} 使用了 {len(ws)} 次（每張圖只能用一次）")
        if plan_path:
            if fid in p and p[fid][2] == "只拍不放":
                issue("usage", ws[0], f"{fid} 在計畫中為「只拍不放」卻被使用")
            if fid not in p:
                issue("usage", ws[0], f"{fid} 不在圖表計畫中")
    if plan_path:
        for fid, (num, sub, kind) in p.items():
            if kind == "放入" and fid not in used:
                issue("usage", plan_path.name, f"{fid}（{num}{sub}）計畫放入但未使用")
    # references
    refd = set()
    for chap, para, key, subs, where in refs:
        if key not in fig_defs:
            issue("ref", where, f"引用了不存在的 圖 {key}")
            continue
        refd.add(key)
    for key, d in fig_defs.items():
        if key not in refd:
            issue("ref", d["where"], f"圖 {key} 從未在敘述中被引用")
    # annotation labels mentioned near the reference (same paragraph)
    for key, d in fig_defs.items():
        for fid in d["subs"]:
            labels = bx.get(fid)
            if not labels:
                continue
            paras = [para for _, para, k, _, _ in refs if k == key]
            text = REF_RE.sub("", "".join(paras))
            for letter, info in labels.items():
                if f"({letter})" not in text:
                    issue("box", d["where"], f"圖 {key} [{fid}] 標註 ({letter}) {info.get('label', '')} 未在引用段落中提到")

    kinds = {}
    for k, w, msg in issues:
        kinds.setdefault(k, []).append((w, msg))
    for k, items in kinds.items():
        print(f"## {k} ({len(items)})")
        for w, msg in items:
            print(f"- {w}: {msg}")
    n_cite = sum(len(v) for v in cited.values())
    skipped = [k for k, v in (("figurePlan", plan_path), ("photoCredits", cred_path), ("annotations", box_pat)) if not v]
    print()
    print(f"TOTAL {len(issues)} issues; chapters {len(files)}; tables {n_tables}; figures defined {len(fig_defs)}; "
          f"images used {len(used)}; photo credits {n_cite} entries on {len({f for v in cited.values() for f in v})} figures")
    if skipped:
        print(f"（設定檔未指定，已略過相關檢查：{', '.join(skipped)}）")
    sys.exit(1 if issues else 0)


if __name__ == "__main__":
    main()
