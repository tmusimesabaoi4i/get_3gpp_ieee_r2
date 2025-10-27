# -*- coding: utf-8 -*-
"""
Office(Word/PPT) → HTML 結合器（テキスト抽出・高速版）
- Word: 段落＋（任意）StoryRanges からテキスト抽出
- PPT : 各スライドの TextFrame / Table からテキスト抽出
- 出力: HTML（各ファイルを <h2>、PPTの各スライドを <h3>）
- 画像出力や Word Filtered HTML は行わない
"""

import os, re, time, random, subprocess, shutil, hashlib
from pathlib import Path
from typing import Iterable, List, Dict, Tuple, Optional

import pythoncom
from win32com.client import gencache, DispatchEx
from concurrent.futures import ProcessPoolExecutor, as_completed
import html as _html

# ==================== 調整フラグ ====================
BATCH_SIZE_DEFAULT      = 5          # 1 HTML の最大収録件数（part単位）
MAX_FILE_KB_DEFAULT     = 1500       # 閾値超はスキップ
DEFAULT_RETRIES         = 1          # COM切断時の再試行回数（各ファイル）
EXCLUSIVE_INSTANCE      = True       # True: DispatchEx で専用インスタンス化
DEFER_QUIT_TO_PARENT    = True       # True: ワーカーで Quit しない（最後に親が Kill）
KILL_AT_START           = True       # 実行開始前に既存 Office Kill
KILL_AT_END             = True       # 実行終了後に Office Kill
USE_WORD_STORY_RANGES   = True       # Word: StoryRanges も走査して拾い漏れ低減

# ==================== 調整フラグ（追加/確認） ====================
PPT_USE_TEXTFRAME2_FALLBACK = False   # Trueで TextFrame2 も試す（やや低速）

# ==================== 定数（追加） ====================
msoGroup = 6


# ==================== 定数/拡張子 ====================
WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
PPT_EXTS  = {".ppt", ".pptx", ".pptm"}
ALLOWED_EXTS = WORD_EXTS | PPT_EXTS

msoTrue  = -1
msoFalse = 0
msoAutomationSecurityForceDisable = 3

# ==================== プロセス一括 Kill ====================
def kill_office_processes() -> None:
    for image_name in ("WINWORD.EXE", "POWERPNT.EXE"):
        try:
            subprocess.run(
                ["taskkill", "/IM", image_name, "/F", "/T"],
                capture_output=True, text=True, check=False
            )
        except Exception:
            pass

# ==================== ユーティリティ ====================
_para_splitter = re.compile(r"[\r\n]+")
def _clean_paragraph_text(s: str) -> str:
    if s is None: return ""
    s = (s.replace("\r", "")
           .replace("\x07","")
           .replace("\u200b","")
           .replace("\ufeff",""))
    return s.strip()

def _split_paragraphs_fast(s: str) -> List[str]:
    if not s: return []
    out = []
    for x in _para_splitter.split(s):
        c = _clean_paragraph_text(x)
        if c: out.append(c)
    return out

def _ensure_paragraphs_list(items: Iterable[str]) -> List[str]:
    out: List[str] = []
    for t in items:
        t = _clean_paragraph_text(t)
        if t: out.append(t)
    return out

def _short_hash(s: str) -> str:
    return hashlib.sha1(s.encode("utf-8", errors="ignore")).hexdigest()[:8]

def _normalize_items(path_lst: Iterable[str | Path | Tuple[str, str]]) -> List[Path]:
    out: List[Path] = []
    for item in path_lst:
        if isinstance(item, (tuple, list)):
            p = Path(str(item[0])) / str(item[1]) if len(item) >= 2 else Path(str(item[0]))
        else:
            p = Path(str(item))
        out.append(p)
    return out

def _prefilter_items(paths: List[Path], max_kb: int = MAX_FILE_KB_DEFAULT) -> List[Path]:
    ok: List[Path] = []
    for p in paths:
        if not p.exists(): continue
        ext = p.suffix.lower()
        if ext not in ALLOWED_EXTS: continue
        try:
            size = p.stat().st_size
            if size > max_kb * 1024:
                print(f"[SKIP size>{max_kb}KB] {p} ({size/1024:.1f} KB)")
                continue
        except Exception as e:
            print(f"[SKIP stat error] {p} / {e}"); continue
        ok.append(p)
    return ok

def _chunk(lst: List[Path], n: int) -> List[List[Path]]:
    if n <= 0: return [lst[:]]
    return [lst[i:i+n] for i in range(0, len(lst), n)]

# ==================== COM 起動 ====================
def _get_word_app():
    try:
        gencache.EnsureDispatch("Word.Application")   # makepy 生成
    except Exception:
        pass
    app = DispatchEx("Word.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    try: app.ScreenUpdating = False
    except Exception: pass
    app.DisplayAlerts = 0
    try:
        app.AutomationSecurity = msoAutomationSecurityForceDisable
    except Exception:
        pass
    try:
        app.Options.CheckGrammarAsYouType = False
        app.Options.CheckSpellingAsYouType = False
        app.Options.AllowReadingMode = False
    except Exception:
        pass
    return app

def _get_ppt_app():
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
    except Exception:
        pass
    app = DispatchEx("PowerPoint.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("PowerPoint.Application")
    try: app.Visible = msoFalse
    except Exception: pass
    try:
        app.AutomationSecurity = msoAutomationSecurityForceDisable
    except Exception:
        pass
    return app

# ==================== COM一時障害判定 ====================
def _is_transient_com_error(exc: Exception) -> bool:
    CODES = {-2147417848, 0x80010108, -2147023174, 0x800706BA}
    try:
        hr = getattr(exc, "hresult", None)
        if hr is None and getattr(exc, "args", None):
            hr = exc.args[0]
        if isinstance(hr, int) and hr in CODES:
            return True
        msg = (str(exc) or "").upper()
        return ("80010108" in msg or "RPC_E_DISCONNECTED" in msg or
                "800706BA" in msg or "RPC サーバーを利用できません" in msg)
    except Exception:
        return False

# ==================== PPT テキスト抽出（高速版） ====================
def _iter_shape_texts_fast(shape) -> Iterable[str]:
    # グループ（例外に頼らず Type 判定）
    try:
        if int(getattr(shape, "Type", 0)) == msoGroup:
            gi = shape.GroupItems
            for i in range(1, gi.Count + 1):
                yield from _iter_shape_texts_fast(gi.Item(i))
            return
    except Exception:
        pass

    # 表：HasText 判定を省略し、TextRange.Text を直接取得（COM往復を最小化）
    try:
        if int(getattr(shape, "HasTable", 0)) == msoTrue:
            table = shape.Table
            rows, cols = table.Rows.Count, table.Columns.Count
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    try:
                        s = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                        if s: yield s
                    except Exception:
                        if PPT_USE_TEXTFRAME2_FALLBACK:
                            try:
                                s2 = table.Cell(r, c).Shape.TextFrame2.TextRange.Text
                                if s2: yield s2
                            except Exception:
                                pass
            return
    except Exception:
        pass

    # 通常テキスト
    try:
        if int(getattr(shape, "HasTextFrame", 0)) == msoTrue:
            tf = shape.TextFrame
            if int(getattr(tf, "HasText", 0)) == msoTrue:
                yield tf.TextRange.Text
                return
    except Exception:
        pass

    # フォールバック（必要時のみ）
    if PPT_USE_TEXTFRAME2_FALLBACK:
        try:
            tf2 = shape.TextFrame2
            if int(getattr(tf2, "HasText", 0)) == msoTrue:
                yield tf2.TextRange.Text
        except Exception:
            pass

def extract_ppt_text(src: Path, ppt_app) -> Dict[int, List[str]]:
    created_here = False
    if ppt_app is None:
        ppt_app = _get_ppt_app(); created_here = True

    pres = None
    result: Dict[int, List[str]] = {}
    try:
        pres = ppt_app.Presentations.Open(FileName=str(src), WithWindow=False, ReadOnly=True)
        # スライドごとに Shapes を走査（インデクサでアクセスを固定）
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            shapes = slide.Shapes
            paras: List[str] = []
            for j in range(1, shapes.Count + 1):
                sh = shapes(j)
                for raw in _iter_shape_texts_fast(sh):
                    if raw:
                        paras.extend(_split_paragraphs_fast(raw))
            result[s_idx] = _ensure_paragraphs_list(paras)
    finally:
        if pres is not None:
            pres.Close()
        if created_here and not DEFER_QUIT_TO_PARENT and ppt_app is not None:
            try: ppt_app.Quit()
            except Exception: pass
    return result
def _extract_ppt_safe_text(src: Path, mgr, retries: int = DEFAULT_RETRIES) -> Dict[int, List[str]]:
    delay = 0.5
    for attempt in range(retries + 1):
        try:
            return extract_ppt_text(src, mgr.ensure())
        except Exception as e:
            if _is_transient_com_error(e) and attempt < retries:
                print(f"[PPT reconnect] {src.name} / retry {attempt+1}")
                mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
            raise
# ==================== Word テキスト抽出（高速版） ====================
def _extract_word_plain_paragraphs(doc) -> List[str]:
    """
    旧：Paragraphs を 1件ずつ列挙 → 往復が多く低速
    新：doc.Content.Text を一括取得して段落分割（圧倒的に往復が少ない）
    """
    try:
        t = doc.Content.Text  # COM呼び出し1回
        return _ensure_paragraphs_list(_split_paragraphs_fast(t))
    except Exception:
        # フォールバック（まれに Content が読めない場面）
        paras: List[str] = []
        try:
            plist = doc.Paragraphs
            for i in range(1, plist.Count + 1):
                try:
                    paras.extend(_split_paragraphs_fast(plist(i).Range.Text))
                except Exception:
                    continue
        except Exception:
            pass
        return _ensure_paragraphs_list(paras)

def _extract_word_story_texts(doc) -> List[str]:
    """
    StoryRanges を for-in ではなく NextStoryRange でつないで走査し、
    各 Range.Text を一発取得（COM往復を削減）
    """
    texts: List[str] = []
    try:
        # 先頭ストーリーを取得
        try:
            rng = doc.StoryRanges(1)
        except Exception:
            rng = None
        visited = 0
        while rng is not None and visited < 512:  # 無限ループ対策の上限
            try:
                t = getattr(rng, "Text", "")
                if t: texts.extend(_split_paragraphs_fast(t))
            except Exception:
                pass
            # 次へ
            try:
                rng = rng.NextStoryRange
            except Exception:
                break
            visited += 1
    except Exception:
        # フォールバック（旧実装）
        try:
            stories = doc.StoryRanges
            try:
                for r in stories:
                    if r is None: continue
                    t = getattr(r, "Text", "")
                    if t: texts.extend(_split_paragraphs_fast(t))
            except Exception:
                count = stories.Count
                for i in range(1, count + 1):
                    r = stories(i)
                    if r is None: continue
                    t = getattr(r, "Text", "")
                    if t: texts.extend(_split_paragraphs_fast(t))
        except Exception:
            pass
    return _ensure_paragraphs_list(texts)

def extract_word_text(src: Path, word_app) -> List[str]:
    created_here = False
    if word_app is None:
        word_app = _get_word_app(); created_here = True

    doc = None
    try:
        doc = word_app.Documents.Open(
            FileName=str(src),
            ReadOnly=True, AddToRecentFiles=False,
            OpenAndRepair=False, ConfirmConversions=False, NoEncodingDialog=True,
            Visible=False
        )
        # 一括取得で主文書を抽出
        paras = _extract_word_plain_paragraphs(doc)
        if USE_WORD_STORY_RANGES:
            extras = _extract_word_story_texts(doc)
            if extras:
                seen = set()
                out: List[str] = []
                # 順序を極力保ちながら重複除去
                for t in paras + extras:
                    if t not in seen:
                        seen.add(t); out.append(t)
                paras = out
        return paras
    finally:
        if doc is not None:
            doc.Close(False)
        if created_here and not DEFER_QUIT_TO_PARENT and word_app is not None:
            try: word_app.Quit()
            except Exception: pass
def _extract_word_safe_text(src: Path, mgr, retries: int = DEFAULT_RETRIES) -> List[str]:
    delay = 0.5
    for attempt in range(retries + 1):
        try:
            return extract_word_text(src, mgr.ensure())
        except Exception as e:
            if _is_transient_com_error(e) and attempt < retries:
                print(f"[WORD reconnect] {src.name} / retry {attempt+1}")
                mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
            raise
# ==================== マネージャ ====================
class _WordManager:
    def __init__(self): self.app = None
    def ensure(self):
        if self.app is None: self.app = _get_word_app()
        return self.app
    def reset(self):
        try:
            if self.app is not None and not DEFER_QUIT_TO_PARENT:
                self.app.Quit()
        except Exception: pass
        self.app = None
        time.sleep(0.1)
    def close(self):
        if DEFER_QUIT_TO_PARENT: return
        self.reset()

class _PptManager:
    def __init__(self): self.app = None
    def ensure(self):
        if self.app is None: self.app = _get_ppt_app()
        return self.app
    def reset(self):
        try:
            if self.app is not None and not DEFER_QUIT_TO_PARENT:
                self.app.Quit()
        except Exception: pass
        self.app = None
        time.sleep(0.1)
    def close(self):
        if DEFER_QUIT_TO_PARENT: return
        self.reset()

# ==================== HTML 生成（Hタグ維持・テキストのみ） ====================
def entries_to_html(entries: List[Tuple[str, Dict]]) -> str:
    """
    entries: [(display_name, data_dict), ...]
      - Word: {"type":"word","paras":[...]}
      - PPT : {"type":"ppt","slides": {no: [paras...]}}
    """
    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="ja">')
    parts.append('<meta charset="UTF-8">')
    parts.append("<title>抽出テキスト</title>")
    parts.append("<style>ul{margin:.3rem 1.2rem;} h2{margin-top:1.2rem;} .word-block,.slide-block{border:1px solid #e3e3e3;padding:.6rem;border-radius:.5rem;} .file-block{margin-bottom:1rem;}</style>")
    parts.append("<body>")

    for file_disp, data in entries:
        parts.append(f"<div class='file-block'><h2>{_html.escape(file_disp)}</h2>")
        if data.get("type") == "ppt":
            slides = data.get("slides", {})
            for slide_no in sorted(slides.keys()):
                paras = slides[slide_no] or []
                parts.append(f"<div class='slide-block'><h3>Slide {slide_no}</h3>")
                if paras:
                    parts.append("<ul>")
                    for p in paras:
                        parts.append(f"<li>{_html.escape(p).replace('\\n','<br>')}</li>")
                    parts.append("</ul>")
                parts.append("</div>")
        else:
            paras = data.get("paras", [])
            if paras:
                parts.append("<div class='word-block'><ul>")
                for p in paras:
                    parts.append(f"<li>{_html.escape(p).replace('\\n','<br>')}</li>")
                parts.append("</ul></div>")
        parts.append("</div>")
    parts.append("</body></html>")
    return "\n".join(parts)

# ==================== ワーカー（パート単位 / 並列維持） ====================
def _worker_make_part(part_index: int, paths: List[str], out_dir: str, out_stem: str) -> tuple[int, str, int]:
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    # 起動ジッタで同時COMアクティベーション衝突を緩和
    time.sleep(0.2 + (part_index % 4) * 0.15 + random.uniform(0.0, 0.05))

    word_mgr = _WordManager()
    ppt_mgr  = _PptManager()
    entries: List[Tuple[str, Dict]] = []

    try:
        total = len(paths)
        for i, s in enumerate(paths, 1):
            src = Path(s)
            if (not src.exists()) or (src.suffix.lower() not in ALLOWED_EXTS):
                continue
            print(f"[part{part_index}] {i}/{total}: {src.name}")
            try:
                if src.suffix.lower() in WORD_EXTS:
                    paras = _extract_word_safe_text(src, word_mgr)
                    if paras:
                        entries.append( (src.name, {"type": "word", "paras": paras}) )
                else:
                    slides = _extract_ppt_safe_text(src, ppt_mgr)
                    if slides:
                        entries.append( (src.name, {"type": "ppt", "slides": slides}) )
            except Exception as e:
                print(f"[part{part_index} SKIP] {src} / {e}")

        out_dir_p = Path(out_dir)
        out_dir_p.mkdir(parents=True, exist_ok=True)
        out_path = out_dir_p / f"{out_stem}_part{part_index}.html"
        html_text = entries_to_html(entries) if entries else "<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし）</p></body>"
        out_path.write_text(html_text, encoding="utf-8")
        return (part_index, str(out_path), len(entries))

    finally:
        try: word_mgr.close()
        except Exception: pass
        try: ppt_mgr.close()
        except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass

# ==================== メイン API（関数名・並列仕様そのまま） ====================
def convert_office_to_html(
    path_lst: Iterable[str | Path | Tuple[str, str]],
    output_html_path: str | Path,
    html_dir: str,
    batch_size: int = BATCH_SIZE_DEFAULT,
    size_kb_limit: int = MAX_FILE_KB_DEFAULT,
    max_agents: Optional[int] = None,
) -> List[Path]:
    """
    path_lst: "C:/a/b.docx" または (dir, filename) の混在でOK
    output_html_path: 出力ルートフォルダ（例: "C:/out"）
    html_dir: まとめファイル名（例: "result.html" → 実際は result_partN.html）
    """
    if KILL_AT_START:
        kill_office_processes()

    items = _normalize_items(path_lst)
    filtered = _prefilter_items(items, max_kb=size_kb_limit)

    out_base = Path(output_html_path) / html_dir
    out_dir  = out_base.parent
    out_stem = out_base.stem

    if not filtered:
        empty = out_dir / f"{out_stem}_part1.html"
        empty.parent.mkdir(parents=True, exist_ok=True)
        empty.write_text("<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし：対象0件）</p></body>", encoding="utf-8")
        print(f"[WRITE empty] {empty}")
        if KILL_AT_END: kill_office_processes()
        return [empty]

    chunks = _chunk(filtered, batch_size)
    total_parts = len(chunks)
    workers = total_parts if max_agents is None else max(1, min(total_parts, max_agents))

    print(f"[PLAN] 有効 {len(filtered)} 件 → {total_parts} part / 同時実行 {workers} / {batch_size}件/part")
    for i, ch in enumerate(chunks, 1):
        print(f"  - part{i}: {len(ch)} 件")

    tasks = [(idx + 1, [str(p) for p in chunk], str(out_dir), out_stem)
             for idx, chunk in enumerate(chunks)]

    results = []
    with ProcessPoolExecutor(max_workers=workers) as ex:
        futs = [ex.submit(_worker_make_part, *t) for t in tasks]
        for f in as_completed(futs):
            results.append(f.result())

    if DEFER_QUIT_TO_PARENT and KILL_AT_END:
        kill_office_processes()

    results.sort(key=lambda x: x[0])
    for idx, path_str, cnt in results:
        print(f"[DONE part{idx}] {path_str}  (収録 {cnt} ファイル)")

    return [Path(p) for _, p, _ in results]

# # =============== 直接実行テストの例 ===============
# if __name__ == "__main__":
#     path_list = [
#         (r"C:\Users\yohei\Downloads\R2-2206906", "R2-2206906_C1-224008.docx"),
#         (r"C:\Users\yohei\Downloads\R2-2206905", "R2-2206905_C1-223972.docx"),
#         (r"C:\Users\yohei\Downloads", "11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx"),
#     ]
#     out_dir  = r"C:\Users\yohei\Downloads\html_out"
#     out_file = "a.html"  # 実際には a_part1.html など分割生成
#     outputs = convert_office_to_html(
#         path_lst=path_list,
#         output_html_path=out_dir,
#         html_dir=out_file,
#         batch_size=5,
#         size_kb_limit=1500,
#         # max_agents=4,
#     )
#     for p in outputs:
#         print(" -", p)


# 直接実行テスト
if __name__ == "__main__":
    path_list = [
        (r"C:\Users\yohei\Downloads", "11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx"),
        (r"C:\Users\yohei\Downloads", "R2-1903631_Discussion on SR BSR for NR Sidelink mode 1.doc"),
    ]
    out_dir  = r"C:\Users\yohei\Downloads"
    out_file = "a.html"  # 実際には a_part1.html など分割生成
    outputs = convert_office_to_html(
        path_lst=path_list,
        output_html_path=out_dir,
        html_dir=out_file,
        batch_size=5,
        size_kb_limit=1500,
        # max_agents=4,  # PPT が多い場合は 3〜4 程度を推奨
    )
    print("✅ 出力一覧:")
    for p in outputs:
        print(" -", p)
