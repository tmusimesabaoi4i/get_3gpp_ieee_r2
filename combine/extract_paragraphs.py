# -*- coding: utf-8 -*-
"""
Word/PPTX 段落を抽出して <li> 化、Heading でファイル名/スライド番号を示す。
要件:
- 開始前に WINWORD.EXE / POWERPNT.EXE を kill（cmdの taskkill を呼ぶ）
- 対象外拡張子はスキップ
- ファイルを開けない場合のみ print してスキップ（他は出力しない）
- CSS なし
- Windows + Microsoft Office + pywin32 前提
"""

# from __future__ import annotations
# import html
# import subprocess
# from pathlib import Path
# from typing import Iterable, List, Dict

# import pythoncom
# import win32com.client as win32

# WORD_EXTS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf"}
# PPT_EXTS  = {".ppt", ".pptx", ".pptm"}


# -------------------- 起動前に Office プロセスを kill --------------------
# def kill_office_processes() -> None:
#     """
#     WINWORD.EXE と POWERPNT.EXE を強制終了。
#     失敗しても例外は投げず、静かに無視（printもしない）。
#     ※ 強制終了により未保存データは失われます。
#     """
#     # /F: 強制終了, /T: 子プロセス含む, /IM: 画像名指定
#     for image_name in ("WINWORD.EXE", "POWERPNT.EXE"):
#         try:
#             subprocess.run(
#                 ["taskkill", "/IM", image_name, "/F", "/T"],
#                 capture_output=True,
#                 text=True,
#                 check=False
#             )
#         except Exception:
#             pass


# # -------------------- 共通ユーティリティ --------------------
# def _clean_paragraph_text(s: str) -> str:
#     if s is None:
#         return ""
#     s = s.replace("\r", "").replace("\x07", "")
#     return s.strip()

# def _ensure_paragraphs_list(items: Iterable[str]) -> List[str]:
#     out: List[str] = []
#     for t in items:
#         t = _clean_paragraph_text(t)
#         if t:
#             out.append(t)
#     return out


# # -------------------- Word: 段落抽出 --------------------
# def extract_paragraphs_from_word(path: Path, word_app=None) -> List[str]:
#     created_here = False
#     if word_app is None:
#         word_app = win32.DispatchEx("Word.Application")
#         word_app.Visible = False
#         created_here = True

#     doc = None
#     paras: List[str] = []
#     try:
#         doc = word_app.Documents.Open(
#             FileName=str(path),
#             ReadOnly=True,
#             AddToRecentFiles=False
#         )
#         for para in doc.Paragraphs:  # 1-based
#             text = _clean_paragraph_text(str(para.Range.Text))
#             if text:
#                 paras.append(text)
#     finally:
#         if doc is not None:
#             doc.Close(False)
#         if created_here and word_app is not None:
#             word_app.Quit()
#     return paras


# # -------------------- PowerPoint: 段落抽出(スライド別) --------------------
# def _iter_shape_paragraphs(shape) -> Iterable[str]:
#     # グループ図形（中を再帰）
#     try:
#         gi = shape.GroupItems
#         for i in range(1, gi.Count + 1):
#             yield from _iter_shape_paragraphs(gi.Item(i))
#         return
#     except Exception:
#         pass

#     # 表（セル内テキスト）
#     try:
#         if getattr(shape, "HasTable", 0) == -1:
#             table = shape.Table
#             for r in range(1, table.Rows.Count + 1):
#                 for c in range(1, table.Columns.Count + 1):
#                     cell_shape = table.Cell(r, c).Shape
#                     if getattr(cell_shape, "TextFrame", None) is not None:
#                         tf = cell_shape.TextFrame
#                         if getattr(tf, "HasText", 0) == -1:
#                             tr = tf.TextRange
#                             count = tr.Paragraphs().Count
#                             for i in range(1, count + 1):
#                                 yield _clean_paragraph_text(tr.Paragraphs(i).Text)
#             return
#     except Exception:
#         pass

#     # 通常の TextFrame
#     try:
#         if getattr(shape, "HasTextFrame", 0) == -1:
#             tf = shape.TextFrame
#             if getattr(tf, "HasText", 0) == -1:
#                 tr = tf.TextRange
#                 count = tr.Paragraphs().Count
#                 for i in range(1, count + 1):
#                     yield _clean_paragraph_text(tr.Paragraphs(i).Text)
#     except Exception:
#         pass


# def extract_paragraphs_from_ppt_grouped(path: Path, ppt_app=None) -> Dict[int, List[str]]:
#     created_here = False
#     if ppt_app is None:
#         ppt_app = win32.DispatchEx("PowerPoint.Application")
#         created_here = True

#     pres = None
#     grouped: Dict[int, List[str]] = {}
#     try:
#         pres = ppt_app.Presentations.Open(
#             FileName=str(path),
#             WithWindow=False,
#             ReadOnly=True
#         )
#         for slide in pres.Slides:
#             slide_no = slide.SlideIndex  # 1-based
#             buf: List[str] = []
#             for shape in slide.Shapes:
#                 for p in _iter_shape_paragraphs(shape):
#                     if p:
#                         buf.append(p)
#             grouped[slide_no] = _ensure_paragraphs_list(buf)
#     finally:
#         if pres is not None:
#             pres.Close()
#         if created_here and ppt_app is not None:
#             ppt_app.Quit()
#     return grouped


# # -------------------- HTML 生成 --------------------
# def paragraphs_to_html(grouped: Dict[str, Dict]) -> str:
#     """
#     grouped:
#       {
#         "file_display": {
#             "type": "word", "paras": [..]
#             # or
#             "type": "ppt",  "slides": { 1:[..], 2:[..], ... }
#         },
#         ...
#       }
#     """
#     parts: List[str] = []
#     parts.append("<!DOCTYPE html>")
#     parts.append('<html lang="en">')
#     parts.append('<meta charset="UTF-8">')
#     parts.append("<title>抽出テキスト</title>")
#     parts.append("<body>")

#     for file_disp, data in grouped.items():
#         parts.append(f"<h2>{html.escape(file_disp)}</h2>")

#         if data.get("type") == "ppt":
#             slides: Dict[int, List[str]] = data.get("slides", {})
#             for slide_no in sorted(slides.keys()):
#                 parts.append(f"<h3>Slide {slide_no}</h3>")
#                 paras = slides[slide_no]
#                 if paras:
#                     parts.append("<ul>")
#                     for p in paras:
#                         safe = html.escape(p).replace("\n", "<br>")
#                         parts.append(f"<li>{safe}</li>")
#                     parts.append("</ul>")
#         else:
#             paras = data.get("paras", [])
#             if paras:
#                 parts.append("<ul>")
#                 for p in paras:
#                     safe = html.escape(p).replace("\n", "<br>")
#                     parts.append(f"<li>{safe}</li>")
#                 parts.append("</ul>")

#     parts.append("</body></html>")
#     return "\n".join(parts)


# # -------------------- メイン --------------------
# def convert_office_to_html(
#     # dir_list: Iterable[str],
#     # file_list: Iterable[str],
#     path_lst: Iterable[str],
#     output_html_path: str | Path,
# ) -> Path:
#     """
#     dir_list と file_list を zip し、各ファイルを処理。
#     - 開始前に Word/PowerPoint プロセスを kill
#     - 対象外拡張子と存在しないファイルは静かにスキップ
#     - 開けないファイルは print してスキップ
#     - Word: <h2>ファイル名</h2> の直下に <li>
#     - PowerPoint: <h2>ファイル名</h2> の下に <h3>Slide n</h3> + <li>
#     """
#     # ① Office プロセスを終了
#     kill_office_processes()

#     # ② COM 初期化
#     pythoncom.CoInitialize()

#     word_app = None
#     ppt_app  = None
#     grouped: Dict[str, Dict] = {}

#     all_n = len(path_lst)
#     rep = 1

#     try:
#         # ③ 対象ファイルをループ処理
#         for item in path_lst:  # ← zip(...) にしない！
#             # item が "(dir, filename)" のペアか、フルパスかを判定
#             if isinstance(item, (tuple, list)):
#                 if len(item) >= 2:
#                     d, f = item[0], item[1]
#                     src = Path(str(d)) / str(f)
#                 else:
#                     # 要素数が1のタプル/リストならその要素をパスとして扱う
#                     src = Path(str(item[0]))
#             else:
#                 # フルパス（文字列 or Path）想定
#                 src = Path(str(item))

#             # 存在しないファイルは静かにスキップ
#             if not src.exists():
#                 continue

#             ext = src.suffix.lower()
#             # 対象外拡張子はスキップ
#             if ext not in WORD_EXTS and ext not in PPT_EXTS:
#                 continue

#             file_disp = src.name  # 見出し表示
#             print("[combining...."+str(round(rep/all_n*100))+" % is done]")
#             rep = rep + 1
#             try:
#                 if ext in WORD_EXTS:
#                     if word_app is None:
#                         word_app = win32.DispatchEx("Word.Application")
#                         word_app.Visible = False
#                     paras = extract_paragraphs_from_word(src, word_app)
#                     grouped[file_disp] = {"type": "word", "paras": _ensure_paragraphs_list(paras)}

#                 elif ext in PPT_EXTS:
#                     if ppt_app is None:
#                         ppt_app = win32.DispatchEx("PowerPoint.Application")
#                     slides = extract_paragraphs_from_ppt_grouped(src, ppt_app)
#                     grouped[file_disp] = {"type": "ppt", "slides": slides}

#             except Exception as e:
#                 # ★ ファイルが開けない等 → print してスキップ
#                 print(f"[SKIP] ファイルを開けませんでした: {src} / エラー: {e}")
#                 continue

#         html_text = paragraphs_to_html(grouped)
#         out_path = Path(output_html_path)
#         out_path.parent.mkdir(parents=True, exist_ok=True)
#         out_path.write_text(html_text, encoding="utf-8")
#         return out_path

#     finally:
#         try:
#             if word_app is not None:
#                 word_app.Quit()
#         except Exception:
#             pass
#         try:
#             if ppt_app is not None:
#                 ppt_app.Quit()
#         except Exception:
#             pass
#         pythoncom.CoUninitialize()
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
"""
Office(doc/docx/rtf/ppt/pptx/pptm) テキスト抽出 → 10件ごとに HTML (combined_partN.html)
- 1500KB 超は事前スキップ
- 並列: multiprocessing（プロセス）で各ワーカーが Office を起動
- COM 切断/サーバー不在は自動復旧（再起動＋指数バックオフ）
- 既定で「ワーカーは Quit せず」→ 全員終了後に親が一括 Kill
"""


import os, re, html, time, random, subprocess
from pathlib import Path
from typing import Iterable, List, Dict, Tuple

import pythoncom
from win32com.client import gencache, DispatchEx
from concurrent.futures import ProcessPoolExecutor, as_completed

# ==================== 調整フラグ ====================
BATCH_SIZE_DEFAULT      = 5         # 1 HTML の最大収録件数
MAX_FILE_KB_DEFAULT     = 1500      # これ超は最初からスキップ
DEFAULT_RETRIES         = 1         # COM切断時の再試行回数（各ファイル）
EXCLUSIVE_INSTANCE      = True      # True: DispatchEx で専用インスタンス化
DEFER_QUIT_TO_PARENT    = True      # True: ワーカーでは Quit せず、最後に親が一括 Kill
KILL_AT_START           = True      # True: 開始前に既存 Office を kill
KILL_AT_END             = True      # True: 全ワーカー終了後に Office を kill

# ==================== 定数/拡張子 ====================
WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
PPT_EXTS  = {".ppt", ".pptx", ".pptm"}
ALLOWED_EXTS = WORD_EXTS | PPT_EXTS

msoTrue  = -1
msoFalse = 0

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
def _clean_paragraph_text(s: str) -> str:
    if s is None: return ""
    s = s.replace("\r", "").replace("\x07", "")
    return s.strip()

def _ensure_paragraphs_list(items: Iterable[str]) -> List[str]:
    out: List[str] = []
    for t in items:
        t = _clean_paragraph_text(t)
        if t: out.append(t)
    return out

_para_splitter = re.compile(r"[\r\n]+")
def _split_paragraphs_fast(s: str) -> List[str]:
    if not s: return []
    return [ _clean_paragraph_text(x) for x in _para_splitter.split(s) if _clean_paragraph_text(x) ]

# ==================== COM アプリ生成 ====================
def _get_word_app():
    """
    EXCLUSIVE_INSTANCE=True の場合は DispatchEx で専用インスタンスを起動。
    makepy は gencache.EnsureDispatch を一度呼んで生成済みでOK（速度向上）。
    """
    try:
        gencache.EnsureDispatch("Word.Application")  # makepy 生成だけ先に（速度向上）
    except Exception:
        pass
    app = DispatchEx("Word.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    try: app.ScreenUpdating = False
    except Exception: pass
    app.DisplayAlerts = 0
    try:
        app.Options.CheckGrammarAsYouType = False
        app.Options.CheckSpellingAsYouType = False
        app.Options.AllowReadingMode = False
    except Exception:
        pass
    return app

def _get_ppt_app():
    try:
        gencache.EnsureDispatch("PowerPoint.Application")  # makepy 生成
    except Exception:
        pass
    app = DispatchEx("PowerPoint.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("PowerPoint.Application")
    try: app.Visible = msoFalse
    except Exception: pass
    # 一部環境で効く抑止（PowerPoint は未サポートでも try で吸収）
    try:
        from win32com.client import constants as c
        app.AutomationSecurity = getattr(c, "msoAutomationSecurityForceDisable", 3)
    except Exception:
        pass
    return app

# ==================== Word/PPT 抽出 ====================
def extract_paragraphs_from_word(path: Path, word_app=None) -> List[str]:
    created_here = False
    if word_app is None:
        word_app = _get_word_app(); created_here = True
    doc = None
    try:
        doc = word_app.Documents.Open(
            FileName=str(path),
            ReadOnly=True, AddToRecentFiles=False,
            OpenAndRepair=False, ConfirmConversions=False, NoEncodingDialog=True
        )
        text = doc.Content.Text
        return _split_paragraphs_fast(text)
    finally:
        if doc is not None:
            doc.Close(False)
        if created_here and not DEFER_QUIT_TO_PARENT and word_app is not None:
            try: word_app.Quit()
            except Exception: pass

def _iter_shape_texts_fast(shape) -> Iterable[str]:
    # グループ
    try:
        gi = shape.GroupItems
        for i in range(1, gi.Count + 1):
            yield from _iter_shape_texts_fast(gi.Item(i))
        return
    except Exception:
        pass
    # 表
    try:
        if int(getattr(shape, "HasTable", 0)) == msoTrue:
            table = shape.Table
            for r in range(1, table.Rows.Count + 1):
                for c in range(1, table.Columns.Count + 1):
                    cell_shape = table.Cell(r, c).Shape
                    try:
                        tf = cell_shape.TextFrame
                        if int(getattr(tf, "HasText", 0)) == msoTrue:
                            yield tf.TextRange.Text; continue
                    except Exception: pass
                    try:
                        tf2 = cell_shape.TextFrame2
                        if int(getattr(tf2, "HasText", 0)) == msoTrue:
                            yield tf2.TextRange.Text
                    except Exception: pass
            return
    except Exception:
        pass
    # 通常テキスト
    try:
        if int(getattr(shape, "HasTextFrame", 0)) == msoTrue:
            tf = shape.TextFrame
            if int(getattr(tf, "HasText", 0)) == msoTrue:
                yield tf.TextRange.Text; return
    except Exception:
        pass
    # 予備
    try:
        tf2 = shape.TextFrame2
        if int(getattr(tf2, "HasText", 0)) == msoTrue:
            yield tf2.TextRange.Text
    except Exception:
        pass

def extract_paragraphs_from_ppt_grouped(path: Path, ppt_app=None) -> Dict[int, List[str]]:
    created_here = False
    if ppt_app is None:
        ppt_app = _get_ppt_app(); created_here = True
    pres = None
    grouped: Dict[int, List[str]] = {}
    try:
        pres = ppt_app.Presentations.Open(FileName=str(path), WithWindow=False, ReadOnly=True)
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            buf: List[str] = []
            shapes = slide.Shapes
            for j in range(1, shapes.Count + 1):
                for raw in _iter_shape_texts_fast(shapes(j)):
                    if raw: buf.extend(_split_paragraphs_fast(raw))
            grouped[s_idx] = _ensure_paragraphs_list(buf)
    finally:
        if pres is not None:
            pres.Close()
        if created_here and not DEFER_QUIT_TO_PARENT and ppt_app is not None:
            try: ppt_app.Quit()
            except Exception: pass
    return grouped

# ==================== HTML 生成 ====================
def paragraphs_to_html(grouped: Dict[str, Dict]) -> str:
    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="en">')
    parts.append('<meta charset="UTF-8">')
    parts.append("<title>抽出テキスト</title>")
    parts.append("<body>")
    for file_disp, data in grouped.items():
        parts.append(f"<h2>{html.escape(file_disp)}</h2>")
        if data.get("type") == "ppt":
            for slide_no in sorted(data.get("slides", {}).keys()):
                parts.append(f"<h3>Slide {slide_no}</h3>")
                paras = data["slides"][slide_no]
                if paras:
                    parts.append("<ul>")
                    for p in paras:
                        parts.append(f"<li>{html.escape(p).replace('\\n', '<br>')}</li>")
                    parts.append("</ul>")
        else:
            paras = data.get("paras", [])
            if paras:
                parts.append("<ul>")
                for p in paras:
                    parts.append(f"<li>{html.escape(p).replace('\\n', '<br>')}</li>")
                parts.append("</ul>")
    parts.append("</body></html>")
    return "\n".join(parts)

# ==================== 入力の正規化・分割 ====================
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

# ==================== COM 一時障害検知 ====================
def _is_transient_com_error(exc: Exception) -> bool:
    """RPC_E_DISCONNECTED(-2147417848) / RPC_S_SERVER_UNAVAILABLE(-2147023174) など"""
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

# ==================== アプリマネージャ & 安全抽出 ====================
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
        if DEFER_QUIT_TO_PARENT: return  # 親がまとめて Kill
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

def _extract_word_safe(src: Path, mgr: _WordManager, retries: int = DEFAULT_RETRIES) -> List[str]:
    delay = 0.5
    for attempt in range(retries + 1):
        try:
            return extract_paragraphs_from_word(src, mgr.ensure())
        except Exception as e:
            if _is_transient_com_error(e) and attempt < retries:
                print(f"[WORD reconnect] {src.name} / retry {attempt+1}")
                mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
            raise

def _extract_ppt_safe(src: Path, mgr: _PptManager, retries: int = DEFAULT_RETRIES) -> Dict[int, List[str]]:
    delay = 0.5
    for attempt in range(retries + 1):
        try:
            return extract_paragraphs_from_ppt_grouped(src, mgr.ensure())
        except Exception as e:
            if _is_transient_com_error(e) and attempt < retries:
                print(f"[PPT reconnect] {src.name} / retry {attempt+1}")
                mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
            raise

# ==================== ワーカー（プロセス） ====================
def _worker_make_part(part_index: int, paths: List[str], out_dir: str, out_stem: str) -> tuple[int, str, int]:
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    # 起動ジッタ：同時 COM アクティベーション衝突の回避
    time.sleep(0.3 + (part_index % 4) * 0.2 + random.uniform(0.0, 0.1))

    word_mgr = _WordManager()
    ppt_mgr  = _PptManager()
    grouped: Dict[str, Dict] = {}

    try:
        total = len(paths)
        for i, s in enumerate(paths, 1):
            src = Path(s)
            if (not src.exists()) or (src.suffix.lower() not in ALLOWED_EXTS):
                continue
            print(f"[part{part_index}] {i}/{total}: {src.name}")
            try:
                if src.suffix.lower() in WORD_EXTS:
                    paras = _extract_word_safe(src, word_mgr, retries=DEFAULT_RETRIES)
                    if paras: grouped[src.name] = {"type": "word", "paras": paras}
                else:
                    slides = _extract_ppt_safe(src, ppt_mgr, retries=DEFAULT_RETRIES)
                    if any(slides.values()): grouped[src.name] = {"type": "ppt", "slides": slides}
            except Exception as e:
                print(f"[part{part_index} SKIP] {src} / {e}")

        out_dir_p = Path(out_dir); out_dir_p.mkdir(parents=True, exist_ok=True)
        out_path = out_dir_p / f"{out_stem}_part{part_index}.html"
        html_text = paragraphs_to_html(grouped) if grouped else \
            "<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし）</p></body>"
        out_path.write_text(html_text, encoding="utf-8")
        return (part_index, str(out_path), len(grouped))

    finally:
        try: word_mgr.close()
        except Exception: pass
        try: ppt_mgr.close()
        except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass

# ==================== 並列メイン API ====================
def convert_office_to_html(
    path_lst: Iterable[str | Path | Tuple[str, str]],
    output_html_path: str | Path,
    batch_size: int = BATCH_SIZE_DEFAULT,
    size_kb_limit: int = MAX_FILE_KB_DEFAULT,
    max_agents: int | None = None,
) -> List[Path]:
    # 起動前 Kill（全員分の開始前に一度だけ）
    if KILL_AT_START:
        kill_office_processes()

    items = _normalize_items(path_lst)
    filtered = _prefilter_items(items, max_kb=size_kb_limit)

    out_base = Path(output_html_path)
    out_dir  = out_base.parent
    out_stem = out_base.stem

    if not filtered:
        empty = out_dir / f"{out_stem}_part1.html"
        empty.parent.mkdir(parents=True, exist_ok=True)
        empty.write_text("<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし：対象0件）</p></body>", encoding="utf-8")
        print(f"[WRITE empty] {empty}")
        # 終了後 Kill（要求に応じて）
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

    # 全員終了 → ここで一括 Kill（≒「誰かが作業中なら閉じない」）
    if DEFER_QUIT_TO_PARENT and KILL_AT_END:
        kill_office_processes()

    results.sort(key=lambda x: x[0])
    for idx, path_str, cnt in results:
        print(f"[DONE part{idx}] {path_str}  (収録 {cnt} ファイル)")

    return [Path(p) for _, p, _ in results]

# # ==================== 直呼び例（Windows は __main__ 必須） ====================
# if __name__ == "__main__":
#     sample_list = [
#         r"C:\Users\yohei\Downloads\DOCS\a.docx",
#         r"C:\Users\yohei\Downloads\DOCS\b.pptx",
#         (r"C:\Users\yohei\Downloads\DOCS", "c.rtf"),
#         (r"C:\Users\yohei\Downloads\DOCS", "d.pptm"),
#     ]
#     out_base = r"C:\Users\yohei\Downloads\combined.html"

#     outputs = convert_office_to_html_parallel(
#         path_lst=sample_list,
#         output_html_path=out_base,
#         batch_size=10,
#         size_kb_limit=1500,
#         # max_agents=4,   # PPTが多い時は 3〜4 推奨
#     )
#     print("出力一覧:")
#     for p in outputs:
#         print(" -", p)




# 直接実行テスト
if __name__ == "__main__":
    dirs = [r"C:\Users\yohei\Downloads\R2-2206906", r"C:\Users\yohei\Downloads\R2-2206905", r"C:\Users\yohei\Downloads",r"C:\Users\yohei\Downloads"]
    files = ["R2-2206906_C1-224008.docx", "R2-2206905_C1-223972.docx","11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx","input_3gpp.xlsx"]
    output = r"C:\Users\yohei\Downloads\a.html"
    out = convert_office_to_html(dirs, files, output)
    print(f"✅ 出力: {out}")
