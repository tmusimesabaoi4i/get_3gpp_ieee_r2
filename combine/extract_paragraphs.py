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


# import os, re, html, time, random, subprocess
# from pathlib import Path
# from typing import Iterable, List, Dict, Tuple

# import pythoncom
# from win32com.client import gencache, DispatchEx
# from concurrent.futures import ProcessPoolExecutor, as_completed

# # ==================== 調整フラグ ====================
# BATCH_SIZE_DEFAULT      = 5         # 1 HTML の最大収録件数
# MAX_FILE_KB_DEFAULT     = 1500      # これ超は最初からスキップ
# DEFAULT_RETRIES         = 1         # COM切断時の再試行回数（各ファイル）
# EXCLUSIVE_INSTANCE      = True      # True: DispatchEx で専用インスタンス化
# DEFER_QUIT_TO_PARENT    = True      # True: ワーカーでは Quit せず、最後に親が一括 Kill
# KILL_AT_START           = True      # True: 開始前に既存 Office を kill
# KILL_AT_END             = True      # True: 全ワーカー終了後に Office を kill

# # ==================== 定数/拡張子 ====================
# WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
# PPT_EXTS  = {".ppt", ".pptx", ".pptm"}
# ALLOWED_EXTS = WORD_EXTS | PPT_EXTS

# msoTrue  = -1
# msoFalse = 0

# # ==================== プロセス一括 Kill ====================
# def kill_office_processes() -> None:
#     for image_name in ("WINWORD.EXE", "POWERPNT.EXE"):
#         try:
#             subprocess.run(
#                 ["taskkill", "/IM", image_name, "/F", "/T"],
#                 capture_output=True, text=True, check=False
#             )
#         except Exception:
#             pass

# # ==================== ユーティリティ ====================
# def _clean_paragraph_text(s: str) -> str:
#     if s is None: return ""
#     s = s.replace("\r", "").replace("\x07", "")
#     return s.strip()

# def _ensure_paragraphs_list(items: Iterable[str]) -> List[str]:
#     out: List[str] = []
#     for t in items:
#         t = _clean_paragraph_text(t)
#         if t: out.append(t)
#     return out

# _para_splitter = re.compile(r"[\r\n]+")
# def _split_paragraphs_fast(s: str) -> List[str]:
#     if not s: return []
#     return [ _clean_paragraph_text(x) for x in _para_splitter.split(s) if _clean_paragraph_text(x) ]

# # ==================== COM アプリ生成 ====================
# def _get_word_app():
#     """
#     EXCLUSIVE_INSTANCE=True の場合は DispatchEx で専用インスタンスを起動。
#     makepy は gencache.EnsureDispatch を一度呼んで生成済みでOK（速度向上）。
#     """
#     try:
#         gencache.EnsureDispatch("Word.Application")  # makepy 生成だけ先に（速度向上）
#     except Exception:
#         pass
#     app = DispatchEx("Word.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("Word.Application")
#     app.Visible = False
#     try: app.ScreenUpdating = False
#     except Exception: pass
#     app.DisplayAlerts = 0
#     try:
#         app.Options.CheckGrammarAsYouType = False
#         app.Options.CheckSpellingAsYouType = False
#         app.Options.AllowReadingMode = False
#     except Exception:
#         pass
#     return app

# def _get_ppt_app():
#     try:
#         gencache.EnsureDispatch("PowerPoint.Application")  # makepy 生成
#     except Exception:
#         pass
#     app = DispatchEx("PowerPoint.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("PowerPoint.Application")
#     try: app.Visible = msoFalse
#     except Exception: pass
#     # 一部環境で効く抑止（PowerPoint は未サポートでも try で吸収）
#     try:
#         from win32com.client import constants as c
#         app.AutomationSecurity = getattr(c, "msoAutomationSecurityForceDisable", 3)
#     except Exception:
#         pass
#     return app

# # ==================== Word/PPT 抽出 ====================
# def extract_paragraphs_from_word(path: Path, word_app=None) -> List[str]:
#     created_here = False
#     if word_app is None:
#         word_app = _get_word_app(); created_here = True
#     doc = None
#     try:
#         doc = word_app.Documents.Open(
#             FileName=str(path),
#             ReadOnly=True, AddToRecentFiles=False,
#             OpenAndRepair=False, ConfirmConversions=False, NoEncodingDialog=True
#         )
#         text = doc.Content.Text
#         return _split_paragraphs_fast(text)
#     finally:
#         if doc is not None:
#             doc.Close(False)
#         if created_here and not DEFER_QUIT_TO_PARENT and word_app is not None:
#             try: word_app.Quit()
#             except Exception: pass

# def _iter_shape_texts_fast(shape) -> Iterable[str]:
#     # グループ
#     try:
#         gi = shape.GroupItems
#         for i in range(1, gi.Count + 1):
#             yield from _iter_shape_texts_fast(gi.Item(i))
#         return
#     except Exception:
#         pass
#     # 表
#     try:
#         if int(getattr(shape, "HasTable", 0)) == msoTrue:
#             table = shape.Table
#             for r in range(1, table.Rows.Count + 1):
#                 for c in range(1, table.Columns.Count + 1):
#                     cell_shape = table.Cell(r, c).Shape
#                     try:
#                         tf = cell_shape.TextFrame
#                         if int(getattr(tf, "HasText", 0)) == msoTrue:
#                             yield tf.TextRange.Text; continue
#                     except Exception: pass
#                     try:
#                         tf2 = cell_shape.TextFrame2
#                         if int(getattr(tf2, "HasText", 0)) == msoTrue:
#                             yield tf2.TextRange.Text
#                     except Exception: pass
#             return
#     except Exception:
#         pass
#     # 通常テキスト
#     try:
#         if int(getattr(shape, "HasTextFrame", 0)) == msoTrue:
#             tf = shape.TextFrame
#             if int(getattr(tf, "HasText", 0)) == msoTrue:
#                 yield tf.TextRange.Text; return
#     except Exception:
#         pass
#     # 予備
#     try:
#         tf2 = shape.TextFrame2
#         if int(getattr(tf2, "HasText", 0)) == msoTrue:
#             yield tf2.TextRange.Text
#     except Exception:
#         pass

# def extract_paragraphs_from_ppt_grouped(path: Path, ppt_app=None) -> Dict[int, List[str]]:
#     created_here = False
#     if ppt_app is None:
#         ppt_app = _get_ppt_app(); created_here = True
#     pres = None
#     grouped: Dict[int, List[str]] = {}
#     try:
#         pres = ppt_app.Presentations.Open(FileName=str(path), WithWindow=False, ReadOnly=True)
#         for s_idx in range(1, pres.Slides.Count + 1):
#             slide = pres.Slides(s_idx)
#             buf: List[str] = []
#             shapes = slide.Shapes
#             for j in range(1, shapes.Count + 1):
#                 for raw in _iter_shape_texts_fast(shapes(j)):
#                     if raw: buf.extend(_split_paragraphs_fast(raw))
#             grouped[s_idx] = _ensure_paragraphs_list(buf)
#     finally:
#         if pres is not None:
#             pres.Close()
#         if created_here and not DEFER_QUIT_TO_PARENT and ppt_app is not None:
#             try: ppt_app.Quit()
#             except Exception: pass
#     return grouped

# # ==================== HTML 生成 ====================
# def paragraphs_to_html(grouped: Dict[str, Dict]) -> str:
#     parts: List[str] = []
#     parts.append("<!DOCTYPE html>")
#     parts.append('<html lang="en">')
#     parts.append('<meta charset="UTF-8">')
#     parts.append("<title>抽出テキスト</title>")
#     parts.append("<body>")
#     for file_disp, data in grouped.items():
#         parts.append(f"<h2>{html.escape(file_disp)}</h2>")
#         if data.get("type") == "ppt":
#             for slide_no in sorted(data.get("slides", {}).keys()):
#                 parts.append(f"<h3>Slide {slide_no}</h3>")
#                 paras = data["slides"][slide_no]
#                 if paras:
#                     parts.append("<ul>")
#                     for p in paras:
#                         parts.append(f"<li>{html.escape(p).replace('\\n', '<br>')}</li>")
#                     parts.append("</ul>")
#         else:
#             paras = data.get("paras", [])
#             if paras:
#                 parts.append("<ul>")
#                 for p in paras:
#                     parts.append(f"<li>{html.escape(p).replace('\\n', '<br>')}</li>")
#                 parts.append("</ul>")
#     parts.append("</body></html>")
#     return "\n".join(parts)

# # ==================== 入力の正規化・分割 ====================
# def _normalize_items(path_lst: Iterable[str | Path | Tuple[str, str]]) -> List[Path]:
#     out: List[Path] = []
#     for item in path_lst:
#         if isinstance(item, (tuple, list)):
#             p = Path(str(item[0])) / str(item[1]) if len(item) >= 2 else Path(str(item[0]))
#         else:
#             p = Path(str(item))
#         out.append(p)
#     return out

# def _prefilter_items(paths: List[Path], max_kb: int = MAX_FILE_KB_DEFAULT) -> List[Path]:
#     ok: List[Path] = []
#     for p in paths:
#         if not p.exists(): continue
#         ext = p.suffix.lower()
#         if ext not in ALLOWED_EXTS: continue
#         try:
#             size = p.stat().st_size
#             if size > max_kb * 1024:
#                 print(f"[SKIP size>{max_kb}KB] {p} ({size/1024:.1f} KB)")
#                 continue
#         except Exception as e:
#             print(f"[SKIP stat error] {p} / {e}"); continue
#         ok.append(p)
#     return ok

# def _chunk(lst: List[Path], n: int) -> List[List[Path]]:
#     if n <= 0: return [lst[:]]
#     return [lst[i:i+n] for i in range(0, len(lst), n)]

# # ==================== COM 一時障害検知 ====================
# def _is_transient_com_error(exc: Exception) -> bool:
#     """RPC_E_DISCONNECTED(-2147417848) / RPC_S_SERVER_UNAVAILABLE(-2147023174) など"""
#     CODES = {-2147417848, 0x80010108, -2147023174, 0x800706BA}
#     try:
#         hr = getattr(exc, "hresult", None)
#         if hr is None and getattr(exc, "args", None):
#             hr = exc.args[0]
#         if isinstance(hr, int) and hr in CODES:
#             return True
#         msg = (str(exc) or "").upper()
#         return ("80010108" in msg or "RPC_E_DISCONNECTED" in msg or
#                 "800706BA" in msg or "RPC サーバーを利用できません" in msg)
#     except Exception:
#         return False

# # ==================== アプリマネージャ & 安全抽出 ====================
# class _WordManager:
#     def __init__(self): self.app = None
#     def ensure(self):
#         if self.app is None: self.app = _get_word_app()
#         return self.app
#     def reset(self):
#         try:
#             if self.app is not None and not DEFER_QUIT_TO_PARENT:
#                 self.app.Quit()
#         except Exception: pass
#         self.app = None
#         time.sleep(0.1)
#     def close(self):
#         if DEFER_QUIT_TO_PARENT: return  # 親がまとめて Kill
#         self.reset()

# class _PptManager:
#     def __init__(self): self.app = None
#     def ensure(self):
#         if self.app is None: self.app = _get_ppt_app()
#         return self.app
#     def reset(self):
#         try:
#             if self.app is not None and not DEFER_QUIT_TO_PARENT:
#                 self.app.Quit()
#         except Exception: pass
#         self.app = None
#         time.sleep(0.1)
#     def close(self):
#         if DEFER_QUIT_TO_PARENT: return
#         self.reset()

# def _extract_word_safe(src: Path, mgr: _WordManager, retries: int = DEFAULT_RETRIES) -> List[str]:
#     delay = 0.5
#     for attempt in range(retries + 1):
#         try:
#             return extract_paragraphs_from_word(src, mgr.ensure())
#         except Exception as e:
#             if _is_transient_com_error(e) and attempt < retries:
#                 print(f"[WORD reconnect] {src.name} / retry {attempt+1}")
#                 mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
#             raise

# def _extract_ppt_safe(src: Path, mgr: _PptManager, retries: int = DEFAULT_RETRIES) -> Dict[int, List[str]]:
#     delay = 0.5
#     for attempt in range(retries + 1):
#         try:
#             return extract_paragraphs_from_ppt_grouped(src, mgr.ensure())
#         except Exception as e:
#             if _is_transient_com_error(e) and attempt < retries:
#                 print(f"[PPT reconnect] {src.name} / retry {attempt+1}")
#                 mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
#             raise

# # ==================== ワーカー（プロセス） ====================
# def _worker_make_part(part_index: int, paths: List[str], out_dir: str, out_stem: str) -> tuple[int, str, int]:
#     pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
#     # 起動ジッタ：同時 COM アクティベーション衝突の回避
#     time.sleep(0.3 + (part_index % 4) * 0.2 + random.uniform(0.0, 0.1))

#     word_mgr = _WordManager()
#     ppt_mgr  = _PptManager()
#     grouped: Dict[str, Dict] = {}

#     try:
#         total = len(paths)
#         for i, s in enumerate(paths, 1):
#             src = Path(s)
#             if (not src.exists()) or (src.suffix.lower() not in ALLOWED_EXTS):
#                 continue
#             print(f"[part{part_index}] {i}/{total}: {src.name}")
#             try:
#                 if src.suffix.lower() in WORD_EXTS:
#                     paras = _extract_word_safe(src, word_mgr, retries=DEFAULT_RETRIES)
#                     if paras: grouped[src.name] = {"type": "word", "paras": paras}
#                 else:
#                     slides = _extract_ppt_safe(src, ppt_mgr, retries=DEFAULT_RETRIES)
#                     if any(slides.values()): grouped[src.name] = {"type": "ppt", "slides": slides}
#             except Exception as e:
#                 print(f"[part{part_index} SKIP] {src} / {e}")

#         out_dir_p = Path(out_dir); out_dir_p.mkdir(parents=True, exist_ok=True)
#         out_path = out_dir_p / f"{out_stem}_part{part_index}.html"
#         html_text = paragraphs_to_html(grouped) if grouped else \
#             "<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし）</p></body>"
#         out_path.write_text(html_text, encoding="utf-8")
#         return (part_index, str(out_path), len(grouped))

#     finally:
#         try: word_mgr.close()
#         except Exception: pass
#         try: ppt_mgr.close()
#         except Exception: pass
#         try: pythoncom.CoUninitialize()
#         except Exception: pass

# # ==================== 並列メイン API ====================
# def convert_office_to_html(
#     path_lst: Iterable[str | Path | Tuple[str, str]],
#     output_html_path: str | Path,
#     html_dir: str,
#     batch_size: int = BATCH_SIZE_DEFAULT,
#     size_kb_limit: int = MAX_FILE_KB_DEFAULT,
#     max_agents: int | None = None,
# ) -> List[Path]:
#     # 起動前 Kill（全員分の開始前に一度だけ）
#     if KILL_AT_START:
#         kill_office_processes()

#     items = _normalize_items(path_lst)
#     filtered = _prefilter_items(items, max_kb=size_kb_limit)

#     out_base = Path(output_html_path) / html_dir
#     out_dir  = out_base.parent
#     out_stem = out_base.stem

#     if not filtered:
#         empty = out_dir / f"{out_stem}_part1.html"
#         empty.parent.mkdir(parents=True, exist_ok=True)
#         empty.write_text("<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし：対象0件）</p></body>", encoding="utf-8")
#         print(f"[WRITE empty] {empty}")
#         # 終了後 Kill（要求に応じて）
#         if KILL_AT_END: kill_office_processes()
#         return [empty]

#     chunks = _chunk(filtered, batch_size)
#     total_parts = len(chunks)
#     workers = total_parts if max_agents is None else max(1, min(total_parts, max_agents))

#     print(f"[PLAN] 有効 {len(filtered)} 件 → {total_parts} part / 同時実行 {workers} / {batch_size}件/part")
#     for i, ch in enumerate(chunks, 1):
#         print(f"  - part{i}: {len(ch)} 件")

#     tasks = [(idx + 1, [str(p) for p in chunk], str(out_dir), out_stem)
#              for idx, chunk in enumerate(chunks)]

#     results = []
#     with ProcessPoolExecutor(max_workers=workers) as ex:
#         futs = [ex.submit(_worker_make_part, *t) for t in tasks]
#         for f in as_completed(futs):
#             results.append(f.result())

#     # 全員終了 → ここで一括 Kill（≒「誰かが作業中なら閉じない」）
#     if DEFER_QUIT_TO_PARENT and KILL_AT_END:
#         kill_office_processes()

#     results.sort(key=lambda x: x[0])
#     for idx, path_str, cnt in results:
#         print(f"[DONE part{idx}] {path_str}  (収録 {cnt} ファイル)")

#     return [Path(p) for _, p, _ in results]

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




# # -*- coding: utf-8 -*-
# import os, re, html, time, random, subprocess, shutil, uuid
# from pathlib import Path
# from typing import Iterable, List, Dict, Tuple

# import pythoncom
# from win32com.client import gencache, DispatchEx
# from concurrent.futures import ProcessPoolExecutor, as_completed

# # ==================== 調整フラグ ====================
# BATCH_SIZE_DEFAULT      = 5         # 1 HTML の最大収録件数
# MAX_FILE_KB_DEFAULT     = 1500      # これ超は最初からスキップ
# DEFAULT_RETRIES         = 1         # COM切断時の再試行回数（各ファイル）
# EXCLUSIVE_INSTANCE      = True      # True: DispatchEx で専用インスタンス化
# DEFER_QUIT_TO_PARENT    = True      # True: ワーカーでは Quit せず、最後に親が一括 Kill
# KILL_AT_START           = True      # True: 開始前に既存 Office を kill
# KILL_AT_END             = True      # True: 全ワーカー終了後に Office を kill

# # ==================== 定数/拡張子 ====================
# WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
# PPT_EXTS  = {".ppt", ".pptx", ".pptm"}
# ALLOWED_EXTS = WORD_EXTS | PPT_EXTS

# msoTrue  = -1
# msoFalse = 0

# # Word SaveAs constants
# wdFormatFilteredHTML = 10
# wdEncodingUTF8       = 65001

# # ==================== プロセス一括 Kill ====================
# def kill_office_processes() -> None:
#     for image_name in ("WINWORD.EXE", "POWERPNT.EXE"):
#         try:
#             subprocess.run(
#                 ["taskkill", "/IM", image_name, "/F", "/T"],
#                 capture_output=True, text=True, check=False
#             )
#         except Exception:
#             pass

# # ==================== ユーティリティ ====================
# def _clean_paragraph_text(s: str) -> str:
#     if s is None: return ""
#     s = s.replace("\r", "").replace("\x07", "")
#     return s.strip()

# def _ensure_paragraphs_list(items: Iterable[str]) -> List[str]:
#     out: List[str] = []
#     for t in items:
#         t = _clean_paragraph_text(t)
#         if t: out.append(t)
#     return out

# _para_splitter = re.compile(r"[\r\n]+")
# def _split_paragraphs_fast(s: str) -> List[str]:
#     if not s: return []
#     return [ _clean_paragraph_text(x) for x in _para_splitter.split(s) if _clean_paragraph_text(x) ]

# def _safe_rel(child: Path, base: Path) -> str:
#     try:
#         return child.relative_to(base).as_posix()
#     except Exception:
#         return os.path.relpath(child, base).replace("\\","/")

# # ==================== COM アプリ生成 ====================
# def _get_word_app():
#     """
#     EXCLUSIVE_INSTANCE=True の場合は DispatchEx で専用インスタンスを起動。
#     makepy は gencache.EnsureDispatch を一度呼んで生成済みでOK（速度向上）。
#     """
#     try:
#         gencache.EnsureDispatch("Word.Application")
#     except Exception:
#         pass
#     app = DispatchEx("Word.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("Word.Application")
#     app.Visible = False
#     try: app.ScreenUpdating = False
#     except Exception: pass
#     app.DisplayAlerts = 0
#     try:
#         app.Options.CheckGrammarAsYouType = False
#         app.Options.CheckSpellingAsYouType = False
#         app.Options.AllowReadingMode = False
#     except Exception:
#         pass
#     return app

# def _get_ppt_app():
#     try:
#         gencache.EnsureDispatch("PowerPoint.Application")
#     except Exception:
#         pass
#     app = DispatchEx("PowerPoint.Application") if EXCLUSIVE_INSTANCE else gencache.EnsureDispatch("PowerPoint.Application")
#     try: app.Visible = msoFalse
#     except Exception: pass
#     try:
#         from win32com.client import constants as c
#         app.AutomationSecurity = getattr(c, "msoAutomationSecurityForceDisable", 3)
#     except Exception:
#         pass
#     return app

# # ==================== PowerPoint: スライド画像 + テキスト抽出 ====================
# def _iter_shape_texts_fast(shape) -> Iterable[str]:
#     # グループ
#     try:
#         gi = shape.GroupItems
#         for i in range(1, gi.Count + 1):
#             yield from _iter_shape_texts_fast(gi.Item(i))
#         return
#     except Exception:
#         pass
#     # 表
#     try:
#         if int(getattr(shape, "HasTable", 0)) == msoTrue:
#             table = shape.Table
#             for r in range(1, table.Rows.Count + 1):
#                 for c in range(1, table.Columns.Count + 1):
#                     cell_shape = table.Cell(r, c).Shape
#                     try:
#                         tf = cell_shape.TextFrame
#                         if int(getattr(tf, "HasText", 0)) == msoTrue:
#                             yield tf.TextRange.Text; continue
#                     except Exception: pass
#                     try:
#                         tf2 = cell_shape.TextFrame2
#                         if int(getattr(tf2, "HasText", 0)) == msoTrue:
#                             yield tf2.TextRange.Text
#                     except Exception: pass
#             return
#     except Exception:
#         pass
#     # 通常テキスト
#     try:
#         if int(getattr(shape, "HasTextFrame", 0)) == msoTrue:
#             tf = shape.TextFrame
#             if int(getattr(tf, "HasText", 0)) == msoTrue:
#                 yield tf.TextRange.Text; return
#     except Exception:
#         pass
#     # 予備
#     try:
#         tf2 = shape.TextFrame2
#         if int(getattr(tf2, "HasText", 0)) == msoTrue:
#             yield tf2.TextRange.Text
#     except Exception:
#         pass

# def extract_ppt_with_images(src: Path, ppt_app, assets_base_dir: Path, html_base_dir: Path, target_width_px: int = 1600) -> Dict[int, Dict[str, object]]:
#     """
#     各スライドを PNG 画像に書き出し、テキストも抽出。
#     戻り値: { slide_no: {"img_rel": str, "paras": List[str]} }
#     """
#     created_here = False
#     if ppt_app is None:
#         ppt_app = _get_ppt_app(); created_here = True

#     pres = None
#     result: Dict[int, Dict[str, object]] = {}
#     try:
#         pres = ppt_app.Presentations.Open(FileName=str(src), WithWindow=False, ReadOnly=True)

#         # スライド画像出力サイズ計算
#         ps = pres.PageSetup
#         sw, sh = float(ps.SlideWidth), float(ps.SlideHeight)
#         width_px  = int(target_width_px)
#         height_px = max(1, int(round(width_px * (sh / sw))))

#         img_dir = assets_base_dir / "ppt" / src.stem
#         img_dir.mkdir(parents=True, exist_ok=True)

#         for s_idx in range(1, pres.Slides.Count + 1):
#             slide = pres.Slides(s_idx)
#             img_name = f"{src.stem}_slide{s_idx:03d}.png"
#             img_path = img_dir / img_name
#             # 画像としてエクスポート
#             slide.Export(str(img_path), "PNG", width_px, height_px)

#             # テキスト抽出（従来ロジック）
#             buf: List[str] = []
#             shapes = slide.Shapes
#             for j in range(1, shapes.Count + 1):
#                 for raw in _iter_shape_texts_fast(shapes(j)):
#                     if raw: buf.extend(_split_paragraphs_fast(raw))
#             paras = _ensure_paragraphs_list(buf)

#             result[s_idx] = {
#                 "img_rel": _safe_rel(img_path, html_base_dir),
#                 "paras": paras
#             }
#     finally:
#         if pres is not None:
#             pres.Close()
#         if created_here and not DEFER_QUIT_TO_PARENT and ppt_app is not None:
#             try: ppt_app.Quit()
#             except Exception: pass
#     return result

# # ==================== Word: 画像を“その位置”で埋め込む（Filtered HTML 利用） ====================
# def extract_word_html_with_images(src: Path, word_app, assets_base_dir: Path, html_base_dir: Path) -> str:
#     """
#     Word の SaveAs2(Filterd HTML)で、画像を含む位置忠実な HTML を一時生成。
#     画像を assets/word/<stem>/ にコピーし、HTML 内の <img src> を相対パスへ張り替え。
#     HTML の <body> 内だけを切り出して返す。
#     """
#     created_here = False
#     if word_app is None:
#         word_app = _get_word_app(); created_here = True

#     # 作業用一時ディレクトリ（ワーカー内でユニーク）
#     tmp_dir = assets_base_dir / f"~wordtmp_{os.getpid()}_{uuid.uuid4().hex[:8]}"
#     tmp_dir.mkdir(parents=True, exist_ok=True)
#     tmp_html = tmp_dir / f"{src.stem}.htm"

#     dest_img_dir = assets_base_dir / "word" / src.stem
#     dest_img_dir.mkdir(parents=True, exist_ok=True)

#     doc = None
#     try:
#         doc = word_app.Documents.Open(
#             FileName=str(src),
#             ReadOnly=True, AddToRecentFiles=False,
#             OpenAndRepair=False, ConfirmConversions=False, NoEncodingDialog=True
#         )
#         # Filtered HTML で出力（UTF-8 指定）
#         # SaveAs2 の引数: FileName, FileFormat, LockComments, Password, AddToRecentFiles,
#         # WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat,
#         # SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks
#         doc.SaveAs2(
#             FileName=str(tmp_html),
#             FileFormat=wdFormatFilteredHTML,
#             Encoding=wdEncodingUTF8
#         )
#     finally:
#         if doc is not None:
#             doc.Close(False)
#         if created_here and not DEFER_QUIT_TO_PARENT and word_app is not None:
#             try: word_app.Quit()
#             except Exception: pass

#     # 画像フォルダ（Word が生成）
#     images_src_dir = tmp_dir / f"{src.stem}_files"
#     # HTML 読み込み
#     html_text = ""
#     if tmp_html.exists():
#         # Word は BOM 付き UTF-8 を出すことが多いので errors=ignore で読む
#         html_text = tmp_html.read_text(encoding="utf-8", errors="ignore")
#     else:
#         # まれに .html という拡張子で出すケースに備える
#         alt = tmp_dir / f"{src.stem}.html"
#         if alt.exists():
#             html_text = alt.read_text(encoding="utf-8", errors="ignore")

#     # 画像コピー & src の相対パス差し替え
#     if images_src_dir.exists():
#         for f in images_src_dir.iterdir():
#             if f.is_file():
#                 shutil.copy2(f, dest_img_dir / f.name)
#         # src=".../<stem>_files/xxx" → src="assets/word/<stem>/xxx"（相対）
#         dest_rel = _safe_rel(dest_img_dir, html_base_dir).rstrip("/")
#         html_text = re.sub(
#             rf'(?i)src="[^"]*{re.escape(src.stem)}_files/',
#             f'src="{dest_rel}/',
#             html_text
#         )

#     # <body> 内だけ抽出
#     m = re.search(r"(?is)<body[^>]*>(.*)</body>", html_text)
#     body_inner = m.group(1) if m else html_text
#     # 不要な条件コメントなど軽く除去
#     body_inner = re.sub(r"<!--\[if.*?endif\]-->", "", body_inner, flags=re.I|re.S)

#     # 一時ディレクトリ掃除
#     try:
#         shutil.rmtree(tmp_dir, ignore_errors=True)
#     except Exception:
#         pass

#     return body_inner

# # ==================== HTML 生成 ====================
# def paragraphs_to_html(grouped: Dict[str, Dict], html_base_dir: Path) -> str:
#     """
#     grouped の構造:
#       - Word: {"type":"word", "html": "<body内html>..."}
#       - PPT : {"type":"ppt",  "slides": {slide_no: {"img_rel": str, "paras": [str,...]}}}
#       - （従来のプレーンテキスト Word の場合は "paras": [...] を許容）
#     """
#     parts: List[str] = []
#     parts.append("<!DOCTYPE html>")
#     parts.append('<html lang="ja">')
#     parts.append('<meta charset="UTF-8">')
#     parts.append("<title>抽出テキスト</title>")
#     parts.append("<style>img{max-width:100%;height:auto;display:block;margin:.4rem 0;} ul{margin:.3rem 1.2rem;} h2{margin-top:1.2rem;} .word-block{border:1px solid #ddd;padding:.6rem;border-radius:.5rem;} .slide-block{border:1px solid #eee;padding:.6rem;border-radius:.5rem;}</style>")
#     parts.append("<body>")

#     for file_disp, data in grouped.items():
#         parts.append(f"<h2>{html.escape(file_disp)}</h2>")
#         if data.get("type") == "ppt":
#             slides = data.get("slides", {})
#             for slide_no in sorted(slides.keys()):
#                 entry = slides[slide_no]
#                 img_rel = entry.get("img_rel")
#                 paras   = entry.get("paras", [])
#                 parts.append(f"<div class='slide-block'><h3>Slide {slide_no}</h3>")
#                 if img_rel:
#                     parts.append(f"<img src='{html.escape(img_rel)}' alt='slide {slide_no}'>")
#                 if paras:
#                     parts.append("<ul>")
#                     for p in paras:
#                         parts.append(f"<li>{html.escape(p).replace('\\n','<br>')}</li>")
#                     parts.append("</ul>")
#                 parts.append("</div>")
#         elif "html" in data:
#             parts.append("<div class='word-block'>")
#             parts.append(data["html"])  # Word の body 内 HTML（画像含む）
#             parts.append("</div>")
#         else:
#             # 互換: 旧 Word テキストのみ
#             paras = data.get("paras", [])
#             if paras:
#                 parts.append("<ul>")
#                 for p in paras:
#                     parts.append(f"<li>{html.escape(p).replace('\\n', '<br>')}</li>")
#                 parts.append("</ul>")
#     parts.append("</body></html>")
#     return "\n".join(parts)

# # ==================== 入力の正規化・分割 ====================
# def _normalize_items(path_lst: Iterable[str | Path | Tuple[str, str]]) -> List[Path]:
#     out: List[Path] = []
#     for item in path_lst:
#         if isinstance(item, (tuple, list)):
#             p = Path(str(item[0])) / str(item[1]) if len(item) >= 2 else Path(str(item[0]))
#         else:
#             p = Path(str(item))
#         out.append(p)
#     return out

# def _prefilter_items(paths: List[Path], max_kb: int = MAX_FILE_KB_DEFAULT) -> List[Path]:
#     ok: List[Path] = []
#     for p in paths:
#         if not p.exists(): continue
#         ext = p.suffix.lower()
#         if ext not in ALLOWED_EXTS: continue
#         try:
#             size = p.stat().st_size
#             if size > max_kb * 1024:
#                 print(f"[SKIP size>{max_kb}KB] {p} ({size/1024:.1f} KB)")
#                 continue
#         except Exception as e:
#             print(f"[SKIP stat error] {p} / {e}"); continue
#         ok.append(p)
#     return ok

# def _chunk(lst: List[Path], n: int) -> List[List[Path]]:
#     if n <= 0: return [lst[:]]
#     return [lst[i:i+n] for i in range(0, len(lst), n)]

# # ==================== COM 一時障害検知 ====================
# def _is_transient_com_error(exc: Exception) -> bool:
#     """RPC_E_DISCONNECTED(-2147417848) / RPC_S_SERVER_UNAVAILABLE(-2147023174) など"""
#     CODES = {-2147417848, 0x80010108, -2147023174, 0x800706BA}
#     try:
#         hr = getattr(exc, "hresult", None)
#         if hr is None and getattr(exc, "args", None):
#             hr = exc.args[0]
#         if isinstance(hr, int) and hr in CODES:
#             return True
#         msg = (str(exc) or "").upper()
#         return ("80010108" in msg or "RPC_E_DISCONNECTED" in msg or
#                 "800706BA" in msg or "RPC サーバーを利用できません" in msg)
#     except Exception:
#         return False

# # ==================== アプリマネージャ & 安全抽出 ====================
# class _WordManager:
#     def __init__(self): self.app = None
#     def ensure(self):
#         if self.app is None: self.app = _get_word_app()
#         return self.app
#     def reset(self):
#         try:
#             if self.app is not None and not DEFER_QUIT_TO_PARENT:
#                 self.app.Quit()
#         except Exception: pass
#         self.app = None
#         time.sleep(0.1)
#     def close(self):
#         if DEFER_QUIT_TO_PARENT: return
#         self.reset()

# class _PptManager:
#     def __init__(self): self.app = None
#     def ensure(self):
#         if self.app is None: self.app = _get_ppt_app()
#         return self.app
#     def reset(self):
#         try:
#             if self.app is not None and not DEFER_QUIT_TO_PARENT:
#                 self.app.Quit()
#         except Exception: pass
#         self.app = None
#         time.sleep(0.1)
#     def close(self):
#         if DEFER_QUIT_TO_PARENT: return
#         self.reset()

# def _extract_word_safe_html(src: Path, mgr: _WordManager, assets_base_dir: Path, html_base_dir: Path, retries: int = DEFAULT_RETRIES) -> str:
#     delay = 0.5
#     for attempt in range(retries + 1):
#         try:
#             return extract_word_html_with_images(src, mgr.ensure(), assets_base_dir, html_base_dir)
#         except Exception as e:
#             if _is_transient_com_error(e) and attempt < retries:
#                 print(f"[WORD reconnect] {src.name} / retry {attempt+1}")
#                 mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
#             raise

# def _extract_ppt_safe_images(src: Path, mgr: _PptManager, assets_base_dir: Path, html_base_dir: Path, retries: int = DEFAULT_RETRIES) -> Dict[int, Dict[str, object]]:
#     delay = 0.5
#     for attempt in range(retries + 1):
#         try:
#             return extract_ppt_with_images(src, mgr.ensure(), assets_base_dir, html_base_dir)
#         except Exception as e:
#             if _is_transient_com_error(e) and attempt < retries:
#                 print(f"[PPT reconnect] {src.name} / retry {attempt+1}")
#                 mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
#             raise

# # ==================== ワーカー（プロセス） ====================
# def _worker_make_part(part_index: int, paths: List[str], out_dir: str, out_stem: str) -> tuple[int, str, int]:
#     pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
#     time.sleep(0.3 + (part_index % 4) * 0.2 + random.uniform(0.0, 0.1))  # 同時 COM アクティベーション緩和

#     word_mgr = _WordManager()
#     ppt_mgr  = _PptManager()
#     grouped: Dict[str, Dict] = {}

#     out_dir_p = Path(out_dir)
#     assets_base_dir = out_dir_p / "assets"  # 画像保存ルート
#     assets_base_dir.mkdir(parents=True, exist_ok=True)

#     try:
#         total = len(paths)
#         for i, s in enumerate(paths, 1):
#             src = Path(s)
#             if (not src.exists()) or (src.suffix.lower() not in ALLOWED_EXTS):
#                 continue
#             print(f"[part{part_index}] {i}/{total}: {src.name}")
#             try:
#                 if src.suffix.lower() in WORD_EXTS:
#                     # 画像位置を保持した HTML 断片
#                     html_fragment = _extract_word_safe_html(src, word_mgr, assets_base_dir, out_dir_p)
#                     if html_fragment:
#                         grouped[src.name] = {"type": "word", "html": html_fragment}
#                 else:
#                     # スライド画像 + テキスト
#                     slides = _extract_ppt_safe_images(src, ppt_mgr, assets_base_dir, out_dir_p)
#                     if slides:
#                         grouped[src.name] = {"type": "ppt", "slides": slides}
#             except Exception as e:
#                 print(f"[part{part_index} SKIP] {src} / {e}")

#         out_dir_p.mkdir(parents=True, exist_ok=True)
#         out_path = out_dir_p / f"{out_stem}_part{part_index}.html"
#         html_text = paragraphs_to_html(grouped, out_dir_p) if grouped else \
#             "<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし）</p></body>"
#         out_path.write_text(html_text, encoding="utf-8")
#         return (part_index, str(out_path), len(grouped))

#     finally:
#         try: word_mgr.close()
#         except Exception: pass
#         try: ppt_mgr.close()
#         except Exception: pass
#         try: pythoncom.CoUninitialize()
#         except Exception: pass

# # ==================== 並列メイン API ====================
# def convert_office_to_html(
#     path_lst: Iterable[str | Path | Tuple[str, str]],
#     output_html_path: str | Path,
#     html_dir: str,
#     batch_size: int = BATCH_SIZE_DEFAULT,
#     size_kb_limit: int = MAX_FILE_KB_DEFAULT,
#     max_agents: int | None = None,
# ) -> List[Path]:
#     # 起動前 Kill
#     if KILL_AT_START:
#         kill_office_processes()

#     items = _normalize_items(path_lst)
#     filtered = _prefilter_items(items, max_kb=size_kb_limit)

#     out_base = Path(output_html_path) / html_dir     # 例: output_html_path="C:/out", html_dir="a.html"
#     out_dir  = out_base.parent                       # → C:/out
#     out_stem = out_base.stem                         # → a

#     if not filtered:
#         empty = out_dir / f"{out_stem}_part1.html"
#         empty.parent.mkdir(parents=True, exist_ok=True)
#         empty.write_text("<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし：対象0件）</p></body>", encoding="utf-8")
#         print(f"[WRITE empty] {empty}")
#         if KILL_AT_END: kill_office_processes()
#         return [empty]

#     chunks = _chunk(filtered, batch_size)
#     total_parts = len(chunks)
#     workers = total_parts if max_agents is None else max(1, min(total_parts, max_agents))

#     print(f"[PLAN] 有効 {len(filtered)} 件 → {total_parts} part / 同時実行 {workers} / {batch_size}件/part")
#     for i, ch in enumerate(chunks, 1):
#         print(f"  - part{i}: {len(ch)} 件")

#     tasks = [(idx + 1, [str(p) for p in chunk], str(out_dir), out_stem)
#              for idx, chunk in enumerate(chunks)]

#     results = []
#     with ProcessPoolExecutor(max_workers=workers) as ex:
#         futs = [ex.submit(_worker_make_part, *t) for t in tasks]
#         for f in as_completed(futs):
#             results.append(f.result())

#     if DEFER_QUIT_TO_PARENT and KILL_AT_END:
#         kill_office_processes()

#     results.sort(key=lambda x: x[0])
#     for idx, path_str, cnt in results:
#         print(f"[DONE part{idx}] {path_str}  (収録 {cnt} ファイル)")

#     return [Path(p) for _, p, _ in results]

# # ==================== 直接実行テスト ====================
# if __name__ == "__main__":
#     # (dir, file) タプル配列で渡す
#     path_list = [
#         (r"C:\Users\yohei\Downloads\R2-2206906", "R2-2206906_C1-224008.docx"),
#         (r"C:\Users\yohei\Downloads\R2-2206905", "R2-2206905_C1-223972.docx"),
#         (r"C:\Users\yohei\Downloads",           "11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx"),
#     ]
#     out_dir  = r"C:\Users\yohei\Downloads\html_out"
#     out_file = "a.html"  # 実際には a_part1.html など分割生成

#     outputs = convert_office_to_html(
#         path_lst=path_list,
#         output_html_path=out_dir,
#         html_dir=out_file,
#         batch_size=5,
#         size_kb_limit=1500,
#         # max_agents=4,  # PPT が多い場合は 3〜4 程度を推奨
#     )
#     print("✅ 出力一覧:")
#     for p in outputs:
#         print(" -", p)


# -*- coding: utf-8 -*-
"""
Office(Word/PPT) → HTML 結合器（最終統合／高速化版）
- Word: 画像を“その位置”のまま抽出（Filtered HTML→画像総当たり回収→相対化）
- PPT : 各スライドを PNG 化して <img> 埋め込み（テキスト抽出はトグル）
"""

import os, re, html, time, random, subprocess, shutil, uuid, hashlib
from pathlib import Path
from typing import Iterable, List, Dict, Tuple, Optional

import pythoncom
from win32com.client import gencache, DispatchEx
from concurrent.futures import ProcessPoolExecutor, as_completed
from urllib.parse import urlparse, unquote

# ==================== 調整フラグ ====================
BATCH_SIZE_DEFAULT      = 5          # 1 HTML の最大収録件数
MAX_FILE_KB_DEFAULT     = 1500       # 閾値超はスキップ
DEFAULT_RETRIES         = 1          # COM切断時の再試行回数（各ファイル）
EXCLUSIVE_INSTANCE      = True       # True: DispatchEx で専用インスタンス化
DEFER_QUIT_TO_PARENT    = True       # True: ワーカーで Quit しない（最後に親が Kill）
KILL_AT_START           = True       # 実行開始前に既存 Office Kill
KILL_AT_END             = True       # 実行終了後に Office Kill
EXTRACT_PPT_TEXT        = True       # PPTのテキスト抽出（False なら画像のみで最速）
PPT_TARGET_WIDTH_PX     = 1600       # スライド画像の横幅
WORD_MINIFY_INLINE_CSS  = False      # TrueでWordの余計なstyleを粗く間引く（必要なら）

# ==================== 定数/拡張子 ====================
WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
PPT_EXTS  = {".ppt", ".pptx", ".pptm"}
ALLOWED_EXTS = WORD_EXTS | PPT_EXTS

msoTrue  = -1
msoFalse = 0

# Office 定数
wdFormatFilteredHTML = 10
wdEncodingUTF8       = 65001
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
# 画像系の属性 (<img src="..."> / background="..." / v:imagedata src="..." / v:fill src="...")
_IMG_ATTR_PATTERN = re.compile(
    r'''
    # <img>, background 属性
    \b(?:src|background)\s*=\s*(?P<q1>["'])(?P<u1>.*?)(?P=q1)
    |
    # VML 系: v:imagedata / v:fill の src
    \b(?:v:imagedata|v:fill)\b[^>]*\bsrc\s*=\s*(?P<q2>["'])(?P<u2>.*?)(?P=q2)
    ''',
    re.IGNORECASE | re.DOTALL | re.VERBOSE
)


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

def _split_paragraphs_fast(s: str) -> List[str]:
    if not s: return []
    return [ _clean_paragraph_text(x) for x in _para_splitter.split(s) if _clean_paragraph_text(x) ]

def _safe_rel(child: Path, base: Path) -> str:
    try:
        return child.relative_to(base).as_posix()
    except Exception:
        return os.path.relpath(child, base).replace("\\","/")

def _short_hash(s: str) -> str:
    return hashlib.sha1(s.encode("utf-8", errors="ignore")).hexdigest()[:8]

def _fast_copy(src: Path, dst: Path) -> None:
    """
    できればハードリンク（同一ボリューム）→だめなら copyfile（メタデータコピー不要で速い）
    """
    try:
        os.link(src, dst)  # Windows 10+ で管理者不要（同一NTFSボリューム）
        return
    except Exception:
        pass
    shutil.copyfile(src, dst)

# ==================== COM アプリ生成 ====================
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

# ==================== PPT：画像化 +（任意）テキスト抽出 ====================
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

def extract_ppt_with_images(src: Path, ppt_app, assets_base_dir: Path, html_base_dir: Path, target_width_px: int) -> Dict[int, Dict[str, object]]:
    created_here = False
    if ppt_app is None:
        ppt_app = _get_ppt_app(); created_here = True

    pres = None
    result: Dict[int, Dict[str, object]] = {}
    try:
        pres = ppt_app.Presentations.Open(FileName=str(src), WithWindow=False, ReadOnly=True)

        ps = pres.PageSetup
        sw, sh = float(ps.SlideWidth), float(ps.SlideHeight)
        width_px  = int(target_width_px)
        height_px = max(1, int(round(width_px * (sh / sw))))

        img_dir = assets_base_dir / "ppt" / src.stem
        img_dir.mkdir(parents=True, exist_ok=True)

        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            img_name = f"{src.stem}_slide{s_idx:03d}.png"
            img_path = img_dir / img_name
            slide.Export(str(img_path), "PNG", width_px, height_px)

            paras: List[str] = []
            if EXTRACT_PPT_TEXT:
                shapes = slide.Shapes
                for j in range(1, shapes.Count + 1):
                    for raw in _iter_shape_texts_fast(shapes(j)):
                        if raw: paras.extend(_split_paragraphs_fast(raw))
                paras = _ensure_paragraphs_list(paras)

            result[s_idx] = {
                "img_rel": _safe_rel(img_path, html_base_dir),
                "paras": paras
            }
    finally:
        if pres is not None:
            pres.Close()
        if created_here and not DEFER_QUIT_TO_PARENT and ppt_app is not None:
            try: ppt_app.Quit()
            except Exception: pass
    return result

# ==================== Word：画像を“その位置”で埋め込む ====================
def _word_set_weboptions(doc) -> None:
    try:
        w = doc.WebOptions
        w.OrganizeInFolder   = True
        w.UseLongFileNames   = True
        w.RelyOnCSS          = True
        w.RelyOnVML          = False   # <img> 優先
        w.OptimizeForBrowser = True
    except Exception:
        pass

def _resolve_candidate_path(u: str, tmp_dir: Path) -> Optional[Path]:
    if not u:
        return None
    s = u.strip().strip('"').strip("'")

    if s.lower().startswith("data:"):
        return None

    if s.lower().startswith("file:"):
        p = urlparse(s)
        candidate = unquote(p.path or "")
        if candidate.startswith("/") and len(candidate) > 3 and candidate[2] == ":":
            candidate = candidate[1:]
        candidate = candidate.replace("/", "\\")
        q = Path(candidate)
        return q if q.exists() else None

    s2 = unquote(s).replace("\\", "/")
    q = (tmp_dir / s2).resolve()
    if q.exists():
        return q
    q2 = (tmp_dir / Path(s2).name).resolve()
    if q2.exists():
        return q2
    return None

def _rewrite_and_copy_all_image_srcs(html_text: str, tmp_dir: Path, dest_img_dir: Path, html_base_dir: Path) -> tuple[str, int]:
    dest_img_dir.mkdir(parents=True, exist_ok=True)
    dest_rel = _safe_rel(dest_img_dir, html_base_dir).rstrip("/")

    copied = 0
    repl_map: Dict[str, str] = {}

    def do_copy(src_path: Path) -> str:
        nonlocal copied
        name = src_path.name
        dst = dest_img_dir / name
        if dst.exists():
            stem, ext = dst.stem, dst.suffix
            k = 1
            while True:
                cand = dest_img_dir / f"{stem}_{k}{ext}"
                if not cand.exists():
                    dst = cand
                    break
                k += 1
        _fast_copy(src_path, dst)
        copied += 1
        return f"{dest_rel}/{dst.name}"

    def _replacer(m: re.Match) -> str:
        whole = m.group(0)
        u = m.group("u1") or m.group("u2")
        if not u:
            return whole
        if u in repl_map:
            return whole.replace(u, repl_map[u])

        p = _resolve_candidate_path(u, tmp_dir)
        if p and p.exists():
            new_url = do_copy(p)
            repl_map[u] = new_url
            return whole.replace(u, new_url)
        return whole

    new_html = _IMG_ATTR_PATTERN.sub(_replacer, html_text)
    return new_html, copied

def _minify_word_inline_styles(body_inner: str) -> str:
    # ざっくりした軽量化（不要ならOFF）
    body_inner = re.sub(r'(?is)\s*class="Mso[^"]*"', "", body_inner)
    body_inner = re.sub(r'(?is)\s*style="[^"]*"', lambda m: ' style="margin:0;padding:0;"', body_inner)
    return body_inner

def extract_word_html_with_images(src: Path, word_app, assets_base_dir: Path, html_base_dir: Path, asset_label: str) -> str:
    created_here = False
    if word_app is None:
        word_app = _get_word_app(); created_here = True

    tmp_dir = assets_base_dir / f"~wordtmp_{os.getpid()}_{uuid.uuid4().hex[:8]}"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_html = tmp_dir / f"{src.stem}.htm"

    # ラベルで衝突回避（同stemでも別フォルダに）
    dest_img_dir = assets_base_dir / "word" / asset_label
    dest_img_dir.mkdir(parents=True, exist_ok=True)

    doc = None
    try:
        doc = word_app.Documents.Open(
            FileName=str(src),
            ReadOnly=True, AddToRecentFiles=False,
            OpenAndRepair=False, ConfirmConversions=False, NoEncodingDialog=True,
            Visible=False
        )
        _word_set_weboptions(doc)
        doc.SaveAs2(FileName=str(tmp_html), FileFormat=wdFormatFilteredHTML, Encoding=wdEncodingUTF8)
    finally:
        if doc is not None:
            doc.Close(False)
        if created_here and not DEFER_QUIT_TO_PARENT and word_app is not None:
            try: word_app.Quit()
            except Exception: pass

    html_text = ""
    for cand in (tmp_html, tmp_dir / f"{src.stem}.html"):
        if cand.exists():
            html_text = cand.read_text(encoding="utf-8", errors="ignore")
            break
    if not html_text:
        try: shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception: pass
        return ""

    html_text, copied_cnt = _rewrite_and_copy_all_image_srcs(html_text, tmp_dir, dest_img_dir, html_base_dir)
    if copied_cnt == 0:
        # 予備：*_files 丸ごとコピー + パス置換
        images_src_dir = tmp_dir / f"{src.stem}_files"
        if images_src_dir.exists():
            for f in images_src_dir.iterdir():
                if f.is_file():
                    dst = dest_img_dir / f.name
                    if not dst.exists():
                        try: _fast_copy(f, dst)
                        except Exception: pass
            dest_rel = _safe_rel(dest_img_dir, html_base_dir).rstrip("/")
            html_text = re.sub(rf'(?i){re.escape(src.stem)}_files/', f'{dest_rel}/', html_text)

    m = re.search(r"(?is)<body[^>]*>(.*)</body>", html_text)
    body_inner = m.group(1) if m else html_text

    if WORD_MINIFY_INLINE_CSS:
        body_inner = _minify_word_inline_styles(body_inner)

    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception:
        pass

    return body_inner

# ==================== HTML 生成 ====================
def paragraphs_to_html(entries: List[Tuple[str, Dict]], html_base_dir: Path) -> str:
    """
    entries: [(display_name, data_dict), ...]
      - Word: data_dict={"type":"word","html": "<body内HTML>"}
      - PPT : data_dict={"type":"ppt","slides": {no: {"img_rel":..., "paras":[...]}}}
      - 旧Wordテキスト: data_dict={"paras":[...]}
    """
    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="en">')
    parts.append('<meta charset="UTF-8">')
    parts.append("<title>抽出テキスト</title>")
    parts.append("<style>img{max-width:100%;height:auto;display:block;margin:.4rem 0;} ul{margin:.3rem 1.2rem;} h2{margin-top:1.2rem;} .word-block,.slide-block{border:1px solid #e3e3e3;padding:.6rem;border-radius:.5rem;}</style>")
    parts.append("<body>")

    for file_disp, data in entries:
        parts.append(f"<h2>{html.escape(file_disp)}</h2>")
        if data.get("type") == "ppt":
            slides = data.get("slides", {})
            for slide_no in sorted(slides.keys()):
                entry = slides[slide_no]
                img_rel = entry.get("img_rel")
                paras   = entry.get("paras", [])
                parts.append(f"<div class='slide-block'><h3>Slide {slide_no}</h3>")
                if img_rel:
                    parts.append(f"<img src='{html.escape(img_rel)}' alt='slide {slide_no}'>")
                if paras:
                    parts.append("<ul>")
                    for p in paras:
                        parts.append(f"<li>{html.escape(p).replace('\\n','<br>')}</li>")
                    parts.append("</ul>")
                parts.append("</div>")
        elif "html" in data:
            parts.append("<div class='word-block'>")
            parts.append(data["html"])
            parts.append("</div>")
        else:
            paras = data.get("paras", [])
            if paras:
                parts.append("<ul>")
                for p in paras:
                    parts.append(f"<li>{html.escape(p).replace('\\n','<br>')}</li>")
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

# ==================== セーフ抽出ラッパ ====================
def _extract_word_safe_html(src: Path, mgr: _WordManager, assets_base_dir: Path, html_base_dir: Path, asset_label: str, retries: int = DEFAULT_RETRIES) -> str:
    delay = 0.5
    for attempt in range(retries + 1):
        try:
            return extract_word_html_with_images(src, mgr.ensure(), assets_base_dir, html_base_dir, asset_label)
        except Exception as e:
            if _is_transient_com_error(e) and attempt < retries:
                print(f"[WORD reconnect] {src.name} / retry {attempt+1}")
                mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
            raise

def _extract_ppt_safe_images(src: Path, mgr: _PptManager, assets_base_dir: Path, html_base_dir: Path, retries: int = DEFAULT_RETRIES) -> Dict[int, Dict[str, object]]:
    delay = 0.5
    for attempt in range(retries + 1):
        try:
            return extract_ppt_with_images(src, mgr.ensure(), assets_base_dir, html_base_dir, PPT_TARGET_WIDTH_PX)
        except Exception as e:
            if _is_transient_com_error(e) and attempt < retries:
                print(f"[PPT reconnect] {src.name} / retry {attempt+1}")
                mgr.reset(); time.sleep(delay); delay = min(delay * 2, 3.0); continue
            raise

# ==================== ワーカー（パート単位） ====================
def _worker_make_part(part_index: int, paths: List[str], out_dir: str, out_stem: str) -> tuple[int, str, int]:
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    # 起動ジッタで同時COMアクティベーション衝突を緩和
    time.sleep(0.2 + (part_index % 4) * 0.15 + random.uniform(0.0, 0.05))

    word_mgr = _WordManager()
    ppt_mgr  = _PptManager()
    entries: List[Tuple[str, Dict]] = []

    out_dir_p = Path(out_dir)
    assets_base_dir = out_dir_p / "assets"
    assets_base_dir.mkdir(parents=True, exist_ok=True)

    try:
        total = len(paths)
        for i, s in enumerate(paths, 1):
            src = Path(s)
            if (not src.exists()) or (src.suffix.lower() not in ALLOWED_EXTS):
                continue
            print(f"[part{part_index}] {i}/{total}: {src.name}")
            try:
                if src.suffix.lower() in WORD_EXTS:
                    label = f"{src.stem}_{_short_hash(str(src.resolve()))}"
                    html_fragment = _extract_word_safe_html(src, word_mgr, assets_base_dir, out_dir_p, label)
                    if html_fragment:
                        entries.append( (src.name, {"type": "word", "html": html_fragment}) )
                else:
                    slides = _extract_ppt_safe_images(src, ppt_mgr, assets_base_dir, out_dir_p)
                    if slides:
                        entries.append( (src.name, {"type": "ppt", "slides": slides}) )
            except Exception as e:
                print(f"[part{part_index} SKIP] {src} / {e}")

        out_dir_p.mkdir(parents=True, exist_ok=True)
        out_path = out_dir_p / f"{out_stem}_part{part_index}.html"
        html_text = paragraphs_to_html(entries, out_dir_p) if entries else \
            "<!DOCTYPE html><meta charset='UTF-8'><title>抽出テキスト</title><body><p>（内容なし）</p></body>"
        out_path.write_text(html_text, encoding="utf-8")
        return (part_index, str(out_path), len(entries))

    finally:
        try: word_mgr.close()
        except Exception: pass
        try: ppt_mgr.close()
        except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass

# ==================== メイン API ====================
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
    html_dir: まとめファイル名（例: "result.html" → 実際は *_partN.html）
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

# # ==================== 直接実行テスト ====================
# if __name__ == "__main__":
#     # 例
#     path_list = [
#         (r"C:\Users\yohei\Downloads\R2-2206906", "R2-2206906_C1-224008.docx"),
#         (r"C:\Users\yohei\Downloads\R2-2206905", "R2-2206905_C1-223972.docx"),
#         (r"C:\Users\yohei\Downloads", "11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx"),
#         # r"C:\Users\yohei\Downloads\something.pptx",  # 直接パスでもOK
#     ]
#     out_dir  = r"C:\Users\yohei\Downloads\html_out"
#     out_file = "a.html"  # 実際には a_part1.html など分割生成

#     outputs = convert_office_to_html(
#         path_lst=path_list,
#         output_html_path=out_dir,
#         html_dir=out_file,
#         batch_size=5,
#         size_kb_limit=1500,
#         # max_agents=4,  # PPT 多めなら 3〜4 推奨
#     )
#     print("✅ 出力一覧:")
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
