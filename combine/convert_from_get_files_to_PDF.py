# # -*- coding: utf-8 -*-
# from __future__ import annotations
# from pathlib import Path
# from typing import Iterable, List, Optional, Any, Union
# import win32com.client as win32

# # ===== 拡張子 =====
# WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
# PPT_EXTS  = {".ppt", ".pptx", ".pptm"}

# # ===== Word 定数 =====
# wdExportFormatPDF              = 17
# wdExportOptimizeForPrint       = 0   # 印刷向け
# wdExportRangeAll               = 0
# wdExportDocumentContent        = 0
# wdExportCreateNoBookmarks      = 0

# # MsoTriState
# msoTrue  = -1
# msoFalse = 0

# # PowerPoint 定数
# ppFixedFormatTypePDF     = 2
# ppFixedFormatIntentPrint = 2
# ppPrintAll               = 1
# ppWindowMinimized        = 2

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


# def _read_absolute_paths_pandas(folder: Path) -> List[Path]:
#     """folder/get_files.xlsx のA列から絶対パスを読む（pandas使用）。無ければ空リスト。"""
#     xlsx = folder / "get_files.xlsx"
#     if not xlsx.exists():
#         print(f"ℹ️ get_files.xlsx なし（処理スキップ）: {xlsx}")
#         return []
#     df = pd.read_excel(xlsx, header=None, usecols=[0], names=["path"], dtype=str, engine="openpyxl")
#     paths: List[Path] = []
#     for s in df["path"].dropna():
#         s = str(s).strip().strip('"').strip("'")
#         if not s:
#             continue
#         p = Path(s)
#         if not p.is_absolute():
#             print(f"↪ 相対パスのためスキップ: {s}")
#             continue
#         paths.append(p)
#     return paths

# def _convert_word_to_pdf(app_word, src: Path, dst: Path) -> bool:
#     """Word → PDF（詳細パラメータ版）"""
#     try:
#         doc = app_word.Documents.Open(str(src), ReadOnly=True, Visible=False)
#         # ご指定の引数をそのまま反映
#         doc.ExportAsFixedFormat(
#             OutputFileName=str(dst),
#             ExportFormat=wdExportFormatPDF,
#             OpenAfterExport=False,
#             OptimizeFor=wdExportOptimizeForPrint,
#             Range=wdExportRangeAll,
#             Item=wdExportDocumentContent,
#             IncludeDocProps=True,
#             KeepIRM=True,
#             CreateBookmarks=wdExportCreateNoBookmarks,
#             DocStructureTags=True,
#             BitmapMissingFonts=True,
#             UseISO19005_1=False,
#         )
#         doc.Close(False)
#         return True
#     except Exception as e:
#         print(f"    └─Word変換失敗: {src.name} → {e}")
#         return False

# def _convert_ppt_to_pdf(app_ppt, src, dst) -> bool:
#     # ★ 文字列に統一（Pathのまま渡さない）
#     src = str(src)
#     dst = str(dst)

#     # 画面は不可視にしない。最小化だけ（環境により効かない場合もある）
#     try:
#         app_ppt.Visible = True
#         app_ppt.WindowState = ppWindowMinimized
#     except Exception:
#         pass

#     def export(pres):
#         # ExportAsFixedFormat(Path, FixedFormatType, Intent,
#         #   FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides,
#         #   PrintRange, RangeType, SlideShowName, IncludeDocProperties,
#         #   KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1)
#         # 位置引数で渡す。不要箇所は None を詰めて RangeType まで届かせる。
#         pres.ExportAsFixedFormat(
#             dst,
#             ppFixedFormatTypePDF,
#             ppFixedFormatIntentPrint,
#             msoTrue,      # FrameSlides
#             1,            # HandoutOrder（既定）
#             1,            # OutputType（既定）
#             msoFalse,     # PrintHiddenSlides
#             None,         # PrintRange
#             ppPrintAll,   # RangeType
#             None,         # SlideShowName
#             msoTrue,      # IncludeDocProperties
#             msoTrue,      # KeepIRMSettings
#             msoTrue,      # DocStructureTags
#             msoTrue,      # BitmapMissingFonts
#             msoFalse      # UseISO19005_1
#         )

#     # 1st: プレゼンをウィンドウ無しで開く（アプリは可視のまま）
#     try:
#         pres = app_ppt.Presentations.Open(
#             src, ReadOnly=True, Untitled=False, WithWindow=False
#         )
#         export(pres)
#         pres.Close()
#         return True
#     except Exception as e1:
#         # 2nd: ウィンドウありで再試行
#         try:
#             pres = app_ppt.Presentations.Open(
#                 src, ReadOnly=True, Untitled=False, WithWindow=True
#             )
#             export(pres)
#             pres.Close()
#             return True
#         except Exception as e2:
#             print(f"    └─PPT変換失敗(非表示): {e1}")
#             print(f"    └─PPT変換失敗(可視化): {e2}")
#             return False

# def convert_from_get_files_to_PDF(
#     paths: Iterable[Union[str, Path]],
#     overwrite: bool = False
# ) -> List[Path]:
#     """
#     渡された「ファイルパスのリスト」を PDF 化。
#     - 同じディレクトリに同名.pdf を出力
#     - overwrite=False なら既存PDFはスキップ
#     - Word/PPT 以外はスキップ
#     - 開けない/失敗時は print してスキップ
#     """
#     # 正規化
#     targets: List[Path] = [Path(p) for p in paths if p]
#     if not targets:
#         return []

#     # 入力のうち、実在ファイル＋対応拡張子だけ抽出
#     todo: List[Path] = []
#     for src in targets:
#         if not src.exists() or not src.is_file():
#             print(f"⚠️ 見つからない/ファイルでない: {src}")
#             continue
#         ext = src.suffix.lower()
#         if ext in WORD_EXTS or ext in PPT_EXTS:
#             todo.append(src)
#         else:
#             print(f"ℹ️ 非対応拡張子スキップ: {src.name}")

#     if not todo:
#         return []

#     app_word: Optional[Any] = None
#     app_ppt:  Optional[Any] = None
#     made: List[Path] = []

#     kill_office_processes()

#     try:
#         if any(p.suffix.lower() in WORD_EXTS for p in todo):
#             app_word = win32.Dispatch("Word.Application")
#             app_word.Visible = False
#             # app_word.DisplayAlerts = 0  # 必要に応じて抑止

#         if any(p.suffix.lower() in PPT_EXTS for p in todo):
#             app_ppt = win32.Dispatch("PowerPoint.Application")
#             app_ppt.Visible = True           # 方針維持
#             try:
#                 app_ppt.WindowState = 2      # ppWindowMinimized=2（最小化）
#             except Exception:
#                 pass

#         for src in todo:
#             dst_pdf = src.with_suffix(".pdf")
#             if dst_pdf.exists() and not overwrite:
#                 print(f"↪ 既存PDFあり（スキップ）: {dst_pdf}")
#                 continue

#             try:
#                 ok = False
#                 if src.suffix.lower() in WORD_EXTS:
#                     if app_word is None:
#                         print(f"⚠️ Word アプリが初期化されていないためスキップ: {src}")
#                         continue
#                     ok = _convert_word_to_pdf(app_word, src, dst_pdf)
#                 else:
#                     if app_ppt is None:
#                         print(f"⚠️ PowerPoint アプリが初期化されていないためスキップ: {src}")
#                         continue
#                     ok = _convert_ppt_to_pdf(app_ppt, src, dst_pdf)

#                 if ok:
#                     print(f"✅ PDF化: {dst_pdf}")
#                     made.append(dst_pdf)
#                 else:
#                     print(f"⚠️ 失敗: {src}")
#             except Exception as e:
#                 # 開けない/変換失敗などは通知して続行
#                 print(f"⚠️ 変換エラー: {src} → {e}")

#     finally:
#         try:
#             if app_word is not None:
#                 app_word.Quit(False)
#         except Exception:
#             pass
#         try:
#             if app_ppt is not None:
#                 app_ppt.Quit()
#         except Exception:
#             pass
            
#     kill_office_processes()
#     return made

# # ---- 動作例 ----
# if __name__ == "__main__":
#     dirs = [
#         r"C:\Users\yohei\Downloads\R2-2206906",
#         r"C:\Users\yohei\Downloads\R2-2206905",
#         r"C:\Users\yohei\Downloads",
#         r"C:\Users\yohei\Downloads"
#     ]
#     files = [
#         "R2-2206906_C1-224008.docx",
#         "R2-2206905_C1-223972.docx",
#         "11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx",
#         "input_3gpp.xlsx"  # 非対応 → スキップされる
#     ]
#     combined = [a + "\\" + b for a, b in zip(dirs, files)]
#     print(combined)
#     out = convert_from_get_files_to_PDF(combined, overwrite=False)
#     print(f"✅ 出力: {out}")


# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Iterable, List, Optional, Any, Union
import subprocess
import multiprocessing as mp
import pythoncom
import win32com.client as win32

# ===== 拡張子 =====
WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
PPT_EXTS  = {".ppt", ".pptx", ".pptm"}

# ===== Word 定数 =====
wdExportFormatPDF              = 17
wdExportOptimizeForPrint       = 0   # 印刷向け
wdExportRangeAll               = 0
wdExportDocumentContent        = 0
wdExportCreateNoBookmarks      = 0

# MsoTriState
msoTrue  = -1
msoFalse = 0

# PowerPoint 定数
ppFixedFormatTypePDF     = 2
ppFixedFormatIntentPrint = 2
ppPrintAll               = 1
ppWindowMinimized        = 2


def kill_office_processes() -> None:
    """WINWORD.EXE と POWERPNT.EXE を強制終了（失敗は無視）。"""
    for image_name in ("WINWORD.EXE", "POWERPNT.EXE"):
        try:
            subprocess.run(
                ["taskkill", "/IM", image_name, "/F", "/T"],
                capture_output=True,
                text=True,
                check=False
            )
        except Exception:
            pass


def _convert_word_to_pdf(app_word, src: Path, dst: Path) -> bool:
    """Word → PDF（詳細パラメータ版）"""
    try:
        doc = app_word.Documents.Open(str(src), ReadOnly=True, Visible=False)
        doc.ExportAsFixedFormat(
            OutputFileName=str(dst),
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=wdExportOptimizeForPrint,
            Range=wdExportRangeAll,
            Item=wdExportDocumentContent,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=wdExportCreateNoBookmarks,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
        doc.Close(False)
        return True
    except Exception as e:
        print(f"    └─Word変換失敗: {src.name} → {e}", flush=True)
        return False


def _convert_ppt_to_pdf(app_ppt, src: Path, dst: Path) -> bool:
    """PowerPoint → PDF"""
    src_s = str(src)
    dst_s = str(dst)

    try:
        app_ppt.Visible = True
        app_ppt.WindowState = ppWindowMinimized
    except Exception:
        pass

    def export(pres):
        pres.ExportAsFixedFormat(
            dst_s,
            ppFixedFormatTypePDF,
            ppFixedFormatIntentPrint,
            msoTrue,      # FrameSlides
            1,            # HandoutOrder
            1,            # OutputType
            msoFalse,     # PrintHiddenSlides
            None,         # PrintRange
            ppPrintAll,   # RangeType
            None,         # SlideShowName
            msoTrue,      # IncludeDocProperties
            msoTrue,      # KeepIRMSettings
            msoTrue,      # DocStructureTags
            msoTrue,      # BitmapMissingFonts
            msoFalse      # UseISO19005_1
        )

    try:
        pres = app_ppt.Presentations.Open(src_s, ReadOnly=True, Untitled=False, WithWindow=False)
        export(pres)
        pres.Close()
        return True
    except Exception as e1:
        try:
            pres = app_ppt.Presentations.Open(src_s, ReadOnly=True, Untitled=False, WithWindow=True)
            export(pres)
            pres.Close()
            return True
        except Exception as e2:
            print(f"    └─PPT変換失敗(非表示): {e1}", flush=True)
            print(f"    └─PPT変換失敗(可視化): {e2}", flush=True)
            return False


def _reserve_output_path(
    output_dir: Path,
    stem: str,
    overwrite: bool,
    lock: Any,
    reserved: Any
) -> Path:
    """
    並列実行用の重複回避。プロセス共有の reserved(dict) と Lock を使って
    同名 .pdf の割当てを原子的に決める。
    """
    if overwrite:
        final = output_dir / f"{stem}.pdf"
        with lock:
            reserved[str(final)] = True
        return final

    with lock:
        i = 0
        while True:
            name = f"{stem}.pdf" if i == 0 else f"{stem}_{i}.pdf"
            final = output_dir / name
            key = str(final)
            if (not final.exists()) and (key not in reserved):
                reserved[key] = True  # 予約
                return final
            i += 1


def _worker_loop(
    worker_id: int,
    task_q: mp.Queue,
    results: Any,
    output_dir_s: str,
    overwrite: bool,
    lock: Any,
    reserved: Any
) -> None:
    """
    ワーカープロセス本体。
    - 自前で COM 初期化/終了
    - 必要に応じて Word/PowerPoint を遅延起動
    - タスクは動的キューから取得（None で終了）
    """
    pythoncom.CoInitialize()
    app_word: Optional[Any] = None
    app_ppt:  Optional[Any] = None
    output_dir = Path(output_dir_s)

    try:
        while True:
            src_s = task_q.get()
            if src_s is None:
                break  # sentinel

            src = Path(src_s)
            if not src.exists() or not src.is_file():
                print(f"[W{worker_id}] ⚠️ 見つからない/ファイルでない: {src}", flush=True)
                continue

            ext = src.suffix.lower()
            if ext not in WORD_EXTS and ext not in PPT_EXTS:
                print(f"[W{worker_id}] ℹ️ 非対応拡張子スキップ: {src.name}", flush=True)
                continue

            dst_pdf = _reserve_output_path(output_dir, src.stem, overwrite, lock, reserved)

            try:
                if ext in WORD_EXTS:
                    if app_word is None:
                        app_word = win32.Dispatch("Word.Application")
                        app_word.Visible = False
                    ok = _convert_word_to_pdf(app_word, src, dst_pdf)
                else:
                    if app_ppt is None:
                        app_ppt = win32.Dispatch("PowerPoint.Application")
                        app_ppt.Visible = True
                        try:
                            app_ppt.WindowState = ppWindowMinimized
                        except Exception:
                            pass
                    ok = _convert_ppt_to_pdf(app_ppt, src, dst_pdf)

                if ok:
                    results.append(str(dst_pdf))
                    print(f"[W{worker_id}] ✅ PDF化: {src.name} → {dst_pdf}", flush=True)
                else:
                    print(f"[W{worker_id}] ⚠️ 失敗: {src}", flush=True)

            except Exception as e:
                print(f"[W{worker_id}] ⚠️ 変換エラー: {src} → {e}", flush=True)

    finally:
        # COM アプリ終了
        try:
            if app_word is not None:
                app_word.Quit(False)
        except Exception:
            pass
        try:
            if app_ppt is not None:
                app_ppt.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def convert_list_to_pdf_in_dir_parallel(
    paths: Iterable[Union[str, Path]],
    output_dir: Union[str, Path],
    pdf_dir: str,
    overwrite: bool = False,
    num_workers: int = 10
) -> List[Path]:
    """
    与えられたファイル群（Word/PPT）を PDF 化して、指定 output_dir に保存。
    動的キューで並列実行（既定 5 ワーカー）。

    - output_dir が無ければ作成
    - overwrite=False は重複回避名（_1, _2 …）を予約ベースで安全に割当て
    - 非存在/非対応はスキップ
    """

    out_dir_path = Path(output_dir) / pdf_dir
    out_dir = Path(out_dir_path).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    # 入力正規化 & 先にフィルタ（無駄なキュー投入を避ける）
    todo: List[str] = []
    for p in paths:
        s = str(p)
        if not s:
            continue
        ext = Path(s).suffix.lower()
        if ext in WORD_EXTS or ext in PPT_EXTS:
            todo.append(s)

    if not todo:
        return []

    # 親プロセス側で先に Office を全滅させる（ロック解除用）
    kill_office_processes()

    manager = mp.Manager()
    results = manager.list()       # 出力ファイルの共有リスト
    reserved = manager.dict()      # 予約された出力ファイル名
    lock = manager.Lock()          # 予約用ロック
    task_q: mp.Queue = mp.Queue()  # 高速なプロセス間 Queue

    # タスク投入（動的キュー）
    for s in todo:
        task_q.put(s)

    # 終了用センチネル
    for _ in range(num_workers):
        task_q.put(None)

    procs: List[mp.Process] = []
    for wid in range(1, num_workers + 1):
        p = mp.Process(
            target=_worker_loop,
            args=(wid, task_q, results, str(out_dir), overwrite, lock, reserved),
            daemon=False
        )
        p.start()
        procs.append(p)

    for p in procs:
        p.join()

    # 念のため Office を片付ける（他のインスタンスに注意）
    kill_office_processes()

    return [Path(s) for s in list(results)]


# ---- 動作例 ----
if __name__ == "__main__":
    # PyInstaller 対策（Windowsのspawn環境）
    mp.freeze_support()

    paths = [
        r"C:\Users\yohei\Downloads\R2-2206906\R2-2206906_C1-224008.docx",
        r"C:\Users\yohei\Downloads\R2-2206905\R2-2206905_C1-223972.docx",
        r"C:\Users\yohei\Downloads\11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx",
        r"C:\Users\yohei\Downloads\input_3gpp.xlsx",  # 非対応 → 自動スキップ
    ]

    out_dir = r"C:\Users\yohei\Downloads\PDF_OUT"
    result = convert_list_to_pdf_in_dir_parallel(paths, out_dir, overwrite=False, num_workers=5)

    print("✅ 出力一覧:")
    for p in result:
        print("  -", p)
