# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from typing import Iterable, List, Optional, Any, Union
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
    """
    WINWORD.EXE と POWERPNT.EXE を強制終了。
    失敗しても例外は投げず、静かに無視（printもしない）。
    ※ 強制終了により未保存データは失われます。
    """
    # /F: 強制終了, /T: 子プロセス含む, /IM: 画像名指定
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


def _read_absolute_paths_pandas(folder: Path) -> List[Path]:
    """folder/get_files.xlsx のA列から絶対パスを読む（pandas使用）。無ければ空リスト。"""
    xlsx = folder / "get_files.xlsx"
    if not xlsx.exists():
        print(f"ℹ️ get_files.xlsx なし（処理スキップ）: {xlsx}")
        return []
    df = pd.read_excel(xlsx, header=None, usecols=[0], names=["path"], dtype=str, engine="openpyxl")
    paths: List[Path] = []
    for s in df["path"].dropna():
        s = str(s).strip().strip('"').strip("'")
        if not s:
            continue
        p = Path(s)
        if not p.is_absolute():
            print(f"↪ 相対パスのためスキップ: {s}")
            continue
        paths.append(p)
    return paths

def _convert_word_to_pdf(app_word, src: Path, dst: Path) -> bool:
    """Word → PDF（詳細パラメータ版）"""
    try:
        doc = app_word.Documents.Open(str(src), ReadOnly=True, Visible=False)
        # ご指定の引数をそのまま反映
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
        print(f"    └─Word変換失敗: {src.name} → {e}")
        return False

def _convert_ppt_to_pdf(app_ppt, src, dst) -> bool:
    # ★ 文字列に統一（Pathのまま渡さない）
    src = str(src)
    dst = str(dst)

    # 画面は不可視にしない。最小化だけ（環境により効かない場合もある）
    try:
        app_ppt.Visible = True
        app_ppt.WindowState = ppWindowMinimized
    except Exception:
        pass

    def export(pres):
        # ExportAsFixedFormat(Path, FixedFormatType, Intent,
        #   FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides,
        #   PrintRange, RangeType, SlideShowName, IncludeDocProperties,
        #   KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1)
        # 位置引数で渡す。不要箇所は None を詰めて RangeType まで届かせる。
        pres.ExportAsFixedFormat(
            dst,
            ppFixedFormatTypePDF,
            ppFixedFormatIntentPrint,
            msoTrue,      # FrameSlides
            1,            # HandoutOrder（既定）
            1,            # OutputType（既定）
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

    # 1st: プレゼンをウィンドウ無しで開く（アプリは可視のまま）
    try:
        pres = app_ppt.Presentations.Open(
            src, ReadOnly=True, Untitled=False, WithWindow=False
        )
        export(pres)
        pres.Close()
        return True
    except Exception as e1:
        # 2nd: ウィンドウありで再試行
        try:
            pres = app_ppt.Presentations.Open(
                src, ReadOnly=True, Untitled=False, WithWindow=True
            )
            export(pres)
            pres.Close()
            return True
        except Exception as e2:
            print(f"    └─PPT変換失敗(非表示): {e1}")
            print(f"    └─PPT変換失敗(可視化): {e2}")
            return False

def convert_from_get_files_to_PDF(
    paths: Iterable[Union[str, Path]],
    overwrite: bool = False
) -> List[Path]:
    """
    渡された「ファイルパスのリスト」を PDF 化。
    - 同じディレクトリに同名.pdf を出力
    - overwrite=False なら既存PDFはスキップ
    - Word/PPT 以外はスキップ
    - 開けない/失敗時は print してスキップ
    """
    # 正規化
    targets: List[Path] = [Path(p) for p in paths if p]
    if not targets:
        return []

    # 入力のうち、実在ファイル＋対応拡張子だけ抽出
    todo: List[Path] = []
    for src in targets:
        if not src.exists() or not src.is_file():
            print(f"⚠️ 見つからない/ファイルでない: {src}")
            continue
        ext = src.suffix.lower()
        if ext in WORD_EXTS or ext in PPT_EXTS:
            todo.append(src)
        else:
            print(f"ℹ️ 非対応拡張子スキップ: {src.name}")

    if not todo:
        return []

    app_word: Optional[Any] = None
    app_ppt:  Optional[Any] = None
    made: List[Path] = []

    kill_office_processes()

    try:
        if any(p.suffix.lower() in WORD_EXTS for p in todo):
            app_word = win32.Dispatch("Word.Application")
            app_word.Visible = False
            # app_word.DisplayAlerts = 0  # 必要に応じて抑止

        if any(p.suffix.lower() in PPT_EXTS for p in todo):
            app_ppt = win32.Dispatch("PowerPoint.Application")
            app_ppt.Visible = True           # 方針維持
            try:
                app_ppt.WindowState = 2      # ppWindowMinimized=2（最小化）
            except Exception:
                pass

        for src in todo:
            dst_pdf = src.with_suffix(".pdf")
            if dst_pdf.exists() and not overwrite:
                print(f"↪ 既存PDFあり（スキップ）: {dst_pdf}")
                continue

            try:
                ok = False
                if src.suffix.lower() in WORD_EXTS:
                    if app_word is None:
                        print(f"⚠️ Word アプリが初期化されていないためスキップ: {src}")
                        continue
                    ok = _convert_word_to_pdf(app_word, src, dst_pdf)
                else:
                    if app_ppt is None:
                        print(f"⚠️ PowerPoint アプリが初期化されていないためスキップ: {src}")
                        continue
                    ok = _convert_ppt_to_pdf(app_ppt, src, dst_pdf)

                if ok:
                    print(f"✅ PDF化: {dst_pdf}")
                    made.append(dst_pdf)
                else:
                    print(f"⚠️ 失敗: {src}")
            except Exception as e:
                # 開けない/変換失敗などは通知して続行
                print(f"⚠️ 変換エラー: {src} → {e}")

    finally:
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
            
    kill_office_processes()
    return made

# ---- 動作例 ----
if __name__ == "__main__":
    dirs = [
        r"C:\Users\yohei\Downloads\R2-2206906",
        r"C:\Users\yohei\Downloads\R2-2206905",
        r"C:\Users\yohei\Downloads",
        r"C:\Users\yohei\Downloads"
    ]
    files = [
        "R2-2206906_C1-224008.docx",
        "R2-2206905_C1-223972.docx",
        "11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx",
        "input_3gpp.xlsx"  # 非対応 → スキップされる
    ]
    combined = [a + "\\" + b for a, b in zip(dirs, files)]
    print(combined)
    out = convert_from_get_files_to_PDF(combined, overwrite=False)
    print(f"✅ 出力: {out}")