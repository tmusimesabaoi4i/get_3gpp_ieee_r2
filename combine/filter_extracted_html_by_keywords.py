# -*- coding: utf-8 -*-
"""
filter_extracted_html_by_keywords（完全版）
- 入力: 本ツールが生成した HTML（<h2>=ファイル名, <h3>=Slide N, <li>=段落）
- 処理: 指定キーワードに一致する <li> だけ再収集
- 出力: 1つの HTML（見出し <h2>, <h3> は維持）
- 照合: 全角/半角・大小無視（NFKC+casefold）。英数字/アンダーバーは語境界で完全一致
- 入力パス: str, Path, WindowsPath(...) の repr 文字列まで受容。重複除去・順序維持
- 出力先: フォルダ/ファイルどちらでもOK。安全なファイル名に正規化し、原子的置換で書き込み
"""

from __future__ import annotations
from pathlib import Path
from typing import Iterable, List, Dict, Optional
from collections import OrderedDict
from html.parser import HTMLParser
import html as _html
import unicodedata, re, time
from tempfile import NamedTemporaryFile

# ========= 正規化・キーワード =========
def _nfkc_casefold(s: str) -> str:
    return unicodedata.normalize("NFKC", s).casefold()

def _compile_keyword_patterns(keywords: Iterable[str]) -> List[re.Pattern]:
    """
    照合ルール:
      - 全角/半角・大小を無視（NFKC+casefold）
      - 英数字/アンダーバーは語境界で完全一致（部分一致を防ぐ）
      - それ以外（和文等）は「含まれる」一致（語境界が曖昧なため）
    """
    pats: List[re.Pattern] = []
    for kw in keywords:
        kw = (kw or "").strip()
        if not kw:
            continue
        kn = _nfkc_casefold(kw)
        pats.append(re.compile(rf'(?<![0-9A-Za-z_]){re.escape(kn)}(?![0-9A-Za-z_])'))
    return pats

# ========= HTML パーサ =========
class _LiFilterParser(HTMLParser):
    """
    想定入力: <h2>=ファイル名, <h3>=Slide N, <li>=段落
    条件一致の <li> だけ result に保持
    """
    def __init__(self, patterns: List[re.Pattern]):
        super().__init__(convert_charrefs=True)
        self.patterns = patterns
        self.current_file: Optional[str] = None
        self.current_slide: Optional[int] = None
        self.in_h2 = False
        self.in_h3 = False
        self.in_li = False
        self._buf_h2: List[str] = []
        self._buf_h3: List[str] = []
        self._buf_li: List[str] = []
        # Ordered: {file: {"word":[...], "ppt": OrderedDict{slide_no:[...]}}}
        self.result: "OrderedDict[str, Dict[str, object]]" = OrderedDict()

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t == "h2":
            self.in_h2 = True; self._buf_h2 = []
        elif t == "h3":
            self.in_h3 = True; self._buf_h3 = []
        elif t == "li":
            self.in_li = True; self._buf_li = []
        elif t == "br" and self.in_li:
            self._buf_li.append("\n")

    def handle_endtag(self, tag):
        t = tag.lower()
        if t == "h2":
            self.in_h2 = False
            name = "".join(self._buf_h2).strip()
            self.current_file = name or None
            self.current_slide = None
            if self.current_file and self.current_file not in self.result:
                self.result[self.current_file] = {"word": [], "ppt": OrderedDict()}
        elif t == "h3":
            self.in_h3 = False
            label = "".join(self._buf_h3).strip()
            m = re.search(r"Slide\s+(\d+)", label, re.IGNORECASE)
            self.current_slide = int(m.group(1)) if m else None
            if self.current_file and self.current_slide is not None:
                self.result[self.current_file]["ppt"].setdefault(self.current_slide, [])  # type: ignore[index]
        elif t == "li":
            self.in_li = False
            raw = "".join(self._buf_li).strip()
            if not raw or not self.current_file:
                return
            text_norm = _nfkc_casefold(raw)
            hit = any(p.search(text_norm) for p in self.patterns) if self.patterns else False
            if not hit:
                return
            if self.current_slide is not None:
                self.result[self.current_file]["ppt"].setdefault(self.current_slide, []).append(raw)  # type: ignore[index]
            else:
                self.result[self.current_file]["word"].append(raw)  # type: ignore[index]

    def handle_data(self, data):
        if self.in_h2:
            self._buf_h2.append(data)
        elif self.in_h3:
            self._buf_h3.append(data)
        elif self.in_li:
            self._buf_li.append(data)

# ========= 入力パス正規化 =========
def _coerce_to_path_list(html_paths) -> List[Path]:
    """
    html_paths: Iterable[Path | str] だけでなく、
               "WindowsPath('C:/...')" の repr 文字列も受け付ける。
    - 重複は除去（Windows を想定して大小無視）、順序は維持
    - 相対は resolve(strict=False) で実質絶対化
    """
    out: List[Path] = []
    seen: set[str] = set()
    if html_paths is None:
        return out

    for x in html_paths:
        s = None
        # Path-like を文字列化（失敗時は str）
        try:
            from os import fspath
            s = fspath(x)
        except Exception:
            s = str(x)
        s = (s or "").strip()
        # "WindowsPath('C:/...')" をアンラップ
        m = re.fullmatch(r"WindowsPath\((['\"])(.+?)\1\)", s)
        if m:
            s = m.group(2)
        p = Path(s)
        try:
            q = p.resolve(strict=False)
        except Exception:
            q = p
        key = str(q).lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(q)
    return out

# ========= 出力 HTML 構築 =========
def _build_output_html(agg: "OrderedDict[str, Dict[str, object]]",
                       keywords: List[str]) -> str:
    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="ja">')
    parts.append('<meta charset="UTF-8">')
    parts.append("<title>抽出テキスト（キーワードフィルタ）</title>")
    parts.append("<style>"
                 "ul{margin:.3rem 1.2rem;} h2{margin-top:1.2rem;}"
                 ".file-block{margin-bottom:1rem;}"
                 ".word-block,.slide-block{border:1px solid #e3e3e3;padding:.6rem;border-radius:.5rem;}"
                 ".meta{color:#666;font-size:.9em;margin-bottom:.6rem}"
                 "</style>")
    parts.append("<body>")
    if keywords:
        kdisp = ", ".join(_html.escape(k) for k in keywords if (k or "").strip())
        parts.append(f"<div class='meta'>抽出キーワード: {kdisp}</div>")

    any_hit = False
    for fname, d in agg.items():
        word_list: List[str] = d.get("word", [])  # type: ignore[assignment]
        ppt_map: "OrderedDict[int, List[str]]" = d.get("ppt", OrderedDict())  # type: ignore[assignment]
        if not word_list and not any(ppt_map.values()):
            continue
        any_hit = True
        parts.append(f"<div class='file-block'><h2>{_html.escape(fname)}</h2>")
        if word_list:
            parts.append("<div class='word-block'><ul>")
            for p in word_list:
                parts.append(f"<li>{_html.escape(p).replace(chr(10), '<br>')}</li>")
            parts.append("</ul></div>")
        for slide_no in sorted(ppt_map.keys()):
            items = ppt_map[slide_no]
            if not items:
                continue
            parts.append(f"<div class='slide-block'><h3>Slide {slide_no}</h3><ul>")
            for p in items:
                parts.append(f"<li>{_html.escape(p).replace(chr(10), '<br>')}</li>")
            parts.append("</ul></div>")
        parts.append("</div>")  # .file-block
    if not any_hit:
        parts.append("<p>（一致は見つかりませんでした）</p>")
    parts.append("</body></html>")
    return "\n".join(parts)

# ========= 出力パスの正規化・原子的書き込み =========
_INVALID_WIN_CHARS = r'[<>:"/\\|?*\x00-\x1F]'
_RESERVED_NAMES = {"CON","PRN","AUX","NUL"} | {f"COM{i}" for i in range(1,10)} | {f"LPT{i}" for i in range(1,10)}

def _sanitize_windows_filename(name: str) -> str:
    clean = re.sub(_INVALID_WIN_CHARS, "_", name)
    clean = clean.rstrip(" .")
    base = clean.split(".")[0].upper()
    if base in _RESERVED_NAMES:
        clean = "_" + clean
    return clean or "untitled.html"

def _decide_output_path(output_html_path: str | Path) -> Path:
    p = Path(output_html_path)
    if p.is_dir():
        p = p / "filtered.html"
    if p.suffix.lower() not in {".html", ".htm"}:
        p = p.with_suffix(".html") if p.suffix else p.with_name(p.name + ".html")
    p = p.with_name(_sanitize_windows_filename(p.name))
    p.parent.mkdir(parents=True, exist_ok=True)
    return p

def _atomic_write_text(path: Path, text: str, encoding: str = "utf-8", retries: int = 3, delay: float = 0.25) -> None:
    for attempt in range(1, retries + 1):
        tmp = None
        try:
            with NamedTemporaryFile("w", encoding=encoding, delete=False, dir=str(path.parent), newline="") as f:
                f.write(text)
                tmp = Path(f.name)
            tmp.replace(path)
            return
        except Exception as e:
            try:
                if tmp and tmp.exists():
                    tmp.unlink()
            except Exception:
                pass
            if attempt < retries:
                time.sleep(delay); delay = min(delay * 2, 1.0)
            else:
                raise RuntimeError(f"atomic write failed at {path}: {e}") from e

# ========= 公開API（あなたの呼び出し形で使える） =========
def filter_extracted_html_by_keywords(
    html_paths: Iterable[str | Path],
    keywords: Iterable[str],
    output_html_path: str | Path,
) -> Path:
    """
    指定HTML群から、指定語を含む <li> 段落だけを再収集して1つのHTMLにまとめる。
    - 見出しは <h2>=ファイル名、<h3>=Slide番号 を維持
    - 照合: NFKC+casefold、英数字は語境界で完全一致
    - output_html_path にディレクトリを渡した場合は 'filtered.html' を自動付与
    戻り値: 生成したHTMLの Path
    """
    # 入力整形
    paths = _coerce_to_path_list(html_paths)
    kws = [k for k in (keywords or []) if (k or "").strip()]
    if not kws:
        raise ValueError("keywords が空です。少なくとも1語を指定してください。")
    pats = _compile_keyword_patterns(kws)

    # 集約
    agg: "OrderedDict[str, Dict[str, object]]" = OrderedDict()
    for p in paths:
        if not p.exists() or not p.is_file():
            print(f"[SKIP] not found: {p}")
            continue
        parser = _LiFilterParser(pats)
        text = None
        for enc in ("utf-8", "utf-8-sig", "cp932"):
            try:
                text = p.read_text(encoding=enc)
                break
            except Exception:
                continue
        if text is None:
            print(f"[SKIP] cannot read: {p}")
            continue
        parser.feed(text); parser.close()
        for fname, d in parser.result.items():
            if fname not in agg:
                agg[fname] = {"word": [], "ppt": OrderedDict()}
            agg[fname]["word"].extend(d.get("word", []))  # type: ignore[index]
            ppt_dst: "OrderedDict[int, List[str]]" = agg[fname]["ppt"]  # type: ignore[index]
            ppt_src: "OrderedDict[int, List[str]]" = d.get("ppt", OrderedDict())  # type: ignore[assignment]
            for no, lst in ppt_src.items():
                ppt_dst.setdefault(no, []).extend(lst)

    # 出力
    out_path = _decide_output_path(output_html_path)
    out_html = _build_output_html(agg, kws)
    _atomic_write_text(out_path, out_html, encoding="utf-8")
    return out_path

# # 使い方例:
if __name__ == "__main__":
    outs = filter_extracted_html_by_keywords(
        html_paths=[
            r"C:\Users\yohei\Downloads\a_part1.html"
        ],
        keywords=["UE", "BSR design", "HARQ"],  # 全角半角/大小は自動で吸収
        output_html_path=r"C:\Users\yohei\Downloads\html_out\filtered.html",
    )
    print("→", outs)
