import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import os
import re
import pandas as pd
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from typing import Union, Optional, Iterable, Union, List, overload

import subprocess

from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
    )

def _sanitize_token(
        s: str
    ) -> str:
    s = (s or "").strip()
    s = re.sub(r'\.html?$', '', s, flags=re.I)
    s = re.sub(r'[\\/:*?"<>|\s]+', "_", s)
    return s

def _change_drive(
        base: Path,
        drive: str
    ) -> Path:
    d = (drive or "").strip().rstrip(":").upper()
    if not d or len(d) != 1 or not d.isalpha():
        raise ValueError(f"Invalid drive: {drive!r}")
    s = str(base)
    s2 = re.sub(r"^[A-Za-z]:", f"{d}:", s)
    return Path(s2)

def _sanitize_filename(
        s: str
    ) -> str:
    s = (s or "").strip()
    return re.sub(r'[\\/:\*\?"<>\|]+', "_", s)

def cell(
        df,
        r: int,
        c: int
    ) -> str:
    try:
        v = df.iat[r, c]
    except Exception:
        v = None
    return "" if (v is None or pd.isna(v)) else str(v).strip()

# def get_downloads_path(
#         drive: Optional[str] = None
#     ) -> Path:
#     home = Path.home()
#     names = ["Downloads", "downloads", "Download", "ダウンロード"]

#     # 1) まずは現在のホーム配下で探索
#     for name in names:
#         p = (home / name)
#         if p.exists() and p.is_dir():
#             return p.resolve()

#     # 2) 見つからなければ、指定ドライブに差し替えて再探索（Windows想定）
#     if drive and (home.drive or os.name == "nt"):
#         alt_home = _change_drive(home, drive)
#         for name in names:
#             p = (alt_home / name)
#             if p.exists() and p.is_dir():
#                 return p.resolve()

#     # 3) どれも無ければエラー
#     tried = [str(home / n) for n in names]
#     if drive and (home.drive or os.name == "nt"):
#         tried += [str(_change_drive(home, drive) / n) for n in names]
#     raise FileNotFoundError(f"{emo.warn} Downloads folder not found. Tried: " + ", ".join(tried))

from pathlib import Path
from typing import Optional
import os

def _change_drive(base_home: Path, drive: str) -> Path:
    """
    base_home の「ドライブ部分」だけを drive に差し替える（Windows想定）。
    例: base_home=C:\\Users\\foo, drive='D' → D:\\Users\\foo
    """
    d = (drive or "").strip().rstrip(":")
    if not d:
        return base_home
    # base_home.parts 例: ('C:\\', 'Users', 'foo')
    tail = Path(*base_home.parts[1:]) if base_home.anchor else base_home
    return Path(f"{d.upper()}:") / tail

def _is_windows() -> bool:
    return os.name == "nt"

def get_downloads_path(drive: Optional[str] = None) -> Path:
    """
    ・drive が与えられた場合: そのドライブ側のホーム配下で Downloads 相当の候補を探索
    ・drive が None の場合: 現在のホーム配下で探索
    ・見つからなければ、（drive 指定時は）現在ホーム側も試してからエラー
    """
    home = Path.home()
    names = ["Downloads", "downloads", "Download", "ダウンロード"]

    tried: list[str] = []

    # A) drive 指定がある場合は、まずそのドライブ側を優先
    if drive and _is_windows():
        alt_home = _change_drive(home, drive)
        for name in names:
            p = alt_home / name
            tried.append(str(p))
            if p.is_dir():
                return p.resolve()

        # （保険）ドライブ側に無ければ、現在ホーム側も試す
        for name in names:
            p = home / name
            tried.append(str(p))
            if p.is_dir():
                return p.resolve()

        # どちらにも無ければエラー
        raise FileNotFoundError(
            f"{emo.warn} Downloads folder not found. Tried: " + ", ".join(tried)
        )

    # B) drive 指定が無い場合は、現在ホーム配下のみで探索して返す
    for name in names:
        p = home / name
        tried.append(str(p))
        if p.is_dir():
            return p.resolve()

    # 見つからなければエラー
    raise FileNotFoundError(
        f"{emo.warn} Downloads folder not found. Tried: " + ", ".join(tried)
    )

def build_case_folder_from_excel(
    folder_abs_path: str, filename: str,
    sheet: Union[int, str] = 0,
    drive: Optional[str] = None,
    ) -> Path:

    excel_path = Path(folder_abs_path) / filename

    if not excel_path.is_absolute():
        raise ValueError(f"{emo.warn} folder_abs_path は絶対パスで指定してください。")
    if not excel_path.exists():
        raise FileNotFoundError(f"{emo.warn} Excel ファイルが見つかりません: {p}")

    suffix = excel_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        engine = "openpyxl"
    elif suffix == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"{emo.warn} 未対応の拡張子です: {suffix}（.xlsx / .xlsm / .xls）")
    df = pd.read_excel(
        excel_path,
        sheet_name=sheet,
        engine=engine,
        header=None,
        dtype=str
    )

    date = _sanitize_filename(cell(df, 0, 1))
    case_id = _sanitize_filename(cell(df, 1, 1))

    if not date or not case_id:
        raise ValueError(f"{emo.warn} B1/B2 が空です。date='{date}' case_id='{case_id}'")

    downloads = get_downloads_path(drive=drive)

    downloads_filename = str(case_id)+'_'+str(date)

    create_subfolder_when_absent(downloads, downloads_filename)

    target = Path(downloads) / downloads_filename

    return target.resolve()

def _get_3gpp_html_name_single(
        url: str
    ) -> str:
    if url is None:
        raise ValueError(f"{emo.warn} URL が None です。")
    url = str(url).strip()
    if not url:
        raise ValueError(f"{emo.warn} URL が空文字です。")

    parsed = urlparse(url)
    segments = [s for s in parsed.path.split('/') if s]

    if not segments:
        raise ValueError(f"{emo.warn} URLにパスがありません: {url}")

    if segments[-1].lower() == "docs":
        if len(segments) < 2:
            raise ValueError(f"{emo.warn} Docs の前にシリーズ名が見つかりません: {url}")
        series = segments[-2]
    else:
        series = segments[-1]

    series = re.sub(r'[\\/:*?"<>|]+', "_", series).strip()
    series = re.sub(r'\.html?$', '', series, flags=re.I)

    if not series:
        raise ValueError(f"{emo.warn} シリーズ名を抽出できませんでした: {url}")

    series_s = _sanitize_token(series)
    return f"{series_s}.html"

@overload
def get_3gpp_html_name(urls: str) -> str: ...
@overload
def get_3gpp_html_name(urls: Iterable[str]) -> List[str]: ...
def get_3gpp_html_name(
        urls: Union[str, Iterable[str]]
    ) -> Union[str, List[str]]:

    if isinstance(urls, str):
        return _get_3gpp_html_name_single(urls)

    results: List[str] = []
    for i, u in enumerate(urls):
        try:
            results.append(_get_3gpp_html_name_single(u))
        except Exception as e:
            raise ValueError(f"{emo.warn} {i} 番目のURLでエラー: {e}") from e
    return results

def _get_ieee_html_name_single(
        url: str
    ) -> str:
    if url is None:
        raise ValueError(f"{emo.warn} URL が None です。")
    url = str(url).strip()
    if not url:
        raise ValueError(f"{emo.warn} URL が空文字です。")

    parsed = urlparse(url)
    segments = [s for s in parsed.path.split('/') if s]

    if not segments:
        raise ValueError(f"{emo.warn} URLにパスがありません: {url}")

    q = parse_qs(parsed.query, keep_blank_values=True)
    q = {k.lower(): v for k, v in q.items()}

    def first(*keys: str) -> str | None:
        for k in keys:
            if k in q and len(q[k]) > 0:
                return q[k][0]
        return None

    year  = first("is_year", "year")
    group = first("is_group", "group")
    n     = first("n")

    if not year:
        raise ValueError(f"is_year/year が見つかりません: {url}")
    if not group:
        raise ValueError(f"is_group/group が見つかりません: {url}")
    if not n:
        raise ValueError(f"n が見つかりません: {url}")

    year_s  = _sanitize_token(year)
    group_s = _sanitize_token(group)
    n_s     = _sanitize_token(n)

    return f"{year_s}_{group_s}_{n_s}.html"

@overload
def get_ieee_html_name(urls: str) -> str: ...
@overload
def get_ieee_html_name(urls: Iterable[str]) -> List[str]: ...
def get_ieee_html_name(
        urls: Union[str, Iterable[str]]
    ) -> Union[str, List[str]]:
    if isinstance(urls, str):
        return _get_ieee_html_name_single(urls)

    results: List[str] = []
    for i, u in enumerate(urls):
        try:
            results.append(_get_ieee_html_name_single(u))
        except Exception as e:
            raise ValueError(f"{emo.warn} {i} 番目のURLでエラー: {e}") from e
    return results


def _normalize_to_http_scheme(target: str) -> str:
    """
    任意の表記（host:port, http://..., https://..., user:pass@host:port など）を
    常に http:// スキームの URL 文字列に正規化して返す。
    """
    if not target:
        raise ValueError("empty proxy string")

    s = target.strip()
    # 全角 ; ＝ を半角に
    s = s.replace("；", ";").replace("＝", "=")

    # 既存スキームは除去して http:// を付け直す（user:pass@ を保持）
    s = re.sub(r'^[a-z][a-z0-9+.\-]*://', '', s, flags=re.I)
    s = s.lstrip('/')  # 変な先頭スラッシュを念のため除去
    return f"http://{s}"

def _extract_mapping_value(proxy_blob: str, prefer: str = "http") -> Optional[str]:
    """
    'http=host:port;https=host:port' のような文字列から希望スキームの値を抜き出す。
    単一 'host:port' 形式ならそれを返す。
    """
    if not proxy_blob:
        return None

    s = proxy_blob.strip().replace("；", ";").replace("＝", "=")

    # 'http=...;https=...' 形式
    if '=' in s:
        mapping = {}
        for part in s.split(';'):
            part = part.strip()
            if not part:
                continue
            if '=' in part:
                k, v = part.split('=', 1)
                mapping[k.strip().lower()] = v.strip()
        # 優先スキーム → 代替スキーム(https) → 代替スキーム(http) → 最初の値
        if prefer in mapping and mapping[prefer]:
            return mapping[prefer]
        if prefer != "https" and "https" in mapping and mapping["https"]:
            return mapping["https"]
        if prefer != "http" and "http" in mapping and mapping["http"]:
            return mapping["http"]
        # どれも無い場合は最初の非空値
        for v in mapping.values():
            if v:
                return v
        return None

    # 単一 'host:port' 形式
    return s or None

def _read_cmd_output(args: list[str]) -> str:
    """
    コマンドを実行して標準出力を文字列で返す（文字化けを避けるためエンコーディング広めに許容）。
    エラー時は空文字を返す。
    """
    try:
        cp = subprocess.run(
            args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if cp.returncode != 0:
            return ""
        return cp.stdout or ""
    except Exception:
        return ""

def get_proxy_from_cmd(prefer: str = "http", allow_env_fallback: bool = True) -> Optional[str]:
    """
    コマンドプロンプトの出力からプロキシ設定を取得し、'http://host:port' 形式の文字列を返す。
    見つからない場合は None を返す。

    探索順:
      1) netsh winhttp show proxy（WinHTTP プロキシ）
      2) reg query HKCU ... Internet Settings（WinINET/IE の ProxyServer/ProxyEnable）
      3) 環境変数 HTTPS_PROXY / HTTP_PROXY（allow_env_fallback=True の場合）

    Parameters
    ----------
    prefer : str
        'http' または 'https' を推奨スキームとして解釈に使う（デフォルト 'http'）。
    allow_env_fallback : bool
        True のとき、最後に環境変数をフォールバックとして参照。

    Returns
    -------
    Optional[str]
        正規化済み 'http://host:port' 形式の文字列。見つからなければ None。
    """
    prefer = (prefer or "http").lower()

    # 1) netsh winhttp show proxy
    out = _read_cmd_output(["netsh", "winhttp", "show", "proxy"])
    if out:
        # 直接接続判定（英/日 両対応）
        if ("Direct access" in out) or ("直接アクセス" in out):
            pass  # 何もしない（次の手段へ）
        else:
            # "Proxy Server(s) : ..." または "プロキシ サーバー : ..."
            m = re.search(r"Proxy Server(?:\(s\))?\s*:\s*(.+)", out, flags=re.I)
            if not m:
                m = re.search(r"プロキシ\s*サーバー\s*:\s*(.+)", out)
            if m:
                value = m.group(1).strip()
                picked = _extract_mapping_value(value, prefer=prefer)
                if picked:
                    return _normalize_to_http_scheme(picked)

    # 2) レジストリ（WinINET / IE）
    # ProxyEnable が 1 のときのみ ProxyServer を採用
    en_out = _read_cmd_output([
        "reg", "query",
        r"HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        "/v", "ProxyEnable"
    ])
    enabled = False
    if en_out:
        # 行末の 0x1 / 0x0 を拾う
        m = re.search(r"\b0x([0-9a-fA-F]+)\b", en_out)
        if m and int(m.group(1), 16) == 1:
            enabled = True

    if enabled:
        sv_out = _read_cmd_output([
            "reg", "query",
            r"HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings",
            "/v", "ProxyServer"
        ])
        if sv_out:
            # 値は 'REG_SZ    http=host:port;https=...' など
            m = re.search(r"ProxyServer\s+REG_\w+\s+(.+)", sv_out)
            if m:
                value = m.group(1).strip()
                picked = _extract_mapping_value(value, prefer=prefer)
                if picked:
                    return _normalize_to_http_scheme(picked)

    # 3) 環境変数（フォールバック）
    if allow_env_fallback:
        env = os.environ
        cand = None
        if prefer == "https":
            cand = env.get("HTTPS_PROXY") or env.get("https_proxy") \
                   or env.get("HTTP_PROXY") or env.get("http_proxy")
        else:
            cand = env.get("HTTP_PROXY") or env.get("http_proxy") \
                   or env.get("HTTPS_PROXY") or env.get("https_proxy")
        if cand:
            return _normalize_to_http_scheme(cand)

    return None

# 使い方例:
if __name__ == "__main__":
    p = get_proxy_from_cmd(prefer="http", allow_env_fallback=True)
    print("Detected proxy:", p)
