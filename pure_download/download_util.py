import re
import requests
from typing import Optional
from urllib.parse import urlparse
from pathlib import PureWindowsPath

def to_double_backslash_literal(
        p: str,
    ) -> str:
    if p is None:
        return ""
    s = str(p).strip()
    m = re.match(r'^\s*(?:Pure)?Windows?Path\s*\(\s*[rR]?[\'"](.+)[\'"]\s*\)\s*$', s)
    if m:
        s = m.group(1)
    s = str(PureWindowsPath(s))
    return s.replace("\\", "\\\\")

def get_landing_and_session(
        kind: str,
    )-> tuple[str, requests.Session]:
    
    kind_norm = (kind or "").strip().lower()
    landing_map = {
        "ieee": "https://mentor.ieee.org/802.11",
        "3gpp": "https://www.3gpp.org/ftp",
    }
    if kind_norm not in landing_map:
        raise ValueError('kind は "IEEE" か "3gpp" を指定してください。')

    LANDING = landing_map[kind_norm]

    sess = requests.Session()
    sess.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9,ja;q=0.8",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    })

    try:
        sess.get(LANDING, timeout=30)  # Cookie 獲得
    except Exception:
        pass

    return LANDING, sess

def normalize_proxy_for_msxml2(
        p: Optional[str],
    ) -> Optional[str]:
    if not p:
        return None
    return p if "://" in p else f"http://{p}"

def cookie_header_from_session(
        session,
        url: str,
    ) -> str:
    if session is None or not getattr(session, "cookies", None):
        return ""
    host = urlparse(url).hostname or ""
    path = urlparse(url).path or "/"
    pairs = []
    try:
        for c in session.cookies:
            if c.domain and not (host == c.domain or host.endswith("." + c.domain.lstrip("."))):
                continue
            if c.path and not path.startswith(c.path):
                continue
            if c.secure and not url.lower().startswith("https://"):
                continue
            if c.name and (c.value is not None):
                pairs.append(f"{c.name}={c.value}")
    except Exception:
        pass
    return "; ".join(pairs)

def sanitize_filename(
        name: str,
    ) -> str:
    for ch in r'<>:"/\|?*':
        name = name.replace(ch, "_")
    name = name.strip().rstrip(".")
    return name or "page"

def is_dir_like(
        path: str,
    ) -> bool:
    return (os.path.isdir(path) or path.endswith(("\\", "/")))
