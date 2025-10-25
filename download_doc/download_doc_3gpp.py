import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import os
from pathlib import Path
from urllib.parse import urlparse
from typing import Optional,List,Dict,Any

from pure_download.download_file import (
    download_file_safely_msxml2,
    )

from pure_download.download_util import (
    get_landing_and_session,
    )

from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
    )

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor
import queue, threading, requests, pythoncom
from requests.adapters import HTTPAdapter

try:
    from urllib3.util.retry import Retry
except Exception:
    from requests.packages.urllib3.util.retry import Retry  # type: ignore

UA_STR = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"  # ← 全エージェントで統一

def _normalize_proxy(p: Optional[str]) -> Optional[str]:
    if not p: return None
    p = p.strip()
    return p if "://" in p else f"http://{p}"

def _make_agent_session(kind: str, proxy: Optional[str], agent_id: int,
                        pool: int = 64, retries: int = 2) -> Tuple[str, requests.Session]:
    landing_map = {"3gpp":"https://mentor.3gpp.org/802.11","3gpp":"https://www.3gpp.org/ftp"}
    landing = landing_map[(kind or "").strip().lower()]

    s = requests.Session()
    retry = Retry(total=retries, connect=retries, read=retries,
                  backoff_factor=0.5, status_forcelist=(500,502,503,504),
                  allowed_methods=False)
    ad = HTTPAdapter(pool_connections=pool, pool_maxsize=pool, max_retries=retry)
    s.mount("http://", ad); s.mount("https://", ad)

    # UA は固定。識別は別ヘッダーで。
    s.headers.update({
        "User-Agent": UA_STR,
        "X-Agent-ID": str(agent_id),    # ← サーバには影響小、ログ識別用
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9,ja;q=0.8",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    })
    s.trust_env = False
    prx = _normalize_proxy(proxy)
    if prx: s.proxies.update({"http": prx, "https": prx})

    try: s.get(landing, timeout=15)
    except Exception: pass
    return landing, s

def fetch_3gpp_docs_queue(base_path: str, download_urls: List[str], proxy: Optional[str]) -> List[Dict[str, Any]]:
    create_subfolder_when_absent(Path(base_path), "DOCS")   # type: ignore[name-defined]
    download_path = Path(base_path) / "DOCS"

    # ===== セッション取得 =====
    LANDING, sess = get_landing_and_session("3gpp")  # type: ignore[name-defined]

    total = len(download_urls)
    if total == 0:
        print(f"{emo.warn} ダウンロード対象がありません。")  # type: ignore[name-defined]
        return []


    tasks: List[Dict[str, Any]] = []
    for i,u in enumerate(download_urls, start=1):
        parsed = urlparse(u)
        pure_filename = os.path.basename(parsed.path) or "download.bin"
        stem = os.path.splitext(pure_filename)[0] or "download"
        ext_guess = os.path.splitext(pure_filename)[1] or ".bin"

        # pure_filename = stem + ".zip"
        tasks.append({"index":i, "url":str(u).strip(),
                      "pure_filename":pure_filename,"stem":stem,"ext_guess":ext_guess})

    results: List[Dict[str, Any]] = []
    results_lock = threading.Lock()
    q: "queue.Queue[Dict[str, Any]]" = queue.Queue()

    # 既存スキップ、未DLをキューへ
    for t in tasks:
        target_zip = download_path / t["pure_filename"]
        if target_zip.exists():
            pct = round(t["index"]/total*100)
            print(f"[Agent-0 {t['index']}/{total} {pct}%] ✅ 既存: {target_zip.name}")
            results.append({
                "index": t["index"], "url": t["url"],
                "filename": t["pure_filename"], "download_path": str(download_path),
                "name": t["pure_filename"], "saved_path": str(target_zip),
                "ext": t["ext_guess"], "skipped": True, "error": None
            })
        else:
            q.put(t)

    if q.empty():
        results.sort(key=lambda r: r.get("index",0)); return results

    MAX_AGENTS = 12
    agent_pool = [_make_agent_session("3gpp", proxy, agent_id=i+1) for i in range(MAX_AGENTS)]

    def agent_run(agent_slot: int):
        # ★ COM を使うなら各スレッドで初期化（MSXML2対策）
        pythoncom.CoInitialize()
        try:
            landing, sess = agent_pool[agent_slot]
            while True:
                try:
                    t = q.get_nowait()
                except queue.Empty:
                    break
                i = t["index"]; pct = round(i/total*100)
                try:
                    ext = download_file_safely_msxml2(  # type: ignore[name-defined]
                        t["url"], download_path, t["stem"],
                        session=sess, referer=landing, proxy=_normalize_proxy(proxy),
                        connect_timeout=10, read_timeout=180, max_retries=5,
                        use_curl_fallback=True,
                        # もしこの関数が UA 指定を受けるなら ↓ を渡す（内部でも UA を固定化）
                        # user_agent=UA_STR,
                    )
                    saved = download_path / f"{t['stem']}{ext or '.zip'}"
                    print(f"[Agent-{agent_slot+1} {i}/{total} {pct}%] ✅ ext:{ext or '(不明)'} → {saved.name}")
                    item = {
                            "index": t["index"], "url": t["url"],
                            "filename": t["pure_filename"], "download_path": str(download_path),
                            "name": t["pure_filename"], "saved_path": str(saved),
                            "ext": t["ext_guess"], "skipped": False, "error": None
                    }
                except Exception as e:
                    print(f"[Agent-{agent_slot+1} {i}/{total} {pct}%] ⚠️ エラー: {e}")
                    item = {
                            "index": t["index"], "url": t["url"],
                            "filename": t["pure_filename"], "download_path": str(download_path),
                            "name": t["pure_filename"], "saved_path": str(saved),
                            "ext": t["ext_guess"], "skipped": False, "error": str(e)
                    }
                finally:
                    with results_lock:
                        results.append(item)
                    q.task_done()
        finally:
            pythoncom.CoUninitialize()
            try: sess.close()
            except Exception: pass

    with ThreadPoolExecutor(max_workers=MAX_AGENTS, thread_name_prefix="agent") as ex:
        futs = [ex.submit(agent_run, slot) for slot in range(MAX_AGENTS)]
        for f in futs: f.result()

    results.sort(key=lambda r: r.get("index",0))
    return results
