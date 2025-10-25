import re
import win32com.client
from typing import Optional
from urllib.parse import urlparse, urljoin

def _strip_scheme(
        s: str
    ) -> str:
    return re.sub(r'^[a-zA-Z][a-zA-Z0-9+.-]*://', '', s)

def _safe_set_headers(
        http,
        hdrs: dict
    ) -> None:
    for k, v in (hdrs or {}).items():
        if not k or v is None:
            continue
        # 簡易チェック：制御文字や改行を含むヘッダは弾く
        if any(c in str(k) for c in ("\r", "\n", "\x00")):
            continue
        try:
            http.setRequestHeader(str(k), str(v))
        except Exception:
            # 一部ヘッダ（Host 等）は setRequestHeader 非対応なのでスキップ
            continue

def _msxml2_get_http_object(
        *,
        progids: tuple[str, ...] = (
            "MSXML2.ServerXMLHTTP.6.0",
            "MSXML2.ServerXMLHTTP.3.0",
            "MSXML2.ServerXMLHTTP",
        ),
        timeouts_ms: tuple[int, int, int, int] | None = (5000, 10000, 10000, 180000),
        insecure_ssl: bool = False,
    ) -> None:
    SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2
    SXH_SERVER_CERT_IGNORE_UNKNOWN_CA = 0x00000100
    SXH_SERVER_CERT_IGNORE_CERT_CN_INVALID = 0x00001000
    SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID = 0x00000200

    attempts: list[str] = []
    last_exc: Exception | None = None
    obj = None

    for pid in progids:
        try:
            obj = win32com.client.Dispatch(pid)
            attempts.append(f"OK {pid}")
            break
        except Exception as e:
            attempts.append(f"NG {pid}: {e!r}")
            last_exc = e
            continue

    if obj is None:
        detail = " | ".join(attempts) or "(no attempts)"
        raise RuntimeError(f"MSXML2.ServerXMLHTTP を生成できません。詳細: {detail}") from last_exc

    if timeouts_ms is not None:
        try:
            resolve, connect, send, receive = timeouts_ms
            obj.setTimeouts(resolve, connect, send, receive)
        except Exception as e:
            attempts.append(f"warn setTimeouts: {e!r}")

    if insecure_ssl:
        try:
            ignore_flags = (
                SXH_SERVER_CERT_IGNORE_UNKNOWN_CA
                | SXH_SERVER_CERT_IGNORE_CERT_CN_INVALID
                | SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID
            )
            obj.setOption(SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, ignore_flags)
        except Exception as e:
            attempts.append(f"warn setOption(ignore_ssl): {e!r}")
    return obj

def msxml2_request(
        method: str,
        url: str,
        headers: dict,
        timeout_ms: tuple[int,int,int,int],
        proxy_hostport: str,
    ) -> None:
    # --- 既定: Accept-Encoding を identity に（未指定時のみ） ---
    hdrs = dict(headers or {})
    if not any(h.lower() == "accept-encoding" for h in hdrs.keys()):
        hdrs["Accept-Encoding"] = "identity"

    # --- プロキシ整形 ---
    proxy = None
    if proxy_hostport:
        proxy = proxy_hostport if "://" in proxy_hostport else f"http://{proxy_hostport}"

    # --- リダイレクト制御 ---
    max_redirects = 10
    redirects_done = 0
    current_method = (method or "GET").upper()
    current_url = url

    last_http = None
    last_status = None
    last_location = None

    while True:
        http = _msxml2_get_http_object()
        try:
            http.setTimeouts(*timeout_ms)
        except Exception:
            pass

        # プロキシ
        try:
            SXH_PROXY_SET_DIRECT = 1
            SXH_PROXY_SET_PROXY  = 2
            if proxy:
                http.setProxy(SXH_PROXY_SET_PROXY, _strip_scheme(proxy))
            else:
                http.setProxy(SXH_PROXY_SET_DIRECT, "")
        except Exception:
            pass

        # 送信準備
        http.open(current_method, current_url, False)
        _safe_set_headers(http, hdrs)

        # 送信
        http.send()

        status = None
        try:
            status = int(http.status)
        except Exception:
            status = None

        # リダイレクト判定
        if status in (301, 302, 303, 307, 308):
            try:
                resp_headers = msxml2_all_headers_dict(http)
                loc = resp_headers.get("location") or resp_headers.get("Location")
            except Exception:
                loc = None

            if loc and redirects_done < max_redirects:
                base = current_url
                new_url = urljoin(base, loc)
                if status == 303:
                    current_method = "GET"
                if not any(k.lower() == "referer" for k in hdrs.keys()):
                    hdrs["Referer"] = base
                try:
                    if urlparse(new_url).netloc != urlparse(base).netloc:
                        if any(k.lower() == "range" for k in list(hdrs.keys())):
                            hdrs.pop("Range", None)
                except Exception:
                    pass

                redirects_done += 1
                current_url = new_url
                last_http = http
                last_status = status
                last_location = loc
                continue
            else:
                return http
        return http
    # ここには来ないが、保険で最後の状態を吐く
    raise RuntimeError(f"MSXML2 request failed (url={current_url}, status={last_status}, location={last_location})")

def msxml2_read_body_bytes(
        http,
    ) -> bytes:
    rb = http.responseBody
    try:
        return bytes(rb)
    except TypeError:
        return bytes(bytearray(rb))

def msxml2_all_headers_dict(
        http,
    ) -> dict:
    raw = http.getAllResponseHeaders() or ""
    hdrs = {}
    for line in raw.splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            hdrs[k.strip().lower()] = v.strip()
    return hdrs

def msxml2_available(
    ) -> bool:
    return all(
        name in globals()
        for name in ("msxml2_request", "msxml2_all_headers_dict", "msxml2_read_body_bytes")
    )

def probe_remote_msxml2(
        url: str,
        headers: dict,
        timeouts_ms: tuple[int,int,int,int],
        proxy: Optional[str],
    ) -> None:
    total_size = None
    accept_ranges = False
    if_range_token = None
    try:
        http = msxml2_request("HEAD", url, headers, timeouts_ms, proxy)
        status = int(http.status)
        if 200 <= status < 400:
            hdrs = msxml2_all_headers_dict(http)
            cl = hdrs.get("content-length")
            if cl and cl.isdigit():
                total_size = int(cl)
            accept_ranges = (hdrs.get("accept-ranges", "").lower() == "bytes")
            etag = hdrs.get("etag")
            last_modified = hdrs.get("last-modified")
            if_range_token = etag or last_modified
    except Exception:
        pass
    return total_size, accept_ranges, if_range_token