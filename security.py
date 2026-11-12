import os

from fastapi import HTTPException, Request


def _load_allowed_ips() -> set[str]:
    """ALLOWED_IP_ADDRESSES（カンマ区切り）から許可IP集合を作る"""
    raw = os.environ.get("ALLOWED_IP_ADDRESSES", "")
    return {ip.strip() for ip in raw.split(",") if ip.strip()}


ALLOWED_IP_ADDRESSES = _load_allowed_ips()


def _get_client_ip(request: Request) -> str:
    """Vercelが付与するヘッダーからクライアントIPを取得する"""
    forwarded = request.headers.get("x-forwarded-for", "")
    if forwarded:
        # 先頭がクライアントIP
        return forwarded.split(",")[0].strip()
    if request.client is None:
        return ""
    return request.client.host


def require_allowed_ip(request: Request) -> None:
    """許可リストにないIPからのリクエストを403で拒否する"""
    if _get_client_ip(request) in ALLOWED_IP_ADDRESSES:
        return
    raise HTTPException(status_code=403, detail="Forbidden")
