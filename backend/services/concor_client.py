import requests
from datetime import datetime
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

CONCOR_API_URL = "https://www.concorindia.co.in/api/multipalContainer"

_session = requests.Session()
_retries = Retry(
    total=2,
    backoff_factor=0.4,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["POST"],
)
_adapter = HTTPAdapter(pool_connections=50, pool_maxsize=50, max_retries=_retries)
_session.mount("https://", _adapter)
_session.mount("http://", _adapter)


def _date_only(value: str) -> str:
    if not value:
        return ""
    try:
        dt = datetime.strptime(value, "%d/%m/%Y %H:%M:%S")
        return dt.strftime("%d-%m-%Y")
    except Exception:
        return value[:10]


def fetch_concor(container_no: str) -> dict:
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0",
        "Origin": "https://www.concorindia.co.in",
        "Referer": "https://www.concorindia.co.in/track-n-trace?lang=en",
    }

    resp = _session.post(
        CONCOR_API_URL,
        headers=headers,
        json={"containerNo": [container_no]},
        timeout=(5, 12),
    )
    resp.raise_for_status()
    data = resp.json()
    track = (((data or {}).get("data") or {}).get(container_no) or {}).get("containerTrack", {}) or {}

    return {
        "train_no": track.get("TRAIN_NUMBER", "") or "",
        "rail_departure": _date_only(track.get("DEPARTURE_DATE_&_TIME", "") or ""),
        "wagon_no": track.get("WAGON_NUMBER", "") or "",
        "train_origin": track.get("TRAIN_ORIGNATING_STATION", "") or "",
        "train_dest": track.get("TRAIN_DESTINATION_STATION", "") or "",
        "shipping_line": track.get("SHIPPING_LINE", "") or "",
    }