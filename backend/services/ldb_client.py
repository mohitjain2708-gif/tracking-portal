import requests
from datetime import datetime
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

LDB_API_URL = "https://www.ldb.co.in/api/ldb/container/search"

_session = requests.Session()
_retries = Retry(
    total=2,
    backoff_factor=0.4,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"],
)
_adapter = HTTPAdapter(pool_connections=50, pool_maxsize=50, max_retries=_retries)
_session.mount("https://", _adapter)
_session.mount("http://", _adapter)


def _fmt_iso(ts: str) -> str:
    if not ts:
        return ""
    try:
        core = ts[:19].replace("T", " ")
        dt = datetime.strptime(core, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        return ts


def classify_ldb_rail_status(location: str, event: str) -> str:
    text = f"{location} {event}".upper()
    if "BIRGUNJ" in text or "BIRGANJ" in text:
        return "Arrived Birgunj"
    if "DANGOAPOSI" in text or "STATION CROSSED" in text or "RAIL" in text:
        return "On Rail"
    if "ICD IN" in text and "VISHAKAPATNAM" in text:
        return "At Vizag ICD"
    if "CFS" in text and ("VIZAG" in text or "VISHAKAPATNAM" in text):
        return "At Vizag CFS"
    if "PORT" in text and ("VISAKHA" in text or "VISHAKAPATNAM" in text):
        return "At Port"
    return "In Transit"


def fetch_ldb(container_no: str) -> dict:
    params = {"cntrNo": container_no, "searchType": "39"}
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
        "Referer": f"https://www.ldb.co.in/ldb/containersearch/39/{container_no}",
    }

    resp = _session.get(LDB_API_URL, params=params, headers=headers, timeout=(5, 12))
    resp.raise_for_status()
    data = resp.json()

    obj = data.get("object", {}) or {}
    last_event = obj.get("lastEvent", {}) or {}

    location = last_event.get("currentLocation", "") or ""
    event = last_event.get("eventName", "") or ""
    latest_time = _fmt_iso(last_event.get("timestampTimezone", "") or "")

    return {
        "latest_location": location,
        "latest_event": event,
        "latest_time": latest_time,
        "ldb_rail_status": classify_ldb_rail_status(location, event),
    }