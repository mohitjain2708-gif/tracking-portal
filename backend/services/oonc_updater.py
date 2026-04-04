from datetime import datetime


def _date_only(value: str) -> str:
    if not value:
        return ""
    return str(value)[:10]


def _parse_date(date_str: str):
    if not date_str:
        return None

    for fmt in ("%d-%m-%Y", "%d-%m-%Y %H:%M:%S", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except Exception:
            pass

    # fallback for strings like 2026-03-25T00:00:00
    try:
        core = str(date_str).strip().replace("T", " ")[:19]
        return datetime.strptime(core, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def _days_since(date_str: str):
    dt = _parse_date(date_str)
    if not dt:
        return ""
    return (datetime.now().date() - dt.date()).days


def _pick_col(hmap: dict, names: list[str]):
    for n in names:
        if n.lower() in hmap:
            return hmap[n.lower()]
    return None


def update_oonc_sheet(ws, header_row: int, hmap: dict, tracking_map: dict):
    container_col = _pick_col(hmap, ["Container No", "ContainerNo", "Container", "CNTR NO"])
    train_no_col = _pick_col(hmap, ["Train No"])
    rail_dep_col = _pick_col(hmap, ["Rail Departure"])
    gate_col = _pick_col(hmap, ["Gate in Birganj", "Gate in Birgunj"])
    transit_col = _pick_col(hmap, ["Rail Transit Time"])
    port_since_col = _pick_col(hmap, ["At India Port Since", "Port Since"])
    last_loc_col = _pick_col(hmap, ["LAST LOCATION", "Last Location"])
    last_loc_date_col = _pick_col(
        hmap,
        ["LAST LOCATION (DATE)", "LAST LOCATION (TIME)", "Last Location (Date)"],
    )

    if not container_col:
        raise ValueError("Container number column not found in OONC")

    for r in range(header_row + 1, ws.max_row + 1):
        container_no = str(ws.cell(r, container_col).value or "").strip().upper()
        if not container_no:
            continue

        tr = tracking_map.get(container_no, {})

        latest_location = tr.get("latest_location", "") or ""
        latest_time_full = tr.get("latest_time", "") or ""
        latest_time_date = _date_only(latest_time_full)
        rail_departure = tr.get("rail_departure", "") or ""
        train_no = tr.get("train_no", "") or ""

        arrived_birgunj = "BIRGUNJ" in latest_location.upper() or "BIRGANJ" in latest_location.upper()
        gate_date = latest_time_date if arrived_birgunj else ""

        # Train No
        if train_no_col:
            ws.cell(r, train_no_col).value = train_no

        # Rail Departure
        if rail_dep_col:
            ws.cell(r, rail_dep_col).value = rail_departure

        # Gate in Birganj
        if gate_col:
            ws.cell(r, gate_col).value = gate_date

        # Rail Transit Time
        if transit_col:
            if gate_date:
                ws.cell(r, transit_col).value = "Arrived"
            elif rail_departure:
                transit_days = _days_since(rail_departure)
                ws.cell(r, transit_col).value = transit_days if transit_days != "" else ""
            else:
                ws.cell(r, transit_col).value = "Not Railed"

        # At India Port Since
        if port_since_col:
            if gate_date:
                ws.cell(r, port_since_col).value = ""
            elif latest_time_date:
                port_days = _days_since(latest_time_date)
                ws.cell(r, port_since_col).value = port_days if port_days != "" else ""
            else:
                ws.cell(r, port_since_col).value = ""

        # LAST LOCATION
        if last_loc_col:
            ws.cell(r, last_loc_col).value = latest_location

        # LAST LOCATION (DATE)
        if last_loc_date_col:
            ws.cell(r, last_loc_date_col).value = latest_time_date