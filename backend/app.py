import os
import tempfile

from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from services.excel_parser import load_workbook_and_oonc, detect_header_row, header_map
from services.ldb_client import fetch_ldb
from services.concor_client import fetch_concor
from services.oonc_updater import update_oonc_sheet


app = FastAPI(title="Tracking Portal API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/api/process-tracking")
async def process_tracking(file: UploadFile = File(...)):
    suffix = os.path.splitext(file.filename)[1] or ".xlsx"

    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, f"input{suffix}")
        out_path = os.path.join(tmpdir, f"processed{suffix}")

        with open(in_path, "wb") as f:
            f.write(await file.read())

        wb, ws = load_workbook_and_oonc(in_path)
        hdr_row = detect_header_row(ws)
        hmap = header_map(ws, hdr_row)

        container_col = None
        for key in ["container no", "containerno", "container", "cntr no"]:
            if key in hmap:
                container_col = hmap[key]
                break

        if not container_col:
            return JSONResponse(
                {"error": "Container number column not found in OONC"},
                status_code=400,
            )

        tracking_map = {}
        preview_rows = []

        for r in range(hdr_row + 1, ws.max_row + 1):
            container_no = str(ws.cell(r, container_col).value or "").strip().upper()
            if not container_no:
                continue
            if container_no in tracking_map:
                continue

            ldb = {}
            concor = {}
            errors = []

            try:
                ldb = fetch_ldb(container_no)
            except Exception as e:
                errors.append(f"LDB: {e}")

            try:
                concor = fetch_concor(container_no)
            except Exception as e:
                errors.append(f"CONCOR: {e}")

            tracking_map[container_no] = {**ldb, **concor, "error": " | ".join(errors)}

        update_oonc_sheet(ws, hdr_row, hmap, tracking_map)
        wb.save(out_path)

        headers = [str(ws.cell(hdr_row, c).value or "") for c in range(1, ws.max_column + 1)]

        for r in range(hdr_row + 1, ws.max_row + 1):
            row = {}
            is_blank = True
            for c in range(1, ws.max_column + 1):
                value = ws.cell(r, c).value
                row[headers[c - 1] or f"Column {c}"] = "" if value is None else value
                if value not in [None, ""]:
                    is_blank = False
            if not is_blank:
                preview_rows.append(row)

        return {
            "headers": headers,
            "rows": preview_rows,
            "tracked_containers": len(tracking_map),
            "download_ready": True,
        }