from __future__ import annotations

import io
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from pydantic import BaseModel, Field

BASE_DIR = Path(__file__).resolve().parent
app = FastAPI(title="RouteGuard", version="1.0.0", description="Excel-based wrong-route cage scan analysis demo")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")

# This is a single-user demonstration app. State is held in memory and is cleared when
# the server restarts. Do not deploy with real operational data without an approved review.
STORE: dict[str, Any] = {
    "loaded": False,
    "filename": None,
    "file_type": None,
    "file_bytes": None,
    "sheet_name": None,
    "df": None,
    "columns": {},
    "records": [],
    "events": [],
    "mode": "empty",
    "uploaded_at": None,
    "last_refresh": None,
}

ALIASES = {
    "expected": [
        "originalscanlocation", "originallocation", "originalscanloc", "originalscannedlocation",
        "originalstaginglocation", "originalarea", "originalscanarea",
        "expectedlocation", "expectedarea", "assignedarea", "assignedlocation",
        "destinationlocation", "routelocation", "routearea", "correctlocation",
        "plannedlocation", "expectedstaginglocation", "assignedstaginglocation",
        "expectedstagingarea", "assignedstagingarea", "expectedscanlocation", "destinationarea",
    ],
    "actual": [
        "actualscanlocation", "actualstaginglocation", "scanlocation", "scannedlocation",
        "actualarea", "scanarea", "stagelocation", "actualstagingarea", "scannedat",
        "actual", "currentlocation", "observedlocation", "scanloc", "locationactual",
    ],
    "cage": ["cageid", "cage", "containerid", "container", "cagebarcode", "barcode", "cartid", "toteid"],
    "route": ["routeid", "route", "routenumber", "routecode", "deliveryroute"],
    "wave": ["wave", "waveid", "wavenumber", "dispatchwave"],
    "success": ["issuccess", "success", "isresolved"],
}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _norm(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def _find_column(columns: list[Any], role: str) -> Any | None:
    normalized = {_norm(col): col for col in columns}
    for alias in ALIASES[role]:
        if alias in normalized:
            return normalized[alias]
    # Secondary match for familiar variants such as "Original Scan Location (Area)".
    for col in columns:
        n = _norm(col)
        if role == "expected" and (n.startswith("original") or n.startswith("expected") or n.startswith("assigned")) and any(x in n for x in ["location", "area", "stage", "scan"]):
            return col
        if role == "actual" and any(x in n for x in ["actual", "scan", "scanned"]) and any(x in n for x in ["location", "area", "stage"]):
            return col
    return None


def _truthy(value: Any) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return str(value).strip().lower() in {"true", "1", "yes", "y", "success", "solved", "resolved", "done", "complete", "completed"}


def _text(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_wrong(expected: Any, actual: Any) -> bool | None:
    exp, act = _text(expected), _text(actual)
    if not exp or not act:
        return None
    # Original Scan Location is the reference value. Compare exact normalized locations.
    # Only case and repeated/leading/trailing whitespace are ignored.
    exp_clean = re.sub(r"\s+", " ", exp.casefold()).strip()
    act_clean = re.sub(r"\s+", " ", act.casefold()).strip()
    return exp_clean != act_clean


def _records_from_df(df: pd.DataFrame, columns: dict[str, Any], preserve_states: dict[str, dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    expected_col, actual_col = columns["expected"], columns["actual"]
    success_col = columns.get("success")
    for index, row in df.iterrows():
        # DataFrame index is retained internally so the matching workbook row can be updated.
        expected = row.get(expected_col, "")
        actual = row.get(actual_col, "")
        wrong = _is_wrong(expected, actual)
        row_id = str(index)
        old = (preserve_states or {}).get(row_id, {})
        raw_success = row.get(success_col, False) if success_col is not None else (wrong is False)
        success = _truthy(raw_success)
        if wrong is None:
            status = "unverified"
        elif not wrong:
            status = "correct"
            # A correct scan is a successful record; initialize empty/missing values to true.
            if success_col is not None and pd.isna(raw_success):
                success = True
        elif success:
            status = "resolved"
        else:
            status = old.get("status", "open")
            if status not in {"open", "acknowledged"}:
                status = "open"

        cage_col, route_col, wave_col = columns.get("cage"), columns.get("route"), columns.get("wave")
        cage = _text(row.get(cage_col, "")) if cage_col is not None else ""
        route = _text(row.get(route_col, "")) if route_col is not None else ""
        wave = _text(row.get(wave_col, "")) if wave_col is not None else ""
        if not cage:
            cage = f"Row {int(index) + 2 if isinstance(index, int) else index + 2}"
        record = {
            "id": row_id,
            "excel_row": int(index) + 2 if isinstance(index, int) else int(index) + 2,
            "cage_id": cage,
            "route_id": route,
            "wave": wave,
            "expected_location": _text(expected),
            "actual_location": _text(actual),
            "is_wrong": wrong is True,
            "unverified": wrong is None,
            "isSuccess": success,
            "status": status,
        }
        records.append(record)
    return records


def _initialize_success_column(df: pd.DataFrame, columns: dict[str, Any], records: list[dict[str, Any]]) -> tuple[pd.DataFrame, dict[str, Any]]:
    actual_success_col = columns.get("success")
    created_column = actual_success_col is None
    if actual_success_col is None:
        actual_success_col = "isSuccess"
        # Prevent a duplicate if a non-exact but equivalent header is present.
        existing = next((c for c in df.columns if _norm(c) == "issuccess"), None)
        if existing is not None:
            actual_success_col = existing
        else:
            df[actual_success_col] = False
        columns["success"] = actual_success_col
    for rec in records:
        idx = int(rec["id"])
        original = df.at[idx, actual_success_col]
        if created_column:
            # Initialize correct records to TRUE and mismatches/unverified records to FALSE.
            initialized = bool(not rec["is_wrong"] and not rec["unverified"])
            df.at[idx, actual_success_col] = initialized
            rec["isSuccess"] = initialized
        elif pd.isna(original):
            initialized = bool(not rec["is_wrong"] and not rec["unverified"])
            df.at[idx, actual_success_col] = initialized
            rec["isSuccess"] = initialized
        else:
            df.at[idx, actual_success_col] = _truthy(original)
            rec["isSuccess"] = _truthy(original)
        if rec["is_wrong"]:
            rec["status"] = "resolved" if rec["isSuccess"] else "open"
    return df, columns


def _set_store(df: pd.DataFrame, filename: str, file_type: str, file_bytes: bytes, sheet_name: str | None,
               columns: dict[str, Any], mode: str = "uploaded", previous_events: list[dict[str, Any]] | None = None) -> None:
    success_col = columns.get("success")
    records = _records_from_df(df, columns)
    df, columns = _initialize_success_column(df, columns, records)
    # Rebuild from the initialized values, then retain actionable acknowledgement state where possible.
    records = _records_from_df(df, columns)
    # For incorrect rows with false isSuccess, use open status at upload time.
    STORE.update({
        "loaded": True,
        "filename": filename,
        "file_type": file_type,
        "file_bytes": file_bytes,
        "sheet_name": sheet_name,
        "df": df,
        "columns": columns,
        "records": records,
        "events": previous_events or [],
        "mode": mode,
        "uploaded_at": _now(),
        "last_refresh": _now(),
    })


def _state() -> dict[str, Any]:
    if not STORE["loaded"]:
        return {"loaded": False, "mode": "empty", "message": "Upload an Excel or CSV scan report, or load sample data.", "last_refresh": _now()}
    records = STORE["records"]
    mismatches = [r for r in records if r["is_wrong"]]
    groups_by_location: dict[str, list[dict[str, Any]]] = {}
    for record in mismatches:
        key = record["actual_location"] or "Unknown location"
        groups_by_location.setdefault(key, []).append(record)
    groups = []
    for location, rows in sorted(groups_by_location.items(), key=lambda item: item[0].lower()):
        groups.append({
            "actual_location": location,
            "count": len(rows),
            "open": sum(r["status"] == "open" for r in rows),
            "acknowledged": sum(r["status"] == "acknowledged" for r in rows),
            "resolved": sum(r["status"] == "resolved" for r in rows),
            "rows": rows,
        })
    return {
        "loaded": True,
        "filename": STORE["filename"],
        "mode": STORE["mode"],
        "source_type": STORE["file_type"],
        "sheet_name": STORE["sheet_name"],
        "uploaded_at": STORE["uploaded_at"],
        "last_refresh": STORE["last_refresh"],
        "auto_refresh_seconds": 5,
        "totals": {
            "records": len(records),
            "correct": sum(r["status"] == "correct" for r in records),
            "wrong": len(mismatches),
            "open": sum(r["status"] == "open" for r in mismatches),
            "acknowledged": sum(r["status"] == "acknowledged" for r in mismatches),
            "resolved": sum(r["status"] == "resolved" for r in mismatches),
            "unverified": sum(r["unverified"] for r in records),
        },
        "groups": groups,
        "records": records,
        "events": list(reversed(STORE["events"][-30:])),
        "columns": {key: (str(value) if value is not None else None) for key, value in STORE["columns"].items()},
        "refresh_note": "Refresh updates the portal view and action status. New scans require a newly uploaded report or an approved live data feed.",
    }


def _read_uploaded(filename: str, raw: bytes) -> tuple[pd.DataFrame, str | None, str]:
    suffix = Path(filename).suffix.lower()
    try:
        if suffix == ".csv":
            df = pd.read_csv(io.BytesIO(raw), dtype=object, keep_default_na=True)
            return df, None, "csv"
        if suffix in {".xlsx", ".xlsm"}:
            bio = io.BytesIO(raw)
            df = pd.read_excel(bio, sheet_name=0, dtype=object, engine="openpyxl")
            wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            sheet_name = wb.sheetnames[0]
            wb.close()
            return df, sheet_name, "xlsx"
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read this file. Check that it is a valid Excel/CSV file. Details: {exc}") from exc
    raise HTTPException(status_code=400, detail="Upload an .xlsx, .xlsm, or .csv file.")


def _excel_bytes() -> bytes:
    df: pd.DataFrame = STORE["df"]
    filename = STORE["filename"] or "routeguard_export.xlsx"
    if STORE["file_type"] == "csv":
        return df.to_csv(index=False).encode("utf-8-sig")

    source = STORE["file_bytes"]
    try:
        wb = load_workbook(io.BytesIO(source))
        sheet_name = STORE["sheet_name"] or wb.sheetnames[0]
        ws = wb[sheet_name]
        header_values = [cell.value for cell in ws[1]]
        success_col_idx = next((i + 1 for i, val in enumerate(header_values) if _norm(val) == "issuccess"), None)
        if success_col_idx is None:
            success_col_idx = ws.max_column + 1
            header_cell = ws.cell(row=1, column=success_col_idx, value="isSuccess")
            if ws.cell(row=1, column=max(1, success_col_idx - 1)).has_style:
                previous = ws.cell(row=1, column=max(1, success_col_idx - 1))
                header_cell._style = previous._style
                header_cell.number_format = previous.number_format
            header_cell.font = Font(name=header_cell.font.name or "Aptos", bold=True, color="FFFFFF")
            header_cell.fill = PatternFill("solid", fgColor="243B53")
        # Update all rows against dataframe order, with header on row 1.
        success_col_name = STORE["columns"].get("success", "isSuccess")
        for offset, rec in enumerate(STORE["records"], start=2):
            val = STORE["df"].iloc[offset - 2].get(success_col_name, rec["isSuccess"])
            ws.cell(row=offset, column=success_col_idx, value=_truthy(val))
        out = io.BytesIO()
        wb.save(out)
        wb.close()
        return out.getvalue()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Could not create the updated workbook: {exc}") from exc


class ActionRequest(BaseModel):
    action: str = Field(..., pattern="^(acknowledge|resolve|reopen)$")
    verified: bool = False
    note: str = Field(default="", max_length=300)


@app.get("/", response_class=HTMLResponse)
def home() -> FileResponse:
    return FileResponse(BASE_DIR / "static" / "index.html")


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {"ok": True, "service": "RouteGuard demo", "station_command_connected": False, "time": _now()}


@app.get("/api/state")
def get_state() -> dict[str, Any]:
    STORE["last_refresh"] = _now()
    return _state()


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)) -> dict[str, Any]:
    filename = Path(file.filename or "scan-report.xlsx").name
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")
    if len(raw) > 15 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File is larger than 15 MB. Upload a smaller scan report.")
    df, sheet_name, file_type = _read_uploaded(filename, raw)
    if df.empty:
        raise HTTPException(status_code=400, detail="The first worksheet or CSV contains no data rows.")
    columns = list(df.columns)
    expected_col = _find_column(columns, "expected")
    actual_col = _find_column(columns, "actual")
    if expected_col is None or actual_col is None:
        missing = []
        if expected_col is None:
            missing.append("expected location/area")
        if actual_col is None:
            missing.append("actual scan/staging location")
        raise HTTPException(status_code=422, detail={
            "message": "RouteGuard could not identify required column(s): " + ", ".join(missing) + ". Rename the headers or use the standard column names.",
            "columns": [str(c) for c in columns],
            "accepted_headers": {"expected": ALIASES["expected"], "actual": ALIASES["actual"]},
        })
    mapped = {
        "expected": expected_col,
        "actual": actual_col,
        "cage": _find_column(columns, "cage"),
        "route": _find_column(columns, "route"),
        "wave": _find_column(columns, "wave"),
        "success": _find_column(columns, "success"),
    }
    _set_store(df.copy(), filename, file_type, raw, sheet_name, mapped, "uploaded")
    return {"ok": True, "message": f"Analyzed {len(df)} rows from {filename}.", "state": _state()}


@app.post("/api/demo")
def load_demo() -> dict[str, Any]:
    df = pd.DataFrame([
        {"Cage ID": "C101", "Route ID": "R-ALPHA-01", "Wave": "Wave 3", "Original Scan Location": "A", "Actual Scan Location": "A", "isSuccess": True},
        {"Cage ID": "C102", "Route ID": "R-ALPHA-01", "Wave": "Wave 3", "Original Scan Location": "A", "Actual Scan Location": "B", "isSuccess": False},
        {"Cage ID": "C103", "Route ID": "R-ALPHA-01", "Wave": "Wave 3", "Original Scan Location": "A", "Actual Scan Location": "B", "isSuccess": False},
        {"Cage ID": "C104", "Route ID": "R-BRAVO-02", "Wave": "Wave 3", "Original Scan Location": "B", "Actual Scan Location": "B", "isSuccess": True},
        {"Cage ID": "C105", "Route ID": "R-CHARLIE-03", "Wave": "Wave 4", "Original Scan Location": "C", "Actual Scan Location": "A", "isSuccess": False},
        {"Cage ID": "C106", "Route ID": "R-CHARLIE-03", "Wave": "Wave 4", "Original Scan Location": "C", "Actual Scan Location": "C", "isSuccess": True},
        {"Cage ID": "C107", "Route ID": "R-DELTA-04", "Wave": "Wave 4", "Original Scan Location": "C", "Actual Scan Location": "B", "isSuccess": False},
        {"Cage ID": "C108", "Route ID": "R-DELTA-04", "Wave": "Wave 4", "Original Scan Location": "C", "Actual Scan Location": "C", "isSuccess": True},
        {"Cage ID": "C109", "Route ID": "R-ALPHA-01", "Wave": "Wave 3", "Original Scan Location": "A", "Actual Scan Location": "B", "isSuccess": True},
        {"Cage ID": "C110", "Route ID": "R-BRAVO-02", "Wave": "Wave 3", "Original Scan Location": "B", "Actual Scan Location": "", "isSuccess": False},
    ])
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Scan Report")
    raw = out.getvalue()
    mapped = {"expected": "Original Scan Location", "actual": "Actual Scan Location", "cage": "Cage ID", "route": "Route ID", "wave": "Wave", "success": "isSuccess"}
    _set_store(df.copy(), "routeguard_sample_scan_report.xlsx", "xlsx", raw, "Scan Report", mapped, "demo")
    STORE["events"] = [{"time": _now(), "cage_id": "C109", "action": "loaded", "message": "Sample report loaded; C109 represents an already-resolved mismatch."}]
    return {"ok": True, "message": "Sample report loaded. This is simulated data, not live Station Command data.", "state": _state()}


@app.post("/api/records/{record_id}/action")
def act_on_record(record_id: str, request: ActionRequest) -> dict[str, Any]:
    if not STORE["loaded"]:
        raise HTTPException(status_code=409, detail="Upload a report or load sample data first.")
    rec = next((r for r in STORE["records"] if r["id"] == record_id), None)
    if rec is None:
        raise HTTPException(status_code=404, detail="The record could not be found. Refresh the report and try again.")
    if not rec["is_wrong"]:
        raise HTTPException(status_code=400, detail="Only wrong-location records can be acknowledged or resolved.")

    if request.action == "acknowledge":
        if rec["status"] == "resolved":
            raise HTTPException(status_code=409, detail="This mismatch is already marked resolved.")
        rec["status"] = "acknowledged"
        message = f"{rec['cage_id']} acknowledged for investigation. isSuccess remains FALSE."
    elif request.action == "resolve":
        if not request.verified:
            raise HTTPException(status_code=400, detail="Confirm that the correction has been verified before resolving.")
        rec["status"] = "resolved"
        rec["isSuccess"] = True
        success_col = STORE["columns"].get("success")
        if success_col is None:
            success_col = "isSuccess"
            STORE["df"][success_col] = False
            STORE["columns"]["success"] = success_col
        STORE["df"].at[int(rec["id"]), success_col] = True
        message = f"{rec['cage_id']} marked resolved; isSuccess updated to TRUE."
    else:  # reopen
        rec["status"] = "open"
        rec["isSuccess"] = False
        success_col = STORE["columns"].get("success", "isSuccess")
        if success_col not in STORE["df"].columns:
            STORE["df"][success_col] = False
            STORE["columns"]["success"] = success_col
        STORE["df"].at[int(rec["id"]), success_col] = False
        message = f"{rec['cage_id']} reopened; isSuccess updated to FALSE."
    event = {"time": _now(), "cage_id": rec["cage_id"], "action": request.action, "message": message, "note": request.note.strip()}
    STORE["events"].append(event)
    STORE["last_refresh"] = _now()
    return {"ok": True, "message": message, "state": _state()}


@app.get("/api/export")
def export_file() -> Response:
    if not STORE["loaded"]:
        raise HTTPException(status_code=409, detail="Load or upload a scan report before exporting.")
    raw = _excel_bytes()
    filename = STORE["filename"] or "routeguard-export.xlsx"
    stem = Path(filename).stem
    if STORE["file_type"] == "csv":
        return Response(raw, media_type="text/csv; charset=utf-8", headers={"Content-Disposition": f'attachment; filename="{stem}_routeguard_updated.csv"'})
    return Response(raw, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="{stem}_routeguard_updated.xlsx"'})


@app.get("/api/events")
def get_events() -> dict[str, Any]:
    return {"events": list(reversed(STORE["events"][-50:]))}
