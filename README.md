# RouteGuard — Wrong-Route Cage Detection Demo

RouteGuard is a local demonstration portal for a Station Command improvement proposal. It reads an Excel/CSV scan report, compares each cage's expected area with its actual scan location, groups mismatches by actual scan location, and lets a user acknowledge or resolve exceptions. Resolving a verified issue changes the row's `isSuccess` value to `TRUE`; acknowledging does not.

**This is not connected to Amazon Station Command.** The five-second refresh refreshes the portal's current in-memory report and actions only. New physical scans require a fresh uploaded report or an approved internal data integration. Use only fabricated or sanitized data unless your organization has approved the use of real operational data.

## Requirements

- Python 3.10+
- Internet access only if you need to install dependencies

## Run on Windows

1. Extract `RouteGuard.zip`.
2. Open PowerShell in the extracted `routeguard` folder.
3. Run:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   uvicorn app.main:app --reload
   ```

4. Open <http://127.0.0.1:8000> in your browser.

You can also run `run_windows.bat` after Python is installed; it creates a virtual environment, installs packages, and starts the app.

## Run on macOS/Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open <http://127.0.0.1:8000>.

## Expected file columns

Required columns are an expected location/area and an actual scanned/staging location. Common accepted names include:

- Original: `Original Scan Location`, `Original Location`, `Original Scan Loc` (also accepts common expected/assigned location headers)
- Actual: `Actual Scan Location`, `Actual Staging Location`, `Scan Location`, `Scanned Location`, `Actual Area`
- Recommended: `Cage ID`, `Route ID`, `Wave`
- Status: `isSuccess` (optional)

The first worksheet in an Excel workbook is analyzed. RouteGuard compares original scan location with actual scan location exactly after normalizing case and whitespace. Original `A` vs actual `A` is correct; original `A` vs actual `B` is wrong. Original `Alpha` vs actual `Alpha A13` is also a mismatch because the location strings differ.

## `isSuccess` and exception states

- If `isSuccess` exists (including different capitalization), RouteGuard updates that existing column rather than adding a duplicate.
- If `isSuccess` is absent, RouteGuard adds it. Correct scans initialize to `TRUE`; mismatches and unverified rows initialize to `FALSE`.
- **Acknowledge** means investigation has started and leaves `isSuccess` `FALSE`.
- **Resolve** requires an explicit confirmation that the correction was verified and sets only that row's `isSuccess` to `TRUE`.
- **Reopen** sets it back to `FALSE`.
- A row that already has `isSuccess = TRUE` is displayed as resolved if it is a mismatch in the report.
- Export creates an updated copy; the original upload on disk is not overwritten.

## Notes and limitations

- This is a single-user demo. State is held in server memory, is not shared across multiple users, and resets when the server restarts.
- Excel exports aim to preserve the uploaded workbook structure and other columns; the tool updates `isSuccess` in the first worksheet. Complex workbooks, macros, unusual headers, formulas, or merged header rows should be tested before relying on the export.
- No data is intentionally written to a database. The server holds the current upload in memory until restarted or another report is uploaded.
- For a real Station Command implementation, use only approved internal APIs/event feeds and obtain the relevant security/privacy approval before integrating.
