# NPTEL Operations Studio — Portable Coordinator Verification Suite

> **Important:** The software is available in the `dist` folder. Do not extract or run it from the `C:` drive. Copy or extract it to the `D:` drive before launching.

## How to Extract and Run

1. Open the project folder and go to the **`dist`** folder.
2. Locate the packaged `NPTEL-Studio` ZIP file (or the supplied application archive) inside `dist`.
3. Copy the ZIP file to a convenient location on the **`D:` drive**.
4. Right-click the ZIP file and select **Extract All…**. Extract the complete folder; do not run the application from inside the ZIP archive.
5. Open the extracted `NPTEL-Studio` folder.
6. Double-click **`NPTEL-Manager.exe`** to launch the application.

**Important:** Keep the extracted folder structure intact. Do not move the EXE out of its folder, because the application uses bundled modules and dependencies.

## Requirements

- Windows 10 or Windows 11 (64-bit)
- No installation required
- No administrative privileges required
- Microsoft Office is not required
- Works fully offline; all modules, UI views, and ImportExcel dependencies are pre-bundled

## Included Features

### Stage 1 — Registration Verification
- Registration verification and fee receipt ingestion
- Verified Students workspace with search
- Pagination with 25 students per page
- Receipt preview
- Unapprove and undo-staging options
- CSV export

### Stage 2 — Exam Scorecard Verification
- Post-exam scorecard ingestion
- Certificate mark parsing
- Marks breakdown:
  - Assignments: **25 marks**
  - Exam: **75 marks**
  - Total: **100 marks**
- Verified workspace
- Dynamic header configurator with **Add Header Info** support

### College Result Sheet Generator
Generates a cumulative result workbook containing four subsheets:

1. Main Result
2. CW
3. End Sem
4. NPTEL

### Offline OCR and PDF Processing
- Windows OCR support
- PDF rasterization
- No cloud API dependencies

### Email Center
- Bulk discrepancy notifications
- BCC support to help protect student email privacy

## Quick Reference

| Item | Details |
|---|---|
| Application package location | `dist` folder |
| Extracted folder | `NPTEL-Studio` |
| Application executable | `NPTEL-Manager.exe` |
| Supported OS | Windows 10/11 (64-bit) |
| Internet connection | Not required |
| Microsoft Office | Not required |
| Installation | Not required |
