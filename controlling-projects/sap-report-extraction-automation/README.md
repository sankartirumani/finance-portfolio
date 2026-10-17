# SAP Report Extraction Automation (Excel VBA + SAP GUI Scripting)

One-click monthly extraction of large SAP controlling reports into Excel, built for a finance controlling team.

| | |
|---|---|
| **Problem** | Every month-end, analysts ran the same heavy SAP reports by hand: open the report, set period, load a long selection list, drill into a line, pick a layout, export. Slow, repetitive, and easy to get wrong (wrong layout, wrong period). |
| **Solution** | An Excel workbook with buttons. Each analyst sets year and period, then clicks once. The macro starts SAP, runs the report, picks the exact layout and saves the export to a shared folder. |
| **Tools** | Excel VBA, SAP GUI Scripting, Windows Script Host |
| **Result** | Hands-off extraction of two reports (cost centre and project/WBS view) run sequentially. A PARAM sheet logs when each last completed. |

## Who can use this

Any team that repeatedly runs the same SAP report and exports it to Excel. You change the settings sheet, not the code.

## Quick start

1. Read [docs/setup-guide.md](docs/setup-guide.md). It covers prerequisites and the one-time SAP scripting check.
2. Import [src/SapExtraction.bas](src/SapExtraction.bas) into a new `.xlsm` (VBA editor, File, Import).
3. Build the Settings sheet from [config/settings-template.csv](config/settings-template.csv).
4. Assign a button to `RunAll` (or call `RunExtraction "SheetName"` for a single report).

## Documentation

- [How it works](docs/how-it-works.md): architecture, run sequence, design decisions
- [Setup guide](docs/setup-guide.md): prerequisites and configuration
- [Lessons learned](docs/lessons-learned.md): what broke, how I fixed it, what I'd improve

## Key design ideas

- **Config-driven.** Transaction, period, paths and layout live in cells, not code. Reusing it for another report means editing cells.
- **Multi-user.** Each colleague has different local paths, so a per-user dropdown lookup chooses the right folders and shortcut.
- **No credentials in code.** Login uses the user's saved SAP shortcut and single sign-on. The macro never sees a password.
- **Fails safe.** Layout selection requires an exact name match and aborts rather than exporting the wrong layout.
- **Resilient to slow systems.** Waits and a popup-probing pattern handle different SAP screens and large result sets.

## Status and caveats

The code here is a **sanitized, generic version** of a tool used internally. Company-specific transaction codes, organisational values and file paths have been replaced with settings. The recorded screen coordinates (`lbl[x,y]`) and control IDs are specific to each report, so you will need to re-record them for your own report (see the setup guide). This generic version has not been run against a live SAP system.

Only use this where SAP GUI Scripting is permitted by your IT policy. See the setup guide.
