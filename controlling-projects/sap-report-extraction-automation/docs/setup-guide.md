# Setup guide

## Prerequisites

- Windows with desktop Excel (macros enabled) and SAP GUI for Windows
- **SAP GUI Scripting enabled.** This is a standard SAP feature. It must be enabled on the server (by your SAP Basis team) and in your SAP GUI options (Options, Accessibility & Scripting, Scripting). If it is off, ask your IT or Basis team. Don't try to bypass it. Many organisations allow it for read-only reporting tasks.
- A saved SAP Logon shortcut (`.sap` file) for your system. Create one via SAP Logon, right-click the entry, Create Shortcut.

## Steps

### 1. Import the code
In Excel press Alt+F11, then File, Import File, and choose `src/SapExtraction.bas`. Save the workbook as `.xlsm`.

### 2. Build the settings sheet
Create a sheet (for example `ExtractA`) and fill it from `config/settings-template.csv`. Column B holds values and the row numbers must match the constants at the top of the module. Add a `PARAM` sheet with a header in row 3 and report names in A4, A5, and so on.

### 3. Save your selection list from SAP (optional)
If your report takes a long list of values, open the report's multiple-selection popup, fill the list, and use the local-file save button. Point the settings to that file.

### 4. Record your report's screen steps
Control IDs and label positions differ per report. To capture yours:

1. In SAP GUI, click the customize-local-layout icon, then *Script Recording and Playback*, then *Record*.
2. Perform the report run manually once.
3. Stop the recording and open the generated VBScript.
4. Copy each `session.findById(...)` line into the `RunReport` / `DrillTo` procedures, replacing literal values with settings cells.

### 5. Add the button
Developer tab, Insert, Button, then assign `RunAll` or a small wrapper that calls `RunExtraction "ExtractA"`.

### 6. Test with a small selection first
Start with a short list and a short period so each run is quick while you debug. Increase the `Application.Wait` times if SAP loads slower than the macro expects.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| "Run-time error 429: ActiveX can't create object" | SAP GUI not running or scripting disabled |
| `The control could not be found by id` | The screen differs from the recording, or SAP hadn't finished loading. Increase the wait. |
| Export aborts with "layout not found" | Layout name differs in your system, or it is user-specific |
| Excel says it is waiting for an OLE action | Normal during large exports. Click OK. |
