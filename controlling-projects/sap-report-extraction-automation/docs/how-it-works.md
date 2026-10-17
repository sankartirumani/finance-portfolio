# How it works

## Architecture

```
Excel workbook (.xlsm)
├── Guide      – instructions for end users
├── Settings   – one sheet per report: shortcut, paths, period, layout, button
├── Input      – per-user lookup lists (local folders and shortcut per username)
├── PARAM      – timestamps of the last successful run per report
└── VBA        – one module: start SAP → run report → export → close
```

## Run sequence

1. **Read settings** from the report's sheet.
2. **Start SAP** by launching the user's saved `.sap` shortcut through Windows Script Host, exactly like a double-click. The shortcut holds the connection and SSO configuration.
3. **Attach** to the running SAP GUI through its scripting engine: `GetObject("SAPGUI").GetScriptingEngine`, then the first connection and session.
4. **Run the report.** Open the transaction, answer the controlling-area popup, fill year and period.
5. **Load the long selection list** from a text file instead of typing it. SAP's multiple-selection popup has an "import from text file" button, and the file was saved from SAP once, so lists of hundreds of values load in one step.
6. **Execute and wait.** Large reports need fixed waits (10 s to 90 s depending on size).
7. **Drill to the target line** using SAP's Find (Ctrl+F) and a double-click.
8. **Choose the layout by exact name** (see below).
9. **Export** to the configured folder and file name, handling the "overwrite?" popup.
10. **Stamp** the completion time in PARAM and optionally close SAP.

## Design decisions

### Settings in cells, code stays generic
The first version had paths and codes inside the macros. When a colleague's OneDrive path differed, the code had to change. Moving every variable to a sheet made the same macro work for everyone.

### Per-user path lookup
OneDrive and SharePoint sync folders differ per user (different roots, different language desktops such as `Desktop` and `Bureau`). An `Input` sheet maps username to paths, and data-validation dropdowns on the settings sheet let a new user pick theirs. A new user adds a row once.

### Exact layout selection
SAP's Find in the layout list matches substrings, so searching `LAYOUT_A` could land on `LAYOUT_A_OLD`. The macro loops "find next" (max 30 times) until the highlighted row equals the requested name, otherwise it aborts with a message. This turns a silent wrong-data failure into a visible stop.

### Probing for optional screens
Different reports show either an extra export-format popup or go straight to the file dialog. Instead of one script per report, the code tries to find the path field. If it isn't there, it confirms the popup and continues.

### Restart SAP between heavy reports
After a very large extraction the GUI becomes sluggish. Closing and relaunching SAP between reports kept the second run reliable.

## Known user prompt

During a large extraction Excel may show "Microsoft Excel is waiting for another application to complete an OLE action." This is normal. The user clicks OK and the macro continues. The Guide sheet documents it with a screenshot.
