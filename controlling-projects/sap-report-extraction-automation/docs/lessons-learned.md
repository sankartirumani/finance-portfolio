# Lessons learned

## What worked

- **Recording, then generalising.** The SAP script recorder gets a working prototype in an hour. The real work is then replacing every literal with a setting.
- **One orchestrating macro plus small helpers.** Early versions had the same start/close code copied into three modules. Delegating to one implementation removed drift between copies.
- **Fail loudly.** The exact-layout check turned an invisible data-quality risk into an obvious stop.
- **Documenting for non-technical users.** A one-page Guide sheet with a screenshot of the expected popup cut support questions.

## What was fragile

| Issue | Why | What I'd do next time |
|---|---|---|
| Fixed `Application.Wait` pauses | The macro guesses how long SAP needs. Too short breaks, too long wastes time. | Poll for the target control (`session.Busy`, or wait until `FindById` succeeds) with a timeout |
| Recorded label coordinates (`lbl[x,y]`) | Break if the report layout shifts | Search for text, or use the layout's own fields and grid objects |
| `On Error Resume Next` around optional popups | Can hide real errors | Wrap in a helper like `TryFind(id)` that returns Nothing and logs |
| `taskkill /IM saplogon.exe` | Closes every SAP window on the machine, including unrelated work | Close only the scripted session or connection |
| Duplicate start/close routines in several modules | Easy for copies to drift | Keep a single shared module (done in later versions) |
| Active-sheet dependence in early versions (`ActiveSheet.Range("B2")`) | Behaviour depends on which tab is selected | Refer to sheets by name (fixed in the final version) |
| Hard-coded user paths | Broke for each new colleague | Per-user lookup list with dropdowns |

## Working within IT policy

- Use SAP's own supported scripting interface. Nothing here bypasses authentication or authorisation.
- Users log in with their normal single sign-on through their own shortcut. The macro runs with that user's SAP rights only, and it sees no passwords.
- Scripting must be enabled by whoever administers SAP in your organisation. Raise it with them early.
- Keep org-specific values out of anything shared. That was also the rule when preparing this public version.

## Ideas for the next version

- Replace fixed waits with state-based waiting
- Log each run (user, period, duration, outcome) to a sheet
- Post-process the export automatically (Power Query load, reconciliation checks)
- Package settings in a single config table instead of per-report sheets
