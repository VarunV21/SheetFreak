# User acceptance checks

Use a disposable `.xlsx` workbook. Save it before testing.

- Install the add-in, restart Excel, and confirm the SheetFreak tab and ribbon key tips are available in a normal `.xlsx`.
- Test currency, accounting, comma and percent formats together, plus dates; zero decimals should preserve the numeric style and underlying values.
- Format several assumptions and a multirow header. Check the blue/yellow colors and every header border.
- Copy a rectangular range containing relative, absolute and mixed references, constants, blanks and text beginning with `=`. Paste to a new location and compare formula text. Try another worksheet and workbook too.
- Confirm overwriting content with Paste exact offers a cancel path.
- Run Auto resize on compact, wide and blank sheets. Confirm A1, 70–120% zoom (blank 100%), gridlines off and a right margin when the 70% minimum permits. Repeat with gridlines already off.
- Run Resize all sheets (ribbon WA). Confirm each visible worksheet gets the same treatment, hidden/very hidden/chart sheets stay untouched, and the original tab is restored. Check another open workbook and another window of the same workbook remain unchanged. Include a protected sheet and confirm any failure is reported and application settings are restored.
- For fill-and-freeze, use a few hundred rows with two formula columns and different formats. Choose a reference column with gaps. Check the preview, choose No, and verify no changes or extra workbook creation. Repeat and accept. Check purple seed formulas (#7030A0), last row values, black font below the seed, AutoFit and confirm no new workbook or sheet is created.
- Change source data: seed formulas update; frozen lower rows stay unchanged. Re-run the tool to refresh those snapshots.
- Check a filtered sheet, grouped tabs, protected sheet, Table, merged cell and spilled formula target; inappropriate operations should be rejected clearly.
- Confirm no Ctrl+Alt bindings are registered and there are no enable/disable shortcut buttons. After a clean restart, Ctrl+Alt+V must still open native Paste Special.
- Uncheck the add-in and restart; check the tab is removed.

Report: Excel version, command used, selected range, expected outcome and actual outcome. Use synthetic data in public issue reports.

- Press Alt, S to open SheetFreak. Verify H applies Header, WC runs Auto resize and WA resizes all sheets. Confirm Format, Formulas, View and Help are separate groups.

- Confirm S actually opens SheetFreak in the installed Excel environment. If it conflicts, use SF and revise Help/README/INSTALL before release. Keep File on F.

## Share Excel (2.1)

- In a saved .xlsx, make unsaved edits, press Alt, S, S and verify Explorer selects a new closed file. Drag it into a test chat without submitting a message. Confirm that the original remains open and unsaved, and the snapshot contains those edits. Use synthetic data only.
- Repeat from a new never-saved workbook and with a worksheet chart selected. Share Excel should use ActiveWorkbook, independent of cell selection.
- Run twice quickly. Confirm the copies are different files and neither overwrites the other.
- Check formulas, formatting, hidden sheets, links and protection are preserved; the snapshot is the entire workbook. Test read-only source files and .xlsm/.xlsb/.xls formats. Confirm an unsaved workbook with VBA gets a clear save-as-.xlsm instruction.
- Test a password-encrypted workbook: the copy must remain protected; do not imply an AI service can read encrypted files. Check blocked/unsupported formats fail clearly.
- Test a failed save and an Explorer failure. Events must be restored; if the copy exists but Explorer fails, its path should be displayed.
- Confirm no Outlook draft, email, automatic upload, data refresh or source save occurs. Verify Alt S S and the drag/upload path in the intended ChatGPT/Claude app during acceptance testing.

## 2.2 presets and shortcuts

- After Alt S, check N/B/G/R/Y/4/K/P on multiple selected cells. Confirm gray italics for Notes, font-only colors, fill-only yellow, currency rather than accounting, comma integers and percent with one decimal. Values/formulas and unrelated formatting must remain unchanged.
- Test positive, negative, zero and text values, nonadjacent cells, protected/read-only sheets, and conditional formatting (which can override visible colors).
- Verify WC and WA are available alongside red R and copy C. Both resize commands cap at 120%, retain the existing 70% floor, right margin and gridlines-off behavior.
- User acceptance required before merging to main.
