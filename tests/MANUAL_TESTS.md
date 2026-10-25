# User acceptance checks

Use a disposable `.xlsx` workbook. Save it before testing.

- Install the add-in, restart Excel, and confirm the SheetFreak tab and direct shortcuts are available in a normal `.xlsx`.
- Test currency, accounting, comma and percent formats together, plus dates; zero decimals should preserve the numeric style and underlying values.
- Format several assumptions and a multirow header. Check the blue/yellow colors and every header border.
- Copy a rectangular range containing relative, absolute and mixed references, constants, blanks and text beginning with `=`. Paste to a new location and compare formula text. Try another worksheet and workbook too.
- Confirm overwriting content with Paste exact offers a cancel path.
- Run Auto resize on a wide worksheet and a blank sheet; confirm A1 and readable zoom.
- For fill-and-freeze, use a few hundred rows with two formula columns and different formats. Choose a reference column with gaps. Check the preview, choose No, and verify no changes or extra workbook creation. Repeat and accept. Check the seed formulas, last row values, black font, AutoFit and confirm no new workbook or sheet is created.
- Change source data: seed formulas update; frozen lower rows stay unchanged. Re-run the tool to refresh those snapshots.
- Check a filtered sheet, grouped tabs, protected sheet, Table, merged cell and spilled formula target; inappropriate operations should be rejected clearly.
- Disable shortcuts and verify ribbon buttons still work. Restart and check shortcuts register again.
- Uncheck the add-in and restart; check the tab and bindings are removed.

Report: Excel version, command used, selected range, expected outcome and actual outcome. Use synthetic data in public issue reports.

- Press Alt, S, F to open SheetFreak. Verify H applies Header and R runs Auto resize. Confirm Format, Formulas, View and Help are separate groups.
