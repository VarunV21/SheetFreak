# Changelog

## 0.2.1 - 2026-10-19

- Raise Auto resize's zoom ceiling from 100% to 200% so compact sheets can zoom in.
- Keep fit-to-width behavior, the 70% minimum, A1 navigation and 100% for empty sheets.
- Add focused native Excel checks for compact, wide and empty sheets, plus unchanged column widths and cell contents.

## 0.2.0 - 2026-10-19

- Use Alt, S, F to open SheetFreak; Excel reserves the V prefix and rejects both V and VV.
- Move SheetFreak directly after Home.
- Use large buttons for the main tools and separate Format, Formulas, View and Help groups.
- Rename A1 and fit width to Auto resize; its ribbon key tip is R.
- Use H for the Header command inside SheetFreak. Existing direct Ctrl+Alt shortcuts are unchanged.
- Remove automatic recovery workbook creation from fill-and-freeze; retain the explicit overwrite confirmation and calculation-error guard.
- Update native regression coverage to assert that no workbook is created on success or error.

## 0.1.0 - 2026-10-19

- First preview of the SheetFreak Excel add-in.
- Zero-decimal display formatting that retains supported numeric format styles.
- Exact formula copy/paste with a private session buffer.
- Assumption and header formatting.
- A1 navigation with bounded fit-to-width zoom.
- General fill-and-freeze with a user-chosen last-row column, explicit confirmation, recovery sheet, calculation-error guard, black text and AutoFit.
- Ribbon buttons, direct keyboard shortcuts, source code, native Excel tests and packaging tool.
