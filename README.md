# SheetFreak

An Excel VBA add-in with a custom ribbon for formatting shortcuts and special formula manipulation. Built for Windows desktop Microsoft 365.

[SheetFreak on my website](https://varun-verma-portfolio.vercel.app/projects/sheetfreak/) · [My portfolio](https://varun-verma-portfolio.vercel.app/)

This repository contains the VBA source, ribbon XML, native Excel tests, and a build helper for version 0.2.1 preview. A compiled `.xlam` is not included.

## Tools and ribbon shortcuts

Press each key in sequence, starting with Alt; do not hold the keys together. These sequences use the ribbon key tips defined in this repository. Excel may reassign them if another add-in conflicts.

| Tool | Ribbon sequence | What it does |
| --- | --- | --- |
| Zero decimals | Alt → S → F → 0 | Removes supported decimal display placeholders while retaining number styles and stored precision. |
| Assumption format | Alt → S → F → A | Applies blue text and pale yellow fill. |
| Header format | Alt → S → F → H | Applies gray fill, bold white text, and borders. |
| Copy exact | Alt → S → F → C | Copies formula text into a private session buffer. |
| Paste exact | Alt → S → F → V | Pastes copied formulas without translating relative references. |
| Fill and freeze | Alt → S → F → F | Fills and calculates the target, then freezes lower rows into values while retaining seed formulas. |
| Auto resize | Alt → S → F → R | Returns to A1 and fits populated columns within 70–200% zoom. |

## Source layout

- [src/](src/) — production VBA modules.
- [ribbon/customUI.xml](ribbon/customUI.xml) — ribbon controls and key tips.
- [tests/](tests/) — native regression modules and manual checks.
- [tools/package.py](tools/package.py) — adds the ribbon to an Excel-produced add-in.

The number-format parser leaves unsupported dates, times, and fractions unchanged. Exact copying records whether each cell contains a formula, so literal text beginning with `=` stays text. Fill and freeze calculates the whole target before freezing values, checks for incompatible selections, and restores Excel's application settings after success or error.

## Build and install

1. Open a blank workbook in Windows desktop Excel.
2. In the VBA editor, import both `.bas` modules from `src/`.
3. Save as an Excel add-in (`.xlam`) and close it.
4. With Python 3, run the following from the repository root, using different input and output filenames:

```text
python tools/package.py SheetFreak-unpackaged.xlam SheetFreak.xlam ribbon/customUI.xml
```

5. In Excel, open **File → Options → Add-ins → Manage: Excel Add-ins → Go → Browse**, select the packaged `SheetFreak.xlam`, and check it. Restart Excel if needed.

Keep the add-in in a permanent folder. Follow your organization's macro-security policy. The packaging helper generalizes creator and last-editor properties but does not remove every possible metadata field; inspect a fresh build before distributing it.

## Testing

Use a disposable workbook. Import `src/SheetFreak.bas` and `tests/SheetFreakTests.bas` into a macro-enabled `.xlsm`, then run `SheetFreakTests.RunAll` from the VBA editor. For zoom checks, load the add-in, import `tests/ZoomTests.bas`, and run the following in the Immediate window (it writes a report to your temporary folder):

```vb
ZoomTests.RunZoomChecks Environ$("TEMP") & "\SheetFreak-zoom-results.txt"
```

See [the manual checklist](tests/MANUAL_TESTS.md) for ribbon, installation, formatting, and workbook checks. Test modules are not part of the release add-in.

## Limitations

- VBA edits clear normal Excel Undo. Fill and freeze overwrites target content, includes hidden rows, and creates no recovery workbook. Save a backup first; cancellation or a later error can leave partial changes.
- Frozen values are snapshots. Rerun the operation when inputs change.
- Exact copying preserves formula text. Unqualified references resolve at the destination sheet or workbook.
- Keyboard layouts and other add-ins can affect ribbon key tips. This preview is unsigned, and compatibility depends on the Excel environment.

## License

MIT. See [LICENSE](LICENSE) for the retained copyright and attribution.
