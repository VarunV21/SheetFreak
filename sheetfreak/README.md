# SheetFreak

A source-available personal Excel toolbox for Windows desktop Microsoft 365. A separate add-in brings keyboard shortcuts and a custom ribbon into ordinary workbooks.

This repository contains reviewable VBA source, ribbon XML, tests and a build helper. It does not distribute a compiled `.xlam` file or any working/model workbooks. Version represented: 0.2.1 preview.

## What it does

| Tool | Shortcut | Behavior |
| --- | --- | --- |
| Zero decimals | Ctrl+Alt+0 | Removes decimal display placeholders while preserving supported currencies, percentages, conditions and literals. Stored values retain their precision. |
| Copy / paste exact | Ctrl+Alt+C / Ctrl+Alt+V | Copies formula text without Excel's relative-reference translation, using a private session buffer. |
| Assumption format | Ctrl+Alt+A | Applies blue text and pale yellow fill. |
| Header format | Ctrl+Alt+T | Applies gray fill, white bold text and cell borders. |
| Auto resize | Ctrl+Alt+H | Fits populated columns within 70–200% zoom and returns to A1. |
| Fill and freeze | Ctrl+Alt+F | Fills seed formulas, calculates before freezing lower rows into values, and retains live seed formulas. |

Ribbon entry: Alt, S, F. Excel may reassign key tips when another add-in conflicts. Shortcuts can be disabled from the ribbon.

## Engineering decisions

- The zero-decimal formatter is a pure function. It preserves unsupported date, time, fraction and text formats instead of guessing.
- Exact copying records both `Formula2` and each cell's `HasFormula`, so literal strings beginning with an equals sign remain text.
- Fill and freeze populates and calculates the whole target before converting any block into values. This preserves dependencies across rows and columns.
- Validation rejects incompatible selections, filters, protection, tables, merged cells, array formulas and spill ranges for operations that cannot support them.
- The fill engine restores calculation mode, events, screen updating, cancellation behavior and status state on success or error. A calculation error stops freezing.
- Limits bound operations to 250,000 formatting cells, 100,000 copied cells and 5 million fill cells. These are implementation limits, not performance benchmarks.

## Build in Excel

1. Start a new blank workbook in desktop Microsoft 365 Excel for Windows.
2. Import the two `.bas` files from `src/` using the VBA editor.
3. Save as an Excel add-in (`.xlam`) and close it.
4. With Python 3, package the ribbon into a different output file:

```text
python tools/package.py SheetFreak-unpackaged.xlam SheetFreak.xlam ribbon/customUI.xml
```

The helper preserves the Excel-produced compiled VBA project and rewrites creator and last-editor properties. It does not guarantee removal of every workbook metadata field. Inspect and sanitize generated artifacts before sharing them. No compiled binary from the original archive is included here.

Load a locally built add-in through Excel's Add-ins dialog, using your normal organization security policy. It does not require programmatic VBA project access. Use a disposable workbook for initial checks.

## Validation evidence

The supplied archive records 43 native engine checks on 2026-10-19 for version 0.2.0, including a two-column fill covering 100,000 lower rows. Version 0.2.1 has eight additional recorded zoom checks. The engine suite was not rerun for that zoom-only change. These are historical records supplied with the project, not independently rerun verification of this source snapshot.

To run the native suite in a disposable `.xlsm`, import `src/SheetFreak.bas` and `tests/SheetFreakTests.bas`, then run `SheetFreakTests.RunAll` from the editor. The separate `ZoomTests` module expects the add-in to be loaded. See `tests/MANUAL_TESTS.md` for UI checks. Test modules are excluded from release builds.

## Limitations

VBA edits clear normal Excel Undo. Fill and freeze overwrites target content, includes hidden rows, and makes no recovery workbook; save a backup before running it. Cancellation or a later error can leave partial changes. The confirmation dialog defaults to No. Frozen results are snapshots and must be regenerated when inputs change.

Exact copying preserves formula text, not its original meaning across sheets or workbooks. Unqualified references resolve at the destination. Floating shapes do not determine Auto resize width. Keyboard layouts and other add-ins can conflict. The preview is unsigned and needs broader compatibility testing before wider distribution.

No telemetry, networking, external services, automatic workbook saving or worksheet UDFs appear in the reviewed production source.

## License

MIT. See `LICENSE`. The original contributor attribution is retained.
