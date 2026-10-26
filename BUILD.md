# Build and release

## Current status

Version 2.0.0 was compiled from this source and passed 45 engine, 12 zoom and 16 workbook-view native Excel checks. Ribbon Fill and Freeze, cancel behavior, colors, Alt S, and persistent loading after restart were also checked. The installer upgrade required the checkbox step documented in INSTALL.md. These checks do not establish compatibility with every Office configuration. Offline checks do not establish that Excel has compiled, loaded or tested the new VBA. The existing local 0.2.1 binary must not be relabeled or packaged as 2.0.0.

## Build in Excel

1. Use a new disposable workbook in Windows desktop Excel. Import the production `.bas` module from `src/` through the VBA editor. Do not enable programmatic VBA-project access.
2. Choose **Debug → Compile VBAProject**. Save as `SheetFreak-unpackaged.xlam` and close the build workbook.
3. From the repository root, with Python 3, run:

```text
python tools/package.py SheetFreak-unpackaged.xlam build/SheetFreak.xlam ribbon/customUI.xml
```

4. Load this candidate, not an older installed copy, in Excel. Check Help shows 2.0.0, and test the ribbon and ribbon key tips.

## Native validation

Run the source engine tests in a separate disposable `.xlsm` containing the production source and `tests/SheetFreakTests.bas`. `SheetFreakTests.RunAll` writes `native-test-results.txt` beside that workbook. It intentionally creates a test workbook; the production Fill and freeze operation must not create one.

Against the **candidate** `SheetFreak.xlam`, import the two view test modules into a disposable test host and run:

```vb
ZoomTests.RunZoomChecks Environ$("TEMP") & "\zoom-test-results.txt"
WorkbookViewTests.RunWorkbookViewChecks Environ$("TEMP") & "\workbook-view-test-results.txt"
```

Complete [the manual checks](tests/MANUAL_TESTS.md), including canceling Fill and freeze, purple seed formulas, the requested S tab key tip, margins at different window sizes, grouped/protected sheets, frozen panes and two windows for one workbook. Verify that the selected window changes and the other window does not. Check application events and screen updating restore when a sheet cannot be resized.

Test both manual installation and the optional installer, updates at the same path, rollback after a registration failure, blocked scripts, and loading after a full Excel restart. Use disposable workbooks; never close or discard unsaved work to run these checks.

## Package a tested release

```text
python tools/check_source.py
python -m unittest discover -s tests -p "test_*.py"
python tools/release.py build/SheetFreak.xlam --reports build/reports --confirm-native-tested
```

Place this build's three native reports in `build/reports/` with the filenames above. The release helper requires passing report markers and records source and binary hashes. This is a human validation gate, not automatic proof that VBA matches the source: the maintainer must verify the candidate was built from this checkout and complete the manual checks before passing the flag.

The helper creates `dist/SheetFreak-v2.0.0.zip`, containing the add-in, optional installer, checksum, installation guide, license and version. No user workbooks or machine-specific installation notes are included.

After user testing, merge the reviewed change, tag the exact tested commit and publish the ZIP as a GitHub Release asset. Keep draft changes unreleased until native validation and user acceptance pass. Users should download that asset, not GitHub's source-code ZIP.

The requested tab key tip is S. During native testing verify Alt, S opens SheetFreak and H/R/W dispatch the correct tools. If Excel reassigns S because of a conflict, retain the known SF fallback and update all public instructions and Help to match before release. F is reserved for File. Confirm Ctrl+Alt+V still opens native Paste Special after a clean restart.
