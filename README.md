# SheetFreak

An Excel toolbox for people who are freaks in the sheets. Install once; use its ribbon and shortcuts in normal `.xlsx` workbooks. Windows desktop Excel / Microsoft 365.

**[Download releases](https://github.com/VarunV21/SheetFreak/releases)** · **[Installation guide](INSTALL.md)** · [My website](https://varun-verma-portfolio.vercel.app/projects/sheetfreak/)

This repository is the canonical source. Version **2.0.0** was compiled and tested in Windows desktop Excel, including 73 native checks, ribbon Fill and Freeze, and loading after a full restart. Download the built release asset for installation.

**In development: 2.1.0 adds Share Excel (Alt, S, S).** A preview has been compiled and installed locally; full native regression and user testing are pending. The feature is not in the 2.0.0 download.

## Install

1. Under **Releases → Assets**, download `SheetFreak-v2.0.0.zip`. Do not choose GitHub's **Source code (zip)**.
2. Extract the ZIP into a permanent folder.
3. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go → Browse**. Select `SheetFreak.xlam`, ensure **Sheetfreak is checked**, then click **OK**.
4. Fully close and reopen Excel. Confirm the SheetFreak tab appears in a blank workbook.

No Python, developer tools or macro-enabled working workbook required. The release also includes an optional installer that copies the add-in to a stable folder and registers it. See [INSTALL.md](INSTALL.md), including updates and company-managed devices.

## Tools

Press ribbon keys one after another, starting with Alt. SheetFreak uses ribbon key tips only; it does not register Ctrl+Alt shortcuts or override native keyboard bindings. The tab key tip S was verified in Excel. Excel may reassign it when another tab conflicts.

| Tool | After Alt → S | What it does |
| --- | --- | --- |
| Zero decimals | 0 | Hide decimal places while preserving supported number styles and stored precision. |
| Assumption | A | Blue text and pale yellow fill for editable model inputs. |
| Header | H | Dark gray fill, bold white text, white borders around every cell. |
| Copy exact | C | Copy formulas into a private session buffer. |
| Paste exact | V | Paste formulas without shifting references. Does not copy formatting. |
| Fill and freeze | F | Choose a last-row column, confirm, fill/calculate formulas, then freeze lower rows as black-text values. Seed formulas remain purple; affected columns AutoFit. |
| Auto resize | R | Current worksheet: A1, gridlines off, fit populated columns with a margin, 70–150% zoom. |
| Resize all sheets | W | Same view adjustment on all visible worksheets in this workbook; return to the original tab. |

Auto resize reserves roughly 10% horizontal space on the right when the 70% minimum permits. Very wide sheets can still extend beyond the window at that minimum. Empty sheets use 100%. Hidden / very hidden sheets and chart sheets are skipped; a chart tab can still be the original tab returned to. View changes apply to the current workbook window. Column widths, freeze panes and content are unchanged. Both view commands reject grouped tabs. Save your workbook normally to retain view changes after closing it.

## Share Excel (2.1 preview)

Press **Alt, S, S** to write a closed snapshot of the whole workbook, including unsaved edits, then reveal it in File Explorer. The filename adds `_vS` before the extension, for example `Budget_vS.xlsx`. Drag that file into your chat, or use the chat's upload button. Outlook is not required and nothing is uploaded automatically. The original stays open, keeps its name and unsaved status, and is not saved by this command.

Each invocation creates a separate copy under `%TEMP%\SheetFreak\Share`. Copies are retained until you or Windows remove them; delete them when finished or move a copy somewhere permanent if needed. The snapshot includes hidden sheets, formulas, links, macros and existing protection; it is not a sanitized export. It preserves `.xlsx`, `.xlsm`, `.xlsb` or `.xls` format. Other formats must first be saved as a supported workbook; unsaved workbooks containing macros must first be saved as `.xlsm`. Your AI service must support and be able to open the resulting file. No external data refresh or forced recalculation is performed.

## Limitations

- VBA changes clear normal Excel Undo. Fill and freeze overwrites the target, including hidden rows, and creates no recovery workbook. Save your work first. Frozen values are snapshots; rerun when inputs change.
- Number formatting skips unsupported dates, times and fractions. Exact-copy preserves formula text; unqualified references resolve at the destination sheet or workbook.
- SheetFreak is unsigned. Company policies may block it; use your organization's approved installation process. The installer does not weaken security settings or enable VBA-project access.
- Other add-ins can affect ribbon placement and key tips. SheetFreak does not rearrange thinkcell or other add-ins.

## Development and release

- [src/](src/): production VBA; [ribbon/](ribbon/): ribbon XML.
- [tests/](tests/): native Excel tests and the [manual checklist](tests/MANUAL_TESTS.md).
- [BUILD.md](BUILD.md): build, test and release instructions. Editing `.bas` files alone does not update the installed binary.

MIT. See [LICENSE](LICENSE).
