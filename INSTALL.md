# Install SheetFreak

Windows desktop Excel is required. You can keep using ordinary `.xlsx` files.

## Download

Open [SheetFreak Releases](https://github.com/VarunV21/SheetFreak/releases) and download the **SheetFreak-vVERSION.zip** asset attached to a tested release. GitHub's automatically generated **Source code (zip)** contains source, not an installable add-in. If no tested release asset is listed, an installable canonical release is not available yet.

Right-click the ZIP and choose **Extract All**. Never register the add-in from inside the ZIP or a temporary preview folder.

## Standard installation

1. Keep the extracted folder somewhere permanent, for example `Documents\SheetFreak`.
2. In Excel, open **File → Options → Add-ins**.
3. Beside **Manage**, choose **Excel Add-ins**, then **Go**.
4. Choose **Browse** and select the extracted `SheetFreak.xlam`.
5. Ensure **Sheetfreak is checked**, click **OK**, then fully close and reopen Excel.
6. Open a blank workbook. The SheetFreak tab should appear; **Alt, S** is the requested tab shortcut (pending native verification). Excel may reassign the key tip; press Alt to see the displayed letters.

If the tab is missing, revisit the Add-ins list and check the filename, location and checkbox. Simply double-clicking the `.xlam` does not verify persistent installation. If Office blocks the unsigned add-in, follow your organization's policy or ask IT. Do not enable all macros or disable Protected View to install SheetFreak.

## Optional installer

The release ZIP contains `Install-SheetFreak.ps1` next to `SheetFreak.xlam` and `SHA256SUMS.txt`.

1. Save your work and close **all** Excel windows.
2. Right-click `Install-SheetFreak.ps1` and choose **Run with PowerShell**, if permitted by your device's policy.
3. Open Excel normally and confirm the tab appears.

The installer checks the included checksum, copies the file to `%LOCALAPPDATA%\SheetFreak\SheetFreak.xlam`, and registers and enables it in Excel. It uses Excel's own add-in interface; it does not require administrator privileges or VBA-project access. It does not download anything or change security settings. If script execution is blocked, use the standard installation above or ask IT; do not bypass your organization's policy. Installer success still requires a normal Excel restart check.

The installer is prepared for the 2.0.0 release and requires native Windows/Excel testing before publication. Do not present it as a tested installation path until the release checklist passes.

## Updates

Download the new tested release ZIP and fully close Excel before replacing the installed file. Keep its filename and installation location the same, then reopen Excel and check the Help version.

If you used the optional installer, run the new release's installer again. It keeps the previous binary alongside the installed file with a `.previous-...` suffix, and restores it if registration fails. If you originally installed manually in a different folder, update that same folder manually; do not register a second SheetFreak copy.

## Remove

In Excel, go to **File → Options → Add-ins → Manage: Excel Add-ins → Go**, uncheck Sheetfreak and click OK. Restart Excel. You can then remove the extracted / installed files if you no longer need them.
