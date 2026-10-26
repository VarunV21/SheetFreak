# Optional per-user installer. Run only from a tested SheetFreak release ZIP.
# No administrator rights, VBA-project access or security-policy changes needed.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$excelApp = $null
$registeredAddIn = $null
$previouslyInstalled = $false
$backupPath = $null
$copied = $false
$installPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'SheetFreak\SheetFreak.xlam'
try {
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        throw 'Save your work and close ALL Excel windows, then run the installer again.'
    }
    $sourcePath = Join-Path $PSScriptRoot 'SheetFreak.xlam'
    $manifestPath = Join-Path $PSScriptRoot 'SHA256SUMS.txt'
    if (!(Test-Path -LiteralPath $sourcePath) -or !(Test-Path -LiteralPath $manifestPath)) {
        throw 'Download and extract a release ZIP. The source-code ZIP does not contain a built add-in.'
    }
    $hashLine = Get-Content -LiteralPath $manifestPath | Where-Object { $_ -match '^[a-fA-F0-9]{64}\s+SheetFreak\.xlam$' }
    if (@($hashLine).Count -ne 1) { throw 'The release checksum file is invalid.' }
    $expectedHash = ($hashLine -split '\s+')[0]
    if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ne $expectedHash) {
        throw 'The add-in does not match the release checksum. Download the release again.'
    }
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    # Excel requires a workbook window before AddIns.Add can register an XLAM.
    $setupWorkbook = $excelApp.Workbooks.Add()
    $setupWorkbook.Saved = $true
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($setupWorkbook)
    foreach ($existingAddIn in $excelApp.AddIns) {
        try {
            if ($existingAddIn.Name -ieq 'SheetFreak.xlam') {
                # Keep the registered filename and location for reliable upgrades.
                $installPath = $existingAddIn.FullName
                $registeredAddIn = $existingAddIn
                $previouslyInstalled = $registeredAddIn.Installed
                $registeredAddIn.Installed = $false
            }
        } finally {
            if ($existingAddIn -ne $registeredAddIn) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($existingAddIn)
            }
        }
    }
    New-Item -ItemType Directory -Path (Split-Path $installPath) -Force | Out-Null
    if (Test-Path -LiteralPath $installPath) {
        $backupPath = $installPath + '.previous-' + [Guid]::NewGuid().ToString('N')
        Copy-Item -LiteralPath $installPath -Destination $backupPath
    }
    Copy-Item -LiteralPath $sourcePath -Destination $installPath -Force
    $copied = $true
    if ($null -eq $registeredAddIn) {
        $registeredAddIn = $excelApp.AddIns.Add($installPath, $false)
    }
    $registeredAddIn.Installed = $true
    if (!$registeredAddIn.Installed) { throw 'Excel did not keep SheetFreak enabled.' }
    $excelApp.Quit()
    Write-Host "SheetFreak has been copied to $installPath"
    Write-Host 'IMPORTANT: In normal Excel, open File > Options > Add-ins > Excel Add-ins > Go.'
    Write-Host 'Check Sheetfreak and click OK. If absent, Browse to the path above.'
    Write-Host 'Fully restart Excel, confirm the tab appears, then try Alt, S.'
    Write-Host 'If your organization blocks unsigned add-ins, use its approved installation process.'
} catch {
    $installFailure = $_
    if ($null -ne $registeredAddIn) {
        try { $registeredAddIn.Installed = $false } catch { }
    }
    if ($copied -and $null -ne $backupPath) {
        Copy-Item -LiteralPath $backupPath -Destination $installPath -Force
    }
    if ($null -ne $registeredAddIn) {
        try { $registeredAddIn.Installed = $previouslyInstalled } catch { }
    }
    if ($null -ne $excelApp) { try { $excelApp.Quit() } catch { } }
    Write-Error $installFailure -ErrorAction Continue
    Write-Host 'See INSTALL.md for the manual installation steps. No security settings were changed.'
    exit 1
} finally {
    if ($null -ne $registeredAddIn) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($registeredAddIn) }
    if ($null -ne $excelApp) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp) }
}
