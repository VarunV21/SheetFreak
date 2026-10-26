"""Package a maintainer-validated Excel build, never an old local preview."""
from pathlib import Path
import argparse
import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET

from check_source import ROOT, check

REPORTS = {
    "native-test-results.txt": "ALL TESTS PASSED",
    "zoom-test-results.txt": "ALL ZOOM CHECKS PASSED",
    "workbook-view-test-results.txt": "ALL WORKBOOK VIEW CHECKS PASSED",
}


def validate_reports(folder):
    result = {}
    for filename, marker in REPORTS.items():
        data = (Path(folder) / filename).read_bytes()
        text = data.decode("utf-8-sig").splitlines()
        if marker not in text or any("FAIL" in line.upper() for line in text):
            raise ValueError(f"Native report did not pass: {filename}")
        if not any(line.startswith("PASS:") for line in text):
            raise ValueError(f"Native report has no checks: {filename}")
        result[filename] = hashlib.sha256(data).hexdigest()
    return result


def validate_binary(binary, ribbon):
    with zipfile.ZipFile(binary) as archive:
        if archive.testzip() is not None:
            raise ValueError("Corrupt add-in ZIP")
        for name in archive.namelist():
            if name.endswith((".xml", ".rels")):
                ET.fromstring(archive.read(name))
        if not archive.read("xl/vbaProject.bin").startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
            raise ValueError("Missing VBA compound document")
        if archive.read("customUI/customUI.xml") != Path(ribbon).read_bytes():
            raise ValueError("Add-in ribbon differs from this checkout; rebuild the candidate")
        for test_module in (b"SheetFreakTests", b"ZoomTests", b"WorkbookViewTests"):
            if test_module in archive.read("xl/vbaProject.bin"):
                raise ValueError("Test module appears in production VBA")


def release(binary, reports, confirmed=False, output=None):
    if not confirmed:
        raise ValueError("Native Excel and installer validation must be confirmed before release")
    version = check()
    report_hashes = validate_reports(reports)
    validate_binary(binary, ROOT / "ribbon/customUI.xml")
    data = Path(binary).read_bytes()
    sha = hashlib.sha256(data).hexdigest()
    source_hashes = {
        str(path.relative_to(ROOT)).replace("\\", "/"): hashlib.sha256(path.read_bytes()).hexdigest()
        for path in [*(ROOT / "src").glob("*.bas"), ROOT / "ribbon/customUI.xml"]
    }
    provenance = {"version": version, "xlam_sha256": sha, "source_sha256": source_hashes,
                  "native_report_sha256": report_hashes,
                  "validation": "Maintainer attested native build, tests and installation checks"}
    destination = Path(output) if output else ROOT / "dist" / f"SheetFreak-v{version}.zip"
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("SheetFreak/SheetFreak.xlam", data)
        archive.writestr("SheetFreak/SHA256SUMS.txt", f"{sha}  SheetFreak.xlam\n")
        archive.writestr("SheetFreak/BUILD-INFO.json", json.dumps(provenance, indent=2) + "\n")
        for filename in ("INSTALL.md", "LICENSE", "VERSION"):
            archive.write(ROOT / filename, f"SheetFreak/{filename}")
        archive.write(ROOT / "install/Install-SheetFreak.ps1", "SheetFreak/Install-SheetFreak.ps1")
    return destination


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("binary", type=Path)
    parser.add_argument("--reports", type=Path, required=True)
    parser.add_argument("--confirm-native-tested", action="store_true")
    args = parser.parse_args()
    print(release(args.binary, args.reports, args.confirm_native_tested))
