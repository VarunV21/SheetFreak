"""Offline packaging tests using synthetic fixtures; these do not run Excel."""
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
from check_source import ROOT, check
from package import package
from release import REPORTS, release, validate_binary, validate_reports


class ReleaseChecks(unittest.TestCase):
    def test_source_consistency(self):
        self.assertEqual(check(), "2.0.0")

    def test_native_confirmation_required(self):
        with self.assertRaisesRegex(ValueError, "validation must be confirmed"):
            release(Path("not-built.xlam"), Path("no-reports"))

    def test_missing_native_reports_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(FileNotFoundError):
                validate_reports(folder)

    def test_failed_report_rejected_even_with_pass_footer(self):
        with tempfile.TemporaryDirectory() as folder:
            for name, marker in REPORTS.items():
                Path(folder, name).write_text("PASS: example\n" + marker + "\nFAIL: error\n")
            with self.assertRaisesRegex(ValueError, "did not pass"):
                validate_reports(folder)

    def test_empty_report_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            for name, marker in REPORTS.items():
                Path(folder, name).write_text(marker + "\n")
            with self.assertRaisesRegex(ValueError, "no checks"):
                validate_reports(folder)

    def test_missing_vba_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder, "empty.xlam")
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
            with self.assertRaisesRegex(ValueError, "no compiled VBA"):
                package(source, Path(folder, "out.xlam"), ROOT / "ribbon/customUI.xml")

    def test_packaging_replaces_ribbon_and_preserves_property_namespaces(self):
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder, "raw.xlam"), Path(folder, "out.xlam")
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
                archive.writestr("xl/vbaProject.bin", bytes.fromhex("D0CF11E0A1B11AE1"))
                archive.writestr("docProps/core.xml", '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Private Name</dc:creator><cp:lastModifiedBy>Private Name</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created></cp:coreProperties>')
            package(source, output, ROOT / "ribbon/customUI.xml")
            validate_binary(output, ROOT / "ribbon/customUI.xml")
            with zipfile.ZipFile(output) as archive:
                props = archive.read("docProps/core.xml")
                self.assertNotIn(b"Private Name", props)
                self.assertIn(b"xmlns:dcterms=", props)
                self.assertIn(b"SheetFreak contributors", props)
            wrong_ribbon = Path(folder, "wrong.xml")
            wrong_ribbon.write_text("<wrong/>")
            with self.assertRaisesRegex(ValueError, "ribbon differs"):
                validate_binary(output, wrong_ribbon)


if __name__ == "__main__":
    unittest.main()
