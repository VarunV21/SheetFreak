"""Offline consistency checks. This is NOT a VBA compiler or an Excel test."""
from pathlib import Path
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]


def check(root=ROOT):
    source = "\n".join(p.read_text(encoding="utf-8-sig") for p in (root / "src").glob("*.bas"))
    version = (root / "VERSION").read_text().strip()
    if not re.fullmatch(r"\d+\.\d+\.\d+", version):
        raise ValueError("VERSION must use major.minor.patch")
    if f'SheetFreak {version}"' not in source:
        raise ValueError("Help version does not match VERSION")
    ribbon = ET.parse(root / "ribbon/customUI.xml").getroot()
    buttons = [node for node in ribbon.iter() if "onAction" in node.attrib]
    for button in buttons:
        callback = button.attrib["onAction"]
        if not re.search(r"Public Sub " + re.escape(callback) + r"\b", source, re.I):
            raise ValueError(f"Ribbon callback missing: {callback}")
    ids = [node.attrib["id"] for node in ribbon.iter() if "id" in node.attrib]
    if len(ids) != len(set(ids)):
        raise ValueError("Duplicate ribbon ID")
    tips = [b.attrib.get("keytip") for b in buttons if "keytip" in b.attrib]
    if len(tips) != len(set(tips)):
        raise ValueError("Duplicate command key tip")
    # Catch routine-boundary errors such as Exit Sub inside a Function.
    for path in [*(root / "src").glob("*.bas"), *(root / "tests").glob("*.bas")]:
        routine = None
        for number, line in enumerate(path.read_text().splitlines(), 1):
            text = line.strip()
            opening = re.match(r"(?:Public |Private )?(Sub|Function)\s+\w+", text, re.I)
            closing = re.match(r"End (Sub|Function)$", text, re.I)
            if opening:
                if routine:
                    raise ValueError(f"Nested routine: {path.name}:{number}")
                routine = opening[1].lower()
            if re.search(r"\bExit (Sub|Function)\b", text, re.I) and not text.startswith("'"):
                for kind in re.findall(r"\bExit (Sub|Function)\b", text, re.I):
                    if kind.lower() != routine:
                        raise ValueError(f"Wrong Exit: {path.name}:{number}")
            if closing:
                if closing[1].lower() != routine:
                    raise ValueError(f"Unmatched routine end: {path.name}:{number}")
                routine = None
        if routine:
            raise ValueError(f"Unclosed routine: {path.name}")
    if re.search(r"Application[.]OnKey\b", source, re.I):
        raise ValueError("SheetFreak must not register or clear global shortcuts")
    return version


if __name__ == "__main__":
    print(f"PASS: {check()} source/ribbon consistency; native Excel validation still required")
