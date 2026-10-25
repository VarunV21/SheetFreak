"""Add the ribbon to an Excel-produced XLAM; no VBA project access needed."""
import argparse, zipfile, re, xml.etree.ElementTree as ET
from pathlib import Path

def package(source, output, ribbon):
    source, output, ribbon = map(Path, (source, output, ribbon))
    if source.resolve() == output.resolve():
        raise ValueError('Use different source and destination paths')
    ns='http://schemas.openxmlformats.org/package/2006/relationships'
    with zipfile.ZipFile(source) as z:
        entries={i.filename:z.read(i) for i in z.infolist()}
    if 'xl/vbaProject.bin' not in entries:
        raise ValueError('Source has no compiled VBA project')
    root=ET.fromstring(entries['_rels/.rels'])
    relation_type='http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
    for rel in list(root):
        if rel.get('Type') == relation_type:
            root.remove(rel)
    used={r.get('Id') for r in root}
    rid='rIdSheetFreak'
    if rid in used:
        raise ValueError('Ribbon relation id already exists')
    ET.SubElement(root, '{'+ns+'}Relationship', Id=rid, Type=relation_type, Target='customUI/customUI.xml')
    entries['_rels/.rels']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    entries['customUI/customUI.xml']=ribbon.read_bytes()
    ET.fromstring(entries['customUI/customUI.xml'])
    # Strip author metadata from the distributable, not from the source workbook.
    if 'docProps/core.xml' in entries:
        # Preserve namespace declarations used by qualified attribute values such
        # as xsi:type="dcterms:W3CDTF"; generic XML reserialization can drop them.
        props=entries['docProps/core.xml'].decode('utf-8')
        for tag in ('dc:creator','cp:lastModifiedBy'):
            props=re.sub(r'(<'+tag+r'\b[^>]*>).*?(</'+tag+r'>)',r'\g<1>SheetFreak contributors\g<2>',props,flags=re.S)
        entries['docProps/core.xml']=props.encode('utf-8')
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in entries.items(): z.writestr(name,data)
    print(f'Packaged {output.name}: {output.stat().st_size:,} bytes')

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('source');parser.add_argument('output');parser.add_argument('ribbon')
    args=parser.parse_args()
    package(args.source,args.output,args.ribbon)
