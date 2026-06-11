import os
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
d = r'e:\Working Documents\【Important】My Projects\24 Specific avatar\interrogation_app\data\recovered_xml\xl\worksheets'

for f in sorted(os.listdir(d)):
    if not f.endswith('.xml'): continue
    path = os.path.join(d, f)
    with open(path, 'r', encoding='utf-8') as file:
        xml_data = file.read()
    
    # Try to fix truncated XML
    try:
        root = ET.fromstring(xml_data)
    except ET.ParseError:
        idx = xml_data.rfind('</row>')
        if idx != -1:
            fixed_xml = xml_data[:idx+6] + '</sheetData></worksheet>'
            try:
                root = ET.fromstring(fixed_xml)
            except:
                print(f"{f}: Failed to parse")
                continue
        else:
            print(f"{f}: Failed to parse")
            continue
            
    row1 = root.find(".//m:row[@r='1']", NS)
    if row1 is not None:
        headers = []
        for c in row1.findall("m:c", NS):
            t = c.find(".//m:t", NS)
            headers.append(t.text if t is not None else "")
        print(f"{f}: {headers}")
    else:
        print(f"{f}: No row 1")
