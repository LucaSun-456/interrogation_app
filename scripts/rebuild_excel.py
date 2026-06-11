import os
import re
import sys
import xml.etree.ElementTree as ET
from openpyxl import Workbook

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def _col_letters_to_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n

def _sheet_xml_to_rows(xml_path: str) -> list[list]:
    with open(xml_path, 'r', encoding='utf-8') as f:
        xml_data = f.read()
        
    # Some XMLs might have garbage at the end if truncated. 
    # Try to fix it by finding the last valid closing tag.
    try:
        root = ET.fromstring(xml_data)
    except ET.ParseError as e:
        print(f"Parse error in {xml_path}: {e}. Attempting to truncate to last valid row...")
        # Find the last </row> tag
        idx = xml_data.rfind('</row>')
        if idx != -1:
            fixed_xml = xml_data[:idx+6] + '</sheetData></worksheet>'
            try:
                root = ET.fromstring(fixed_xml)
                print("Successfully parsed truncated XML.")
            except ET.ParseError as e2:
                print(f"Still failed to parse: {e2}")
                return []
        else:
            return []

    rows_map: dict[int, dict[int, str]] = {}
    for row in root.findall(".//m:sheetData/m:row", NS):
        r_idx = int(row.attrib.get("r", "0"))
        rows_map.setdefault(r_idx, {})
        for c in row.findall("m:c", NS):
            ref = c.attrib.get("r", "")
            m = re.match(r"^([A-Z]+)(\d+)$", ref)
            if not m:
                continue
            col_idx = _col_letters_to_index(m.group(1))
            cell_type = c.attrib.get("t")
            v = c.find("m:v", NS)
            is_elem = c.find("m:is", NS)
            val = ""
            if cell_type == "inlineStr" and is_elem is not None:
                t = is_elem.find(".//m:t", NS)
                val = t.text if t is not None and t.text else ""
            elif v is not None and v.text is not None:
                val = v.text
            rows_map[r_idx][col_idx] = val
            
    if not rows_map:
        return []
    max_row = max(rows_map)
    max_col = max(max(cols) for cols in rows_map.values())
    table = []
    for r in range(1, max_row + 1):
        line = []
        cols = rows_map.get(r, {})
        for c in range(1, max_col + 1):
            line.append(cols.get(c, ""))
        table.append(line)
    return table

def main():
    extracted_dir = sys.argv[1]
    out_path = sys.argv[2]
    
    sheet_dir = os.path.join(extracted_dir, 'xl', 'worksheets')
    if not os.path.exists(sheet_dir):
        print("No worksheets found.")
        return
        
    sheet_files = sorted([f for f in os.listdir(sheet_dir) if f.startswith('sheet') and f.endswith('.xml')])
    
    wb = Workbook()
    wb.remove(wb.active)
    
    # We know the app's sheet order:
    # 1. Participants
    # 2. Appointments
    # 3. AppointmentSlots
    # 4. Questionnaires
    # 5. QuestionnaireAnswers
    # 6. TrainingSessions
    # 7. SystemMeta
    expected_names = [
        "Participants",
        "Appointments",
        "AppointmentSlots",
        "Questionnaires",
        "QuestionnaireAnswers",
        "TrainingSessions",
        "SystemMeta"
    ]
    
    for i, sf in enumerate(sheet_files):
        print(f"Parsing {sf}...")
        rows = _sheet_xml_to_rows(os.path.join(sheet_dir, sf))
        if not rows:
            print(f"  No rows found or failed to parse {sf}.")
            continue
            
        title = expected_names[i] if i < len(expected_names) else f"sheet_{i+1}"
        ws = wb.create_sheet(title)
        for row in rows:
            ws.append(row)
        print(f"  Added {len(rows)} rows to {title}.")
        
    wb.save(out_path)
    print(f"Saved recovered workbook to {out_path}")

if __name__ == '__main__':
    main()
