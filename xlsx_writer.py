import zipfile
import xml.etree.ElementTree as ET
import os

def write_xlsx(filename, sheets_dict):
    """
    sheets_dict: dict of sheet_name -> list of rows (each row is list of values)
    Writes a valid .xlsx file without external dependencies.
    """
    def col_to_letter(col_idx):
        result = ""
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def escape_xml(s):
        if s is None:
            return ""
        s = str(s)
        return (s.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;")
                 .replace('"', "&quot;")
                 .replace("'", "&apos;"))

    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    content_types.append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    content_types.append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    content_types.append('<Default Extension="xml" ContentType="application/xml"/>')
    content_types.append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')

    workbook_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    workbook_xml.append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    workbook_xml.append('<sheets>')

    wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    wb_rels.append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')

    sheet_xmls = {}

    for idx, (sheet_name, rows) in enumerate(sheets_dict.items(), start=1):
        r_id = f"rId{idx}"
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        workbook_xml.append(f'<sheet name="{escape_xml(sheet_name)}" sheetId="{idx}" r:id="{r_id}"/>')
        wb_rels.append(f'<Relationship Id="{r_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')

        ws = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        ws.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        ws.append('<sheetViews><sheetView tabSelected="1" workbookViewId="0" rightToLeft="1"/></sheetViews>')
        ws.append('<sheetData>')
        
        for r_idx, row in enumerate(rows, start=1):
            ws.append(f'<row r="{r_idx}">')
            for c_idx, val in enumerate(row, start=1):
                cell_ref = f"{col_to_letter(c_idx)}{r_idx}"
                if val is None or val == '':
                    continue
                # check if numeric (int or float)
                is_num = False
                if isinstance(val, (int, float)):
                    is_num = True
                elif isinstance(val, str):
                    # check if string is clean numeric (not phone number starting with 0, not national ID)
                    # For phone/ID, keep as string!
                    if not (val.startswith('0') or val.startswith('INS-') or val.startswith('INC-') or len(val) >= 10):
                        try:
                            f = float(val)
                            is_num = True
                        except ValueError:
                            pass
                
                if is_num:
                    ws.append(f'<c r="{cell_ref}"><v>{val}</v></c>')
                else:
                    ws.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{escape_xml(val)}</t></is></c>')
            ws.append('</row>')
        ws.append('</sheetData>')
        ws.append('</worksheet>')
        sheet_xmls[f"xl/worksheets/sheet{idx}.xml"] = "".join(ws)

    workbook_xml.append('</sheets></workbook>')
    wb_rels.append('</Relationships>')
    content_types.append('</Types>')

    root_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    root_rels.append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    root_rels.append('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>')
    root_rels.append('</Relationships>')

    with zipfile.ZipFile(filename, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', "".join(content_types))
        z.writestr('_rels/.rels', "".join(root_rels))
        z.writestr('xl/workbook.xml', "".join(workbook_xml))
        z.writestr('xl/_rels/workbook.xml.rels', "".join(wb_rels))
        for sheet_path, sheet_xml in sheet_xmls.items():
            z.writestr(sheet_path, sheet_xml)
    print(f"Successfully generated {filename}")

if __name__ == '__main__':
    write_xlsx('test_output.xlsx', {'تجربة': [['الاسم', 'القيمة'], ['اختبار', 123]]})

