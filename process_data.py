import os
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from xlsx_writer import write_xlsx

def load_sheet_all(file_path):
    print(f"Loading {file_path}...")
    with zipfile.ZipFile(file_path, 'r') as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in tree.findall('{*}si'):
                text_parts = [t.text for t in si.findall('.//{*}t') if t.text]
                shared_strings.append(''.join(text_parts))
        
        sheet_path = 'xl/worksheets/sheet1.xml'
        sheet_tree = ET.fromstring(z.read(sheet_path))
        
        headers_ordered = []
        rows_data = []
        
        for row_elem in sheet_tree.findall('.//{*}row'):
            row_idx = int(row_elem.get('r'))
            cells = {}
            for c in row_elem.findall('{*}c'):
                r = c.get('r')
                col_letters = ''.join([ch for ch in r if ch.isalpha()])
                t = c.get('t')
                v = c.find('{*}v')
                val = v.text if v is not None else ''
                if t == 's' and val != '':
                    idx = int(val)
                    if idx < len(shared_strings):
                        val = shared_strings[idx]
                elif t == 'inlineStr':
                    is_t = c.find('.//{*}t')
                    val = is_t.text if is_t is not None else ''
                cells[col_letters] = val.strip() if isinstance(val, str) else val
            
            if row_idx == 1:
                def col_key(s):
                    num = 0
                    for ch in s:
                        num = num * 26 + (ord(ch) - ord('A') + 1)
                    return num
                sorted_cols = sorted(cells.keys(), key=col_key)
                headers_ordered = [cells[c] for c in sorted_cols]
            else:
                rows_data.append(cells)
                
    return headers_ordered, rows_data

def clean_coordinates(coord_str):
    if not coord_str or coord_str.strip() in [',', ';', '', 'None', 'null']:
        return ('', '', 'غير متوفرة', '')
    
    cleaned = coord_str.replace(';', ',').strip()
    parts = cleaned.split(',')
    if len(parts) != 2:
        return ('', '', 'تنسيق غير صالح', '')
    
    try:
        p1 = float(parts[0].strip())
        p2 = float(parts[1].strip())
    except ValueError:
        return ('', '', 'قيمة غير رقمية', '')
    
    # Check if swapped (Lon, Lat) vs (Lat, Lon)
    if 34 <= p1 <= 56 and 16 <= p2 <= 32:
        lat = p2
        lon = p1
        status = 'تم تصحيح التبديل (طول/عرض)'
    elif 16 <= p1 <= 32 and 34 <= p2 <= 56:
        lat = p1
        lon = p2
        status = 'صحيحة'
    else:
        lat = p1
        lon = p2
        status = 'إحداثيات خارج النطاق الجغرافي'
        
    gmaps = f"https://www.google.com/maps?q={lat:.6f},{lon:.6f}"
    return (f"{lat:.7f}", f"{lon:.7f}", status, gmaps)

def process_all():
    sectors_config = [
        {
            'sector_name': 'القطاع الغربي',
            'raw_file': 'الغربي.xlsx',
            'out_file': 'بيانات_القطاع_الغربي_بعد_المعالجة.xlsx'
        },
        {
            'sector_name': 'القطاع الشمالي',
            'raw_file': 'القطاع الشمالي.xlsx',
            'out_file': 'بيانات_القطاع_الشمالي_بعد_المعالجة.xlsx'
        },
        {
            'sector_name': 'القطاع الجنوبي',
            'raw_file': 'الجنوبي.xlsx',
            'out_file': 'بيانات_القطاع_الجنوبي_بعد_المعالجة.xlsx'
        }
    ]
    
    all_sectors_summary = []
    combined_sheets = {}

    for cfg in sectors_config:
        sector_name = cfg['sector_name']
        raw_file = cfg['raw_file']
        out_file = cfg['out_file']
        
        headers, raw_rows = load_sheet_all(raw_file)
        
        def col_to_letter(idx):
            res = ""
            while idx > 0:
                idx, rem = divmod(idx - 1, 26)
                res = chr(65 + rem) + res
            return res
        
        header_name_to_idx = {h: idx for idx, h in enumerate(headers)}
        
        visit_col_idx = header_name_to_idx.get('رقم الزيارة', 0)
        coord_col_idx = header_name_to_idx.get('الاحداثيات', 16)
        muni_col_idx = header_name_to_idx.get('البلدية', 2)
        insp_col_idx = header_name_to_idx.get('اسم المراقب', 3)
        
        seen_visits = set()
        dedup_records = []
        dup_count = 0
        
        for r_cells in raw_rows:
            row_vals = []
            for c_idx in range(1, len(headers) + 1):
                col_let = col_to_letter(c_idx)
                row_vals.append(r_cells.get(col_let, ''))
                
            vid = row_vals[visit_col_idx]
            if vid in seen_visits:
                dup_count += 1
                continue
            seen_visits.add(vid)
            
            orig_coords = row_vals[coord_col_idx]
            lat, lon, coord_status, gmaps = clean_coordinates(orig_coords)
            
            row_vals.append(lat)
            row_vals.append(lon)
            row_vals.append(coord_status)
            row_vals.append(gmaps)
            
            dedup_records.append(row_vals)
            
        new_headers = list(headers) + ['خط العرض (معالج)', 'خط الطول (معالج)', 'حالة الإحداثيات', 'رابط خرائط جوجل']
        
        total_visits = len(dedup_records)
        all_inspectors = set(r[insp_col_idx] for r in dedup_records)
        total_inspectors = len(all_inspectors)
        
        coord_stats = defaultdict(int)
        for r in dedup_records:
            coord_stats[r[-2]] += 1
            
        print(f"\n=======================================================")
        print(f"تقرير المعالجة: {sector_name}")
        print(f"=======================================================")
        print(f"السجلات الأصلية: {len(raw_rows)}")
        print(f"عدد الزيارات المكررة المحذوفة: {dup_count}")
        print(f"إجمالي الزيارات الفريدة بعد المعالجة: {total_visits}")
        print(f"إجمالي المراقبين: {total_inspectors}")
        print(f"حالة الإحداثيات: {dict(coord_stats)}")
        
        muni_inspectors = defaultdict(lambda: defaultdict(int))
        for r in dedup_records:
            muni = r[muni_col_idx]
            insp = r[insp_col_idx]
            muni_inspectors[muni][insp] += 1
            
        summary_rows = [
            [f"ملخص مؤشرات {sector_name}", ""],
            ["المؤشر", "القيمة"],
            ["إجمالي الزيارات بالقطاع", total_visits],
            ["إجمالي عدد المراقبين بالقطاع", total_inspectors],
            ["عدد السجلات الأصلية", len(raw_rows)],
            ["عدد الزيارات المكررة المستبعدة", dup_count],
            ["الإحداثيات الصحيحة", coord_stats['صحيحة']],
            ["الإحداثيات غير المتوفرة (المعالجة)", coord_stats['غير متوفرة']],
            ["", ""],
            [f"جدول بلديات {sector_name}", "", "", ""],
            ["البلدية", "إجمالي الزيارات", "عدد المراقبين", "نسبة الزيارات من القطاع"]
        ]
        
        sorted_munis = sorted(muni_inspectors.items(), key=lambda x: sum(x[1].values()), reverse=True)
        for muni, insps in sorted_munis:
            m_visits = sum(insps.values())
            m_insps = len(insps)
            m_pct = f"{(m_visits / total_visits) * 100:.1f}%"
            summary_rows.append([muni, m_visits, m_insps, m_pct])
            
        summary_rows.append(["", ""])
        summary_rows.append([f"تفصيل المراقبين حسب كل بلدية - {sector_name}", "", "", ""])
        summary_rows.append(["البلدية", "اسم المراقب", "عدد الزيارات", "النسبة من زيارات البلدية", "النسبة من زيارات القطاع"])
        
        for muni, insps in sorted_munis:
            m_visits = sum(insps.values())
            sorted_insps = sorted(insps.items(), key=lambda x: x[1], reverse=True)
            for insp_name, count in sorted_insps:
                pct_muni = f"{(count / m_visits) * 100:.1f}%"
                pct_sector = f"{(count / total_visits) * 100:.1f}%"
                summary_rows.append([muni, insp_name, count, pct_muni, pct_sector])
                
        if sector_name == 'القطاع الجنوبي' or not os.path.exists(out_file):
            sheets_data = {
                'ملخص الزيارات والمراقبين': summary_rows,
                'الزيارات بعد المعالجة': [new_headers] + dedup_records
            }
            write_xlsx(out_file, sheets_data)
        
        combined_sheets[f"ملخص_{sector_name}"] = summary_rows
        all_sectors_summary.append({
            'sector': sector_name,
            'total_visits': total_visits,
            'total_inspectors': total_inspectors,
            'raw_count': len(raw_rows),
            'dup_count': dup_count,
            'munis': sorted_munis
        })
        
    tot_visits = sum(s['total_visits'] for s in all_sectors_summary)
    tot_insps = sum(s['total_inspectors'] for s in all_sectors_summary)
    tot_raw = sum(s['raw_count'] for s in all_sectors_summary)
    tot_dup = sum(s['dup_count'] for s in all_sectors_summary)
    tot_munis = sum(len(s['munis']) for s in all_sectors_summary)
    
    master_rows = [
        ["تقرير المقارنة الشامل بين القطاعات الرقابية بمنطقة عسير", "", "", "", ""],
        ["البيان", "القطاع الغربي", "القطاع الشمالي", "القطاع الجنوبي", "الإجمالي العام"],
        ["إجمالي الزيارات الفريدة", all_sectors_summary[0]['total_visits'], all_sectors_summary[1]['total_visits'], all_sectors_summary[2]['total_visits'], tot_visits],
        ["إجمالي عدد المراقبين", all_sectors_summary[0]['total_inspectors'], all_sectors_summary[1]['total_inspectors'], all_sectors_summary[2]['total_inspectors'], tot_insps],
        ["عدد السجلات الأصلية", all_sectors_summary[0]['raw_count'], all_sectors_summary[1]['raw_count'], all_sectors_summary[2]['raw_count'], tot_raw],
        ["عدد الزيارات المكررة المستبعدة", all_sectors_summary[0]['dup_count'], all_sectors_summary[1]['dup_count'], all_sectors_summary[2]['dup_count'], tot_dup],
        ["عدد البلديات التابعة للقطاع", len(all_sectors_summary[0]['munis']), len(all_sectors_summary[1]['munis']), len(all_sectors_summary[2]['munis']), tot_munis],
        ["متوسط الزيارات لكل مراقب", 
         f"{all_sectors_summary[0]['total_visits']/all_sectors_summary[0]['total_inspectors']:.1f}", 
         f"{all_sectors_summary[1]['total_visits']/all_sectors_summary[1]['total_inspectors']:.1f}", 
         f"{all_sectors_summary[2]['total_visits']/all_sectors_summary[2]['total_inspectors']:.1f}", 
         f"{tot_visits/tot_insps:.1f}"]
    ]
    
    master_sheets = {
        'المقارنة العامة للقطاعات': master_rows,
        'القطاع الجنوبي - البلديات والمراقبين': combined_sheets['ملخص_القطاع الجنوبي'],
        'القطاع الغربي - البلديات والمراقبين': combined_sheets['ملخص_القطاع الغربي'],
        'القطاع الشمالي - البلديات والمراقبين': combined_sheets['ملخص_القطاع الشمالي']
    }
    write_xlsx('تقرير_شامل_مراقبين_القطاعات.xlsx', master_sheets)
    print("\nMaster files updated successfully!")

if __name__ == '__main__':
    process_all()
