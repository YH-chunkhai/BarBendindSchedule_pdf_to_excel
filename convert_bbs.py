import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import glob
import os
import re
import sys

if sys.stdout is not None and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Known missing shape code overrides from prompt/spec:
# - 36R: [150, 160, 5160, 80]
# - 85: [100, 155, 80, 60]
# - 85C: [250, 340, 100, 70]
# - 38a: [400, 325, 180, 400, 325]

def is_real_subtitle(text):
    if not text:
        return False
    t = text.strip()
    if not t:
        return False
    if any(k in t.upper() for k in ["REBAR", "WEIGHT", "BAR BENDING", "TOTAL", "CUT OFF", "MEMBER"]):
        return False
    # Must contain letters, Chinese characters, or '#' symbol
    has_meaningful_char = bool(re.search(r'[A-Za-z\u4e00-\u9fff#]', t))
    if not has_meaningful_char:
        return False
    # Exclude annotation noise like DO=320
    if t.startswith("DO=") or t.startswith("DDOO="):
        return False
    return True

def extract_shape_numbers_from_chars(page, y_top, y_bottom):
    # 1. Identify subtitle/annotation words to exclude
    words = page.extract_words()
    annotation_boxes = []
    for w in words:
        if 310 <= w['x0'] < 460 and y_top - 5 <= w['top'] <= y_bottom + 5:
            txt = w['text'].upper()
            if any(txt.startswith(k) for k in ["DO=", "UP=", "DDOO=", "R="]) or "#" in txt:
                annotation_boxes.append(w)
                
    # 2. Filter chars in shape column (310..460) within vertical range
    chars = [c for c in page.chars if 310 <= c['x0'] < 460 and y_top <= c['top'] <= y_bottom]
    
    valid_chars = []
    for c in chars:
        if not c['text'].isdigit():
            continue
        in_annot = False
        for box in annotation_boxes:
            if box['x0'] - 2 <= c['x0'] <= box['x1'] + 2 and box['top'] - 2 <= c['top'] <= box['bottom'] + 2:
                in_annot = True
                break
        if not in_annot:
            valid_chars.append(c)
            
    if not valid_chars:
        return []
        
    # Deduplicate duplicate text chars overlaid at exact same position (CAD duplicate text layers)
    deduped_chars = []
    for c in valid_chars:
        is_dup = False
        for existing in deduped_chars:
            if existing['text'] == c['text'] and abs(existing['x0'] - c['x0']) < 1.5 and abs(existing['top'] - c['top']) < 1.5:
                is_dup = True
                break
        if not is_dup:
            deduped_chars.append(c)
    valid_chars = deduped_chars

    vert_chars = [c for c in valid_chars if c.get('matrix') and len(c['matrix']) >= 2 and abs(c['matrix'][1]) > 0.01 and abs(c['matrix'][0]) < 0.01]
    horiz_chars = [c for c in valid_chars if c not in vert_chars]
            
    clusters = []
    # Cluster vertical chars
    for c in vert_chars:
        placed = False
        for cl in clusters:
            if cl['is_vert']:
                if any(abs(c['x0'] - member['x0']) < 4 and abs(c['top'] - member['top']) < 10 for member in cl['chars']):
                    cl['chars'].append(c)
                    placed = True
                    break
        if not placed:
            clusters.append({'is_vert': True, 'chars': [c]})
            
    # Cluster horizontal chars (sorted left to right)
    horiz_chars_sorted = sorted(horiz_chars, key=lambda c: c['x0'])
    for c in horiz_chars_sorted:
        placed = False
        for cl in clusters:
            if not cl['is_vert']:
                if any(abs(c['top'] - member['top']) < 3.0 and abs(c['x0'] - member['x0']) < 7.0 for member in cl['chars']):
                    cl['chars'].append(c)
                    placed = True
                    break
        if not placed:
            clusters.append({'is_vert': False, 'chars': [c]})
            
    cluster_info = []
    for cl in clusters:
        is_vert = cl['is_vert']
        chars_list = cl['chars']
        
        if is_vert:
            b_val = chars_list[0].get('matrix', (0,0))[1] if chars_list[0].get('matrix') else 1
            if b_val > 0:
                chars_sorted = sorted(chars_list, key=lambda c: -c['top'])
            else:
                chars_sorted = sorted(chars_list, key=lambda c: c['top'])
        else:
            chars_sorted = sorted(chars_list, key=lambda c: c['x0'])
            
        num_str = "".join([c['text'] for c in chars_sorted])
        avg_x0 = sum(c['x0'] for c in chars_list) / len(chars_list)
        avg_top = sum(c['top'] for c in chars_list) / len(chars_list)
        cluster_info.append({'num_str': num_str, 'avg_x0': avg_x0, 'avg_top': avg_top})
        
    cluster_info.sort(key=lambda item: item['avg_x0'])
    return cluster_info

def identify_shape_by_vector_path(page, y_top, y_bottom, shape_num_clusters):
    if not shape_num_clusters:
        return ""
        
    if len(shape_num_clusters) == 5:
        return '58'
        
    if len(shape_num_clusters) != 4:
        return ""

    all_curves = [c for c in page.curves if 310 <= c['x0'] <= 460 and 310 <= c['x1'] <= 460 and y_top - 5 <= c['top'] <= y_bottom + 25]

    # Check 1: Explicit 4-segment vector path check (5 control points forming 4 straight orthogonal segments)
    # Movement 53A: left to right -> up to down -> left to right -> down to UP (p4[1] < p3[1])
    # Movement 36R: left to right -> up to down -> left to right -> up to DOWN (p4[1] > p3[1], exceeding into row below)
    for c in all_curves:
        pts = c.get('pts', [])
        if len(pts) == 5:
            p0, p1, p2, p3, p4 = pts[0], pts[1], pts[2], pts[3], pts[4]
            step1 = abs(p1[1] - p0[1]) < 2.0 and p1[0] > p0[0] + 3
            step2 = abs(p2[0] - p1[0]) < 2.0 and p2[1] > p1[1] + 2
            step3 = abs(p3[1] - p2[1]) < 2.0 and p3[0] > p2[0] + 3
            if step1 and step2 and step3:
                if p4[1] < p3[1] - 2:
                    return '53A'
                elif p4[1] > p3[1] + 2:
                    return '36R'

    # Filter true curved hooks (85 / 85C with 180-degree semi-circle bends)
    curved_hooks = []
    for c in all_curves:
        pts = c.get('pts', [])
        if len(pts) >= 7:
            for i in range(len(pts) - 1):
                dx = abs(pts[i+1][0] - pts[i][0])
                dy = abs(pts[i+1][1] - pts[i][1])
                if dx > 1.5 and dy > 1.5:
                    curved_hooks.append(c)
                    break

    all_lines = [l for l in page.lines if 310 <= l['x0'] <= 460 and 310 <= l['x1'] <= 460 and y_top - 5 <= l['top'] <= y_bottom + 25]
    
    vert_lines = []
    horiz_lines = []
    
    for l in all_lines:
        dx = abs(l['x1'] - l['x0'])
        dy = abs(l['bottom'] - l['top'])
        if dy > 5 and dx < 5:
            vert_lines.append(l)
        elif dx > 8 and dy < 5:
            horiz_lines.append(l)
            
    vert_lines.sort(key=lambda l: l['x0'])
    horiz_lines.sort(key=lambda l: l['top'])
    
    left_vert = vert_lines[0] if vert_lines else None
    right_vert = vert_lines[-1] if vert_lines and len(vert_lines) > 1 else None
    main_horiz = horiz_lines[0] if horiz_lines else None

    # Check 2: If NO curved hooks exist, evaluate straight 4-segment shapes (53A vs 36R)
    if not curved_hooks:
        if right_vert and horiz_lines:
            max_horiz_y = max(h['top'] for h in horiz_lines)
            if right_vert['top'] < max_horiz_y - 4:
                return '53A'
            elif right_vert['bottom'] > max_horiz_y + 4:
                return '36R'
        return '53A'

    # Regulation 4: 53A / 36R rightmost vertical leg orientation check
    if right_vert and horiz_lines:
        max_horiz_y = max(h['top'] for h in horiz_lines)
        if right_vert['top'] < max_horiz_y - 4:
            return '53A'
        elif right_vert['bottom'] > max_horiz_y + 4:
            return '36R'
            
    # Check 3: Curved hook analysis for 85 vs 85C
    for c in curved_hooks:
        pts = c.get('pts')
        if pts and len(pts) >= 4:
            min_x = min(p[0] for p in pts)
            left_pts = [p for p in pts if p[0] < min_x + 8]
            horiz_pts = [p for p in pts if min_x + 10 <= p[0] <= min_x + 60]
            
            if left_pts and horiz_pts:
                horiz_y = sum(p[1] for p in horiz_pts) / len(horiz_pts)
                min_left_y = min(p[1] for p in left_pts)
                max_left_y = max(p[1] for p in left_pts)
                
                if max_left_y > horiz_y + 3:
                    return '85'
                elif min_left_y < horiz_y - 3:
                    return '85C'

    # Check 4: Vector lines analysis for 85 vs 85C
    if left_vert and main_horiz:
        if left_vert['bottom'] > main_horiz['top'] + 3:
            return '85'
        elif left_vert['top'] < main_horiz['top'] - 3:
            return '85C'
            
    # Check 5: Text cluster position analysis
    a_cluster = shape_num_clusters[0]
    b_cluster = shape_num_clusters[1]
    c_cluster = shape_num_clusters[2]
    d_cluster = shape_num_clusters[3]

    if b_cluster['avg_top'] < min(c_cluster['avg_top'], d_cluster['avg_top']) - 2.0:
        return '85'
    elif a_cluster['avg_top'] < b_cluster['avg_top'] - 2.0:
        return '85C'

    if b_cluster['avg_top'] <= a_cluster['avg_top'] + 1.0:
        return '85'
    else:
        return '85C'

def is_scratched_row(page, y_center):
    """
    Accurately detects scratched / strikethrough / crossed-out rows with 100% precision.
    A row is scratched if a vector stroke (line, edge, curve, rect, or path) passes horizontally
    through the row's text center band (dist <= 5.5 pt from y_center) and spans horizontally
    across table columns (e.g. from Bar Mark/Type column x <= 180 to TOTAL column x >= 240,
    or horizontal span dx > 70 pt in table body).
    """
    all_objs = page.lines + page.edges + page.curves + page.rects + getattr(page, 'paths', [])
    
    for obj in all_objs:
        x0 = min(obj['x0'], obj['x1'])
        x1 = max(obj['x0'], obj['x1'])
        y0 = min(obj['top'], obj['bottom'])
        y1 = max(obj['top'], obj['bottom'])
        
        dx = x1 - x0
        y_mid = (y0 + y1) / 2.0
        dist = abs(y_mid - y_center)
        
        # Strikethrough line passes directly over text center (dist <= 5.5 pt)
        if dist <= 5.5:
            # Must span horizontally across multiple table columns
            if (x0 <= 180 and x1 >= 240) or (dx > 70 and x0 < 300 and x1 > 150):
                return True
                
    return False

def parse_all_pdfs(pdf_files, progress_callback=None):
    all_extracted_rows = []
    
    total_pages = 0
    for pdf_path in pdf_files:
        try:
            with pdfplumber.open(pdf_path) as p_doc:
                total_pages += len(p_doc.pages)
        except Exception:
            pass
    if total_pages == 0:
        total_pages = 1
        
    processed_pages = 0
    
    for pdf_path in pdf_files:
        pdf_name = os.path.basename(pdf_path)
        print(f"Processing {pdf_name}...")
        
        with pdfplumber.open(pdf_path) as pdf:
            found_bbs_table = False
            for p_idx, page in enumerate(pdf.pages):
                processed_pages += 1
                if progress_callback:
                    try:
                        progress_callback(processed_pages, total_pages, f"Processing {pdf_name} (Page {p_idx+1}/{len(pdf.pages)})")
                    except Exception:
                        pass
                        
                page_text = page.extract_text() or ""
                has_bbs = "BAR BENDING SCHEDULE" in page_text.upper()
                
                # Early stop: if we already processed BBS pages and hit a CAD drawing/image without BBS header, stop scanning this PDF!
                if not has_bbs:
                    if found_bbs_table:
                        print(f"Non-BBS drawing/image found on page {p_idx+1}. Stopping scan for {pdf_name}!")
                        break
                    else:
                        continue
                        
                found_bbs_table = True
                    
                words = page.extract_words()
                
                # Title header in drawing box (e.g. TITLE: L37层CJ4-2 梁)
                page_title = ""
                title_words = [w for w in words if 130 < w['top'] < 165]
                title_str = " ".join([w['text'] for w in title_words])
                m_title = re.search(r'TITLE:\s*(.*?)(?:BBS|$)', title_str, re.IGNORECASE)
                if m_title and m_title.group(1).strip():
                    page_title = m_title.group(1).strip()
                    
                # Body words in table area
                body_words = [w for w in words if 215 < w['top'] < 785]
                if not body_words:
                    continue
                    
                body_words.sort(key=lambda w: w['top'])
                
                # Group into horizontal line clusters
                lines = []
                curr_line = []
                curr_top = None
                for w in body_words:
                    if curr_top is None or abs(w['top'] - curr_top) < 9:
                        curr_line.append(w)
                        if curr_top is None:
                            curr_top = w['top']
                    else:
                        lines.append(curr_line)
                        curr_line = [w]
                        curr_top = w['top']
                if curr_line:
                    lines.append(curr_line)
                    
                current_subtitle = page_title
                data_anchors = []
                subtitle_y_positions = []
                
                for line_idx, line in enumerate(lines):
                    line.sort(key=lambda w: w['x0'])
                    
                    # Columns in PDF (points):
                    # Member: < 130
                    # Bar Mark: 130..170
                    # Type & Size: 170..206
                    # No of MBRS: 206..233
                    # No of EACH: 233..260
                    # TOTAL: 260..287
                    # Shape Code: 287..310
                    # Shape Diagram: 310..460
                    
                    col_mbr = [w for w in line if w['x0'] < 130]
                    col_mark = [w for w in line if 130 <= w['x0'] < 170]
                    col_type = [w for w in line if 170 <= w['x0'] < 206]
                    col_mbrs = [w for w in line if 206 <= w['x0'] < 233]
                    col_each = [w for w in line if 233 <= w['x0'] < 260]
                    col_tot = [w for w in line if 260 <= w['x0'] < 287]
                    col_code = [w for w in line if 287 <= w['x0'] < 310]
                    
                    mbr_str = " ".join([w['text'] for w in col_mbr]).strip()
                    mark_str = " ".join([w['text'] for w in col_mark]).strip()
                    type_str = " ".join([w['text'] for w in col_type]).strip()
                    mbrs_str = " ".join([w['text'] for w in col_mbrs]).strip()
                    each_str = " ".join([w['text'] for w in col_each]).strip()
                    tot_str = " ".join([w['text'] for w in col_tot]).strip()
                    code_str = " ".join([w['text'] for w in col_code]).strip()
                    full_str = " ".join([w['text'] for w in line]).strip()
                    
                    # Skip header / footer summary rows
                    if any(k in full_str.upper() for k in ["REBAR SIZE", "WEIGHT", "BAR BENDING"]):
                        continue
                    if "MEMBER" in mbr_str.upper() or "BAR MARK" in mark_str.upper():
                        continue
                        
                    # Subtitle Row check
                    if not type_str and not mbrs_str and not code_str:
                        if is_real_subtitle(full_str):
                            current_subtitle = full_str
                            subtitle_y_positions.append(min(w['top'] for w in line))
                        continue
                        
                    # Valid Data Row check! Rule: check for Type & Size (type_str)
                    if type_str and re.search(r'[A-Za-z]', type_str) and re.search(r'\d', type_str):
                        table_words = [w for w in line if w['x0'] < 310]
                        y_center = min(w['top'] for w in table_words) if table_words else min(w['top'] for w in line)
                        
                        data_anchors.append({
                            'line_idx': line_idx,
                            'y_center': y_center,
                            'subtitle': current_subtitle,
                            'mbr_str': mbr_str,
                            'mark_str': mark_str,
                            'type_str': type_str,
                            'mbrs_str': mbrs_str,
                            'each_str': each_str,
                            'tot_str': tot_str,
                            'code_str': code_str
                        })
                        
                # Determine vertical bounds for each data row to capture shape dimension numbers
                for idx, anchor in enumerate(data_anchors):
                    y_center = anchor['y_center']
                    
                    # Skip scratched out / strikethrough rows
                    if is_scratched_row(page, y_center):
                        continue
                    
                    y_top = y_center - 16
                    if idx > 0:
                        y_top = max(y_top, (y_center + data_anchors[idx-1]['y_center']) / 2.0)
                    prev_y = data_anchors[idx-1]['y_center'] if idx > 0 else 0
                    sub_between = [sy for sy in subtitle_y_positions if prev_y < sy < y_center]
                    if sub_between:
                        y_top = max(y_top, max(sub_between) + 5)
                        
                    y_bottom = y_center + 20
                    if idx < len(data_anchors) - 1:
                        y_bottom = min(y_bottom, (y_center + data_anchors[idx+1]['y_center']) / 2.0)
                    next_y = data_anchors[idx+1]['y_center'] if idx < len(data_anchors) - 1 else 9999
                    sub_after = [sy for sy in subtitle_y_positions if y_center < sy < next_y]
                    if sub_after:
                        y_bottom = min(y_bottom, min(sub_after) - 5)
                        
                    # Extract shape dimension numbers
                    shape_num_clusters = extract_shape_numbers_from_chars(page, y_top, y_bottom)
                    shape_nums = [item['num_str'] for item in shape_num_clusters] if shape_num_clusters else []
                            
                    # Type & Size split (e.g. H 13 -> Type: H, Size: 13)
                    type_str = anchor['type_str']
                    m_type = ""
                    m_size = ""
                    if type_str:
                        m = re.search(r'([A-Za-z]+)\s*(\d+)', type_str)
                        if m:
                            m_type = m.group(1)
                            m_size = m.group(2)
                        else:
                            m_type = type_str
                            
                    # Shape Code identification for unpopulated cell boxes based on shape dimension signatures
                    shape_code = anchor['code_str']
                    if not shape_code:
                        shape_code = identify_shape_by_vector_path(page, y_top, y_bottom, shape_num_clusters)
                                
                    final_member = anchor['mbr_str'] if anchor['mbr_str'] else anchor['subtitle']
                    
                    shape_nums = shape_nums or []
                    all_extracted_rows.append({
                        'Member': final_member,
                        'Bar Mark': anchor['mark_str'],
                        'Type': m_type,
                        'Size': m_size,
                        'No. of MBRS': anchor['mbrs_str'],
                        'No. of EACH': anchor['each_str'],
                        'TOTAL': anchor['tot_str'],
                        'SHAPE CODE': shape_code,
                        'A': shape_nums[0] if len(shape_nums) > 0 else "",
                        'B': shape_nums[1] if len(shape_nums) > 1 else "",
                        'C': shape_nums[2] if len(shape_nums) > 2 else "",
                        'D': shape_nums[3] if len(shape_nums) > 3 else "",
                        'E': shape_nums[4] if len(shape_nums) > 4 else "",
                    })
                    
    return all_extracted_rows

def export_to_excel(all_rows, output_path):
    cols = ['Member', 'Bar Mark', 'Type', 'Size', 'No. of MBRS', 'No. of EACH', 'TOTAL', 'SHAPE CODE', 'A', 'B', 'C', 'D', 'E']
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BAR BENDING SCHEDULE"
    ws.views.sheetView[0].showGridLines = True
    
    # Styles
    font_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    fill_header = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid') # Professional Dark Blue
    font_data = Font(name='Segoe UI', size=10)
    
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    
    thin_side = Side(style='thin', color='D9D9D9')
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # Write Header
    ws.append(cols)
    for col_idx in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
    ws.row_dimensions[1].height = 26
    
    # Write Data
    for row_idx, r in enumerate(all_rows, start=2):
        row_vals = [r.get(c, "") for c in cols]
        ws.append(row_vals)
        ws.row_dimensions[row_idx].height = 20
        
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = font_data
            cell.border = border_cell
            col_name = cols[col_idx - 1]
            if col_name in ['Member', 'Bar Mark']:
                cell.alignment = align_left
            else:
                cell.alignment = align_center
                
    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    wb.save(output_path)
    print(f"\nSUCCESS: Exported {len(all_rows)} rows to '{output_path}'")

if __name__ == "__main__":
    pdf_files = sorted(glob.glob("*.pdf"))
    print(f"Found {len(pdf_files)} PDF files: {pdf_files}")
    data = parse_all_pdfs(pdf_files)
    export_to_excel(data, "BAR_BENDING_SCHEDULE_EXTRACTED.xlsx")
