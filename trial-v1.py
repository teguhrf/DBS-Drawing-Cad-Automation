import openpyxl
import re
from pyproj import Proj

def process_and_route(input_file, output_file):
    # Setup UTM Zone 47S Projection
    utm_proj = Proj(proj='utm', zone=47, south=True, ellps='WGS84')
    
    print(f"Reading {input_file}...")
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active
    
    pole_data = {}
    pole_list = []
    
    # --- PASS 1: Read coordinates, pole types, and umbang ---
    for row in range(1, ws.max_row + 1):
        coord_str = ws.cell(row=row, column=1).value       # Column A: Coordinate
        main_pole_name = ws.cell(row=row, column=2).value  # Column B: Full Pole Name
        
        if not coord_str or not main_pole_name:
            continue
            
        try:
            # Parse and convert coordinates
            lat_str, lon_str = str(coord_str).split(',')
            lat = float(lat_str.strip())
            lon = float(lon_str.strip())
            x, y = utm_proj(lon, lat)
            
            pole_name_clean = str(main_pole_name).strip()
            
            # --- UMBANG LOGIC CHECKING (Column D) ---
            umbang_raw = ws.cell(row=row, column=4).value
            umbang_val = str(umbang_raw).strip() if umbang_raw is not None else ""
            
            umbang_block = None
            # Validate if it's a number between 1 and 9
            if umbang_val.isdigit() and 1 <= int(umbang_val) <= 9:
                umbang_block = f"U{umbang_val}"
            
            # --- POLE TYPE LOGIC CHECKING (Columns H to K) ---
            val_h = str(ws.cell(row=row, column=8).value).strip()
            val_i = str(ws.cell(row=row, column=9).value).strip()
            val_j = str(ws.cell(row=row, column=10).value).strip()
            val_k = str(ws.cell(row=row, column=11).value).strip()
            
            active_flags = []
            if val_h in ["1", "1.0"]: active_flags.append("TIANG_MERAH")
            if val_i in ["1", "1.0"]: active_flags.append("TIANG_BIRU")
            if val_j in ["1", "1.0"]: active_flags.append("TIANG_MAG")
            if val_k in ["1", "1.0"]: active_flags.append("TIANG_SP")
            
            # Fallback default if multiple or none selected
            block_name = active_flags[0] if len(active_flags) == 1 else "TIANG_BIRU" 
            
            # Store all validated data
            pole_data[pole_name_clean] = {
                'x': x, 
                'y': y, 
                'row': row, 
                'block': block_name,
                'umbang': umbang_block
            }
            pole_list.append(pole_name_clean)
            
        except Exception:
            continue

    # --- PASS 2: Execute Parent-Child Routing Logic ---
    pole_set = set(pole_list)
    pole_relations = {}
    
    # Universal radar to find joint poles
    def get_joint_pole(missing_target):
        m = re.search(r'^(.*?\s+)([A-Z]+)\s+(\d+)$', missing_target)
        if m:
            prefix = m.group(1)
            feeder = m.group(2)
            num = m.group(3)
            for t in pole_set:
                m_t = re.search(r'^(.*?\s+)([A-Z]+)\s+' + num + r'$', t)
                if m_t and prefix == m_t.group(1) and feeder in m_t.group(2):
                    return t
        return None

    for pole in pole_list:
        parent = None
        
        match_letter = re.search(r'^(.*)\s+/\d+[A-Za-z]$', pole)
        match_child = re.search(r'^(.*)\s+/(\d+)$', pole)
        match_main = re.search(r'^(.*?\s+)([A-Z]+)\s+(\d+)$', pole)
        
        if match_letter:
            target_parent = match_letter.group(1).strip()
            parent = target_parent if target_parent in pole_set else (get_joint_pole(target_parent) or target_parent)
        elif match_child:
            base = match_child.group(1).strip()
            num = int(match_child.group(2))
            if num != 1:
                target_parent = f"{base} /{num - 1}"
                parent = target_parent if target_parent in pole_set else target_parent
            else:
                target_parent = base
                parent = target_parent if target_parent in pole_set else (get_joint_pole(target_parent) or target_parent)
        elif match_main:
            prefix = match_main.group(1)
            feeder = match_main.group(2)
            num = int(match_main.group(3))
            if num != 1:
                target_parent = f"{prefix}{feeder} {num - 1}"
                parent = target_parent if target_parent in pole_set else (get_joint_pole(target_parent) or target_parent)
                
        pole_relations[pole] = parent

    # --- PASS 3: Generate AutoCAD Commands ---
    for pole, data in pole_data.items():
        row = data['row']
        x = data['x']
        y = data['y']
        block_name = data['block']
        umbang_block = data['umbang']
        
        text_x = x + 2
        text_y = y + 1
        
        # 1. Output to Column P (16): Main Pole Block Insert
        ws.cell(row=row, column=16).value = f'(command "-insert" "{block_name}" "{x:.3f},{y:.3f}" 1 1 0)'
        
        # --- DISPLAY TEXT FORMATTING ---
        # Bypass string splitting if the name explicitly starts with "PE "
        if pole.upper().startswith("PE "):
            display_label_b = pole
        else:
            display_name_match = re.search(r'([A-Z]+\s+\d+.*)$', pole)
            display_label_b = display_name_match.group(1) if display_name_match else pole 
            
        label_c = ws.cell(row=row, column=3).value
        
        # 2. Output to Column Q (17) and R (18): Text Labels
        if display_label_b: ws.cell(row=row, column=17).value = f"-TEXT {text_x:.3f},{text_y:.3f} 0 {display_label_b}"
        if label_c: ws.cell(row=row, column=18).value = f"-TEXT {text_x:.3f},{text_y:.3f} 0 {label_c}"
        
        # 3. Output to Column S (19): Line Routing Command
        parent = pole_relations.get(pole)
        if parent and parent in pole_data:
            x_parent = pole_data[parent]['x']
            y_parent = pole_data[parent]['y']
            ws.cell(row=row, column=19).value = f"LINE {x_parent:.3f},{y_parent:.3f} {x:.3f},{y:.3f} "

        # 4. Output to Column T (20): Umbang Block Insert (if exists)
        if umbang_block:
            ws.cell(row=row, column=20).value = f'(command "-insert" "{umbang_block}" "{x:.3f},{y:.3f}" 1 1 0)'

    wb.save(output_file)
    print(f"Done! Umbang logic and PE naming bypass generated successfully to {output_file}")

# --- EXECUTION ---
process_and_route("data_input.xlsx", "data_output.xlsx")