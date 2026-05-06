import openpyxl
import re
from pyproj import Proj

def process_and_route(input_file, output_excel, output_scr):
    utm_proj = Proj(proj='utm', zone=47, south=True, ellps='WGS84')
    
    print(f"Reading {input_file}...")
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active
    
    pole_data = {}
    pole_list = []
    
    # --- SETUP HEADER OUTPUT (REARRANGED) ---
    ws.cell(row=1, column=22).value = "Input_Kabel"         
    ws.cell(row=1, column=23).value = "Input_tiang"         
    ws.cell(row=1, column=24).value = "Input_Umbang"        
    ws.cell(row=1, column=25).value = "Input_No Tiang"      
    ws.cell(row=1, column=26).value = "Input_No Tiang Lama" 
    ws.cell(row=1, column=27).value = "Input_details"       
    
    # --- PASS 1: Read and Clean Data ---
    for row in range(2, ws.max_row + 1):  
        coord_str = ws.cell(row=row, column=1).value
        main_pole_name = ws.cell(row=row, column=2).value
        
        if not coord_str or not main_pole_name:
            continue
            
        try:
            # Parse Coordinates
            lat_str, lon_str = str(coord_str).split(',')
            lat = float(lat_str.strip())
            lon = float(lon_str.strip())
            x, y = utm_proj(lon, lat)
            
            pole_name_clean = str(main_pole_name).strip()
            
            # Old Pole Name (Column C / 3)
            old_pole_val = ws.cell(row=row, column=3).value
            old_pole_name = str(old_pole_val).strip() if old_pole_val else ""
            
            # Umbang Logic (Column D / 4)
            umbang_raw = ws.cell(row=row, column=4).value
            umbang_val = str(umbang_raw).strip() if umbang_raw is not None else ""
            umbang_block = f"U{umbang_val}" if umbang_val.isdigit() and 1 <= int(umbang_val) <= 9 else None
            
            # Details Concatenation (Columns E, F, G / 5, 6, 7)
            col_e = str(ws.cell(row=row, column=5).value).strip() if ws.cell(row=row, column=5).value is not None else ""
            col_f = str(ws.cell(row=row, column=6).value).strip() if ws.cell(row=row, column=6).value is not None else ""
            col_g = str(ws.cell(row=row, column=7).value).strip() if ws.cell(row=row, column=7).value is not None else ""
            details_text = " ".join(filter(None, [col_e, col_f, col_g]))
            
            # Pole Type Logic (Columns H, I, J, K / 8, 9, 10, 11)
            val_h = str(ws.cell(row=row, column=8).value).strip()
            val_i = str(ws.cell(row=row, column=9).value).strip()
            val_j = str(ws.cell(row=row, column=10).value).strip()
            val_k = str(ws.cell(row=row, column=11).value).strip()
            
            active_flags = []
            if val_h in ["1", "1.0"]: active_flags.append("TIANG_MERAH")
            if val_i in ["1", "1.0"]: active_flags.append("TIANG_BIRU")
            if val_j in ["1", "1.0"]: active_flags.append("TIANG_MAG")
            if val_k in ["1", "1.0"]: active_flags.append("TIANG_SP")
            
            block_name = active_flags[0] if len(active_flags) == 1 else "TIANG_BIRU"
            
            # CABLE CLEANING LOGIC (IN-PLACE ERASE & DEFAULT FALLBACK)
            dominant_found = False
            for col_idx in range(12, 22):
                cable_val = ws.cell(row=row, column=col_idx).value
                
                if not dominant_found and cable_val is not None and str(cable_val).strip() != "":
                    dominant_found = True
                elif dominant_found:
                    ws.cell(row=row, column=col_idx).value = None
            
            if not dominant_found:
                ws.cell(row=row, column=15).value = 1
            
            pole_data[pole_name_clean] = {
                'x': x, 'y': y, 'row': row, 
                'block': block_name, 'umbang': umbang_block,
                'details': details_text, 'old_pole': old_pole_name
            }
            pole_list.append(pole_name_clean)
            
        except Exception:
            continue

    # --- PASS 2: Execute Routing Logic ---
    pole_set = set(pole_list)
    pole_relations = {}
    
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
        
        # UPDATED REGEX: Pisahin base, angka, dan huruf suffix secara presisi
        match_suffix_child = re.search(r'^(.*)\s+/(\d+)([A-Za-z]+)$', pole)
        match_child = re.search(r'^(.*)\s+/(\d+)$', pole)
        match_main = re.search(r'^(.*?\s+)([A-Z]+)\s+(\d+)$', pole)
        
        if match_suffix_child:
            base = match_suffix_child.group(1).strip()
            num = int(match_suffix_child.group(2))
            suffix = match_suffix_child.group(3)
            
            if num != 1:
                # Contoh: A 10 /5B -> nyari A 10 /4B
                target_parent = f"{base} /{num - 1}{suffix}"
                parent = target_parent if target_parent in pole_set else (get_joint_pole(target_parent) or target_parent)
            else:
                # Contoh: A 10 /1B -> nyari A 10
                target_parent = base
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
    script_commands = ["OSMODE 0"] 
    
    for pole, data in pole_data.items():
        row = data['row']
        x = data['x']
        y = data['y']
        block_name = data['block']
        umbang_block = data['umbang']
        details_text = data['details']
        old_pole_text = data['old_pole']
        
        text_x = x + 2
        y_pole = y + 1
        y_lama = y
        y_details = y - 1
        
        # 1. Line Routing
        parent = pole_relations.get(pole)
        if parent and parent in pole_data:
            x_parent = pole_data[parent]['x']
            y_parent = pole_data[parent]['y']
            cmd_line = f"LINE {x_parent:.3f},{y_parent:.3f} {x:.3f},{y:.3f} "
            ws.cell(row=row, column=22).value = cmd_line
            script_commands.append(cmd_line)
            
        # 2. Pole Block
        cmd_insert = f'(command "-insert" "{block_name}" "{x:.3f},{y:.3f}" 1 1 0)'
        ws.cell(row=row, column=23).value = cmd_insert
        script_commands.append(cmd_insert)

        # 3. Umbang Block
        if umbang_block:
            cmd_umbang = f'(command "-insert" "{umbang_block}" "{x:.3f},{y:.3f}" 1 1 0)'
            ws.cell(row=row, column=24).value = cmd_umbang
            script_commands.append(cmd_umbang)
        
        # PE Naming Bypass
        if pole.upper().startswith("PE "):
            display_label_b = pole
        else:
            display_name_match = re.search(r'([A-Z]+\s+\d+.*)$', pole)
            display_label_b = display_name_match.group(1) if display_name_match else pole 
            
        # 4. Main Pole Text
        if display_label_b: 
            cmd_text_1 = f"-TEXT {text_x:.3f},{y_pole:.3f} 0 {display_label_b}"
            ws.cell(row=row, column=25).value = cmd_text_1
            script_commands.append(cmd_text_1)
        
        # 5. Old Pole Text
        if old_pole_text: 
            cmd_text_2 = f"-TEXT {text_x:.3f},{y_lama:.3f} 0 {old_pole_text}"
            ws.cell(row=row, column=26).value = cmd_text_2
            script_commands.append(cmd_text_2)
        
        # 6. Combined Details
        if details_text:
            cmd_details = f"-TEXT {text_x:.3f},{y_details:.3f} 0 {details_text}"
            ws.cell(row=row, column=27).value = cmd_details
            script_commands.append(cmd_details)

    script_commands.append("OSMODE 15359")
    script_commands.append("") 

    wb.save(output_excel)
    with open(output_scr, "w") as scr_file:
        scr_file.write("\n".join(script_commands))

    print(f"Beres bro! Suffix huruf di jalur kabel udah fix.")

# --- EXECUTION ---
process_and_route("data_input.xlsx", "data_output.xlsx", "plot_otomatis.scr")