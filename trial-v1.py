import openpyxl
import re
import os
import warnings
from pyproj import Proj

# Matiin warning remeh dari openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def process_and_route(input_file, output_excel, output_folder):
    utm_proj = Proj(proj='utm', zone=47, south=True, ellps='WGS84')
    
    print(f"Reading {input_file}...")
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active
    
    pole_data = {}
    pole_list = []
    
    # --- SETUP HEADER OUTPUT ---
    ws.cell(row=1, column=22).value = "Input_Kabel"         
    ws.cell(row=1, column=23).value = "Input_tiang"         
    ws.cell(row=1, column=24).value = "Input_Umbang"        
    ws.cell(row=1, column=25).value = "Input_No Tiang"      
    ws.cell(row=1, column=26).value = "Input_No Tiang Lama" 
    ws.cell(row=1, column=27).value = "Input_details"       
    
    cable_headers = {}
    for col_idx in range(12, 22):
        val = ws.cell(row=1, column=col_idx).value
        cable_headers[col_idx] = str(val).strip() if val else f"Cable_{col_idx}"
        
    def get_cable_layer(cable_name):
        name_up = str(cable_name).upper()
        if "185" in name_up: return "CABLE 185"
        if "95" in name_up: return "CABLE 95"
        if "3 X 16" in name_up or "3X16" in name_up: return "CABLE 316"
        if "1X 16" in name_up or "1X16" in name_up: return "CABLE SERVICE"
        if "19064" in name_up or "9064" in name_up: return "PVC 9064"
        if "7083" in name_up or "7044" in name_up: return "PVC 7083"
        if "7173" in name_up or "3132" in name_up: return "BARE 7173"
        if "7122" in name_up: return "BARE 7122"
        return "CABLE SERVICE" 
    
    # --- PASS 1: Read and Clean Data ---
    for row in range(2, ws.max_row + 1):  
        coord_str = ws.cell(row=row, column=1).value
        main_pole_name = ws.cell(row=row, column=2).value
        
        if not coord_str or not main_pole_name:
            continue
            
        try:
            lat_str, lon_str = str(coord_str).split(',')
            lat = float(lat_str.strip())
            lon = float(lon_str.strip())
            x, y = utm_proj(lon, lat)
            
            pole_name_clean = str(main_pole_name).strip()
            
            old_pole_val = ws.cell(row=row, column=3).value
            old_pole_name = str(old_pole_val).strip() if old_pole_val else ""
            
            umbang_raw = ws.cell(row=row, column=4).value
            umbang_val = str(umbang_raw).strip() if umbang_raw is not None else ""
            umbang_block = f"U{umbang_val}" if umbang_val.isdigit() and 1 <= int(umbang_val) <= 9 else None
            
            col_e = str(ws.cell(row=row, column=5).value).strip() if ws.cell(row=row, column=5).value is not None else ""
            col_f = str(ws.cell(row=row, column=6).value).strip() if ws.cell(row=row, column=6).value is not None else ""
            col_g = str(ws.cell(row=row, column=7).value).strip() if ws.cell(row=row, column=7).value is not None else ""
            details_text = " ".join(filter(None, [col_e, col_f, col_g]))
            
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
            
            dominant_found = False
            cable_name_for_layer = "Cable 1X 16 Nmp" 
            
            for col_idx in range(12, 22):
                cable_val = ws.cell(row=row, column=col_idx).value
                if not dominant_found and cable_val is not None and str(cable_val).strip() != "":
                    dominant_found = True
                    cable_name_for_layer = cable_headers[col_idx]
                elif dominant_found:
                    ws.cell(row=row, column=col_idx).value = None
            
            if not dominant_found:
                ws.cell(row=row, column=15).value = 1 
            
            final_layer_kabel = get_cable_layer(cable_name_for_layer)
            
            pole_data[pole_name_clean] = {
                'x': x, 'y': y, 'row': row, 
                'block': block_name, 'umbang': umbang_block,
                'details': details_text, 'old_pole': old_pole_name,
                'cable_layer': final_layer_kabel
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
        match_suffix_child = re.search(r'^(.*)\s+/(\d+)([A-Za-z]+)$', pole)
        match_child = re.search(r'^(.*)\s+/(\d+)$', pole)
        match_main = re.search(r'^(.*?\s+)([A-Z]+)\s+(\d+)$', pole)
        
        if match_suffix_child:
            base = match_suffix_child.group(1).strip()
            num = int(match_suffix_child.group(2))
            suffix = match_suffix_child.group(3)
            if num != 1:
                target_parent = f"{base} /{num - 1}{suffix}"
                parent = target_parent if target_parent in pole_set else (get_joint_pole(target_parent) or target_parent)
            else:
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

    # --- PASS 3: Generate Split AutoCAD Commands with Auto-Grouping ---
    os.makedirs(output_folder, exist_ok=True)
    
    scripts_by_area = {}
    current_area = None
    
    for pole in pole_list: 
        data = pole_data[pole]
        
        if pole.upper().startswith("PE "):
            current_area = pole
        elif current_area is None:
            current_area = pole
            
        clean_area_name = re.sub(r'[\\/*?:"<>|]', "", current_area).strip()
        
        if clean_area_name not in scripts_by_area:
            # Pake setvar biar aman dari bug Enter di file gabungan
            scripts_by_area[clean_area_name] = [
                '(setvar "OSMODE" 0)',
                "(setq firstEnt (entlast))" 
            ]
            
        cmds = scripts_by_area[clean_area_name]
        
        row = data['row']
        x = data['x']
        y = data['y']
        block_name = data['block']
        umbang_block = data['umbang']
        details_text = data['details']
        old_pole_text = data['old_pole']
        cable_layer = data['cable_layer']
        
        text_x = x + 2
        y_pole = y + 1
        y_lama = y
        y_details = y - 1
        
        parent = pole_relations.get(pole)
        if parent and parent in pole_data:
            x_parent = pole_data[parent]['x']
            y_parent = pole_data[parent]['y']
            cmd_line = f"LINE {x_parent:.3f},{y_parent:.3f} {x:.3f},{y:.3f} "
            ws.cell(row=row, column=22).value = cmd_line
            cmds.append(f'(command "-layer" "m" "{cable_layer}" "")')
            cmds.append(cmd_line)
            
        cmd_insert = f'(command "-insert" "{block_name}" "{x:.3f},{y:.3f}" 1 1 0)'
        ws.cell(row=row, column=23).value = cmd_insert
        cmds.append('(command "-layer" "m" "0" "")')
        cmds.append(cmd_insert)

        if umbang_block:
            cmd_umbang = f'(command "-insert" "{umbang_block}" "{x:.3f},{y:.3f}" 1 1 0)'
            ws.cell(row=row, column=24).value = cmd_umbang
            cmds.append('(command "-layer" "m" "UMBANG" "")')
            cmds.append(cmd_umbang)
        
        if pole.upper().startswith("PE "):
            display_label_b = pole
        else:
            display_name_match = re.search(r'([A-Z]+\s+\d+.*)$', pole)
            display_label_b = display_name_match.group(1) if display_name_match else pole 
            
        if display_label_b: 
            cmd_text_1 = f"-TEXT {text_x:.3f},{y_pole:.3f} 0 {display_label_b}"
            ws.cell(row=row, column=25).value = cmd_text_1
            cmds.append('(command "-layer" "m" "POLE NUMBER" "")')
            cmds.append(cmd_text_1)
        
        if old_pole_text: 
            cmd_text_2 = f"-TEXT {text_x:.3f},{y_lama:.3f} 0 {old_pole_text}"
            ws.cell(row=row, column=26).value = cmd_text_2
            cmds.append('(command "-layer" "m" "TIANG LAMA" "")')
            cmds.append(cmd_text_2)
        
        if details_text:
            cmd_details = f"-TEXT {text_x:.3f},{y_details:.3f} 0 {details_text}"
            ws.cell(row=row, column=27).value = cmd_details
            cmds.append('(command "-layer" "m" "details" "")')
            cmds.append(cmd_details)

    master_script_lines = []
    
    # Finalisasi dan eksekusi
    for area_name, cmds in scripts_by_area.items():
        cmds.append('(setq ss (ssadd))')
        cmds.append('(setq en firstEnt)')
        cmds.append('(while (setq en (if en (entnext en) (entnext))) (ssadd en ss))')
        cmds.append('(if (> (sslength ss) 0) (command "-group" "c" "*" "Auto-Grouped by Script" ss ""))')
        
        cmds.append('(command "-layer" "s" "0" "")')
        cmds.append('(setvar "OSMODE" 15359)')
        
        file_path = os.path.join(output_folder, f"{area_name}.scr")
        with open(file_path, "w") as scr_file:
            # Amanin pakai enter di ujung file satuan
            scr_file.write("\n".join(cmds) + "\n")
            
        master_script_lines.extend(cmds)

    # File master dengan handling enter yang rapi
    master_path = os.path.join(output_folder, "00_RUN_ALL.scr")
    with open(master_path, "w") as master_file:
        master_file.write("\n".join(master_script_lines) + "\n")

    wb.save(output_excel)
    print(f"Beres bro! File gabungan 00_RUN_ALL.scr udah kebal error.")

# --- EXECUTION ---
process_and_route("data_input.xlsx", "data_output.xlsx", "output src")