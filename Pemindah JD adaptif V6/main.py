import os
import glob
import textwrap
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.cell.cell import MergedCell

def detect_template_and_grading(source_path):
    file_name = os.path.basename(source_path).lower()
    tpl_dir = "TemplateJobDescBaruKosongan"
    
    # 1. KATEGORI MANAGERIAL
    if any(kw in file_name for kw in ["sub section head", "sub sec head", "sub sec. head"]):
        return os.path.join(tpl_dir, "Template GOL 123_Managerial.xlsx")
    if any(kw in file_name for kw in ["section head", "sec. head", "sec head"]):
        return os.path.join(tpl_dir, "Template GOL 4_Managerial.xlsx")
    if any(kw in file_name for kw in ["sub. dept head", "sub dept head", "sub department head"]):
        return os.path.join(tpl_dir, "Template GOL 5_Managerial.xlsx")

    # 2. KATEGORI SPESIALIS
    if any(kw in file_name for kw in ["pelaksana", "operator", "senior pelaksana", "staff engineer", "staff officer"]):
        return os.path.join(tpl_dir, "Template GOL 123_Spesialis.xlsx")
    if any(kw in file_name for kw in ["assistant engineer", "assistant officer", "ast. officer", "ast. engineer", "ast officer", "ast engineer"]):
        return os.path.join(tpl_dir, "Template GOL 4_Spesialis.xlsx")
    if any(kw in file_name for kw in ["engineer", "officer", "senior officer", "senior engineer"]):
        return os.path.join(tpl_dir, "Template GOL 5_Spesialis.xlsx")

    return None


def extract_specific_tasks(source_sheet):
    """
    Ekstraksi Tugas Pokok Spesifik (SUDAH PERFECT - JANGAN DIOTAK-ATIK)
    """
    items = []
    current_item = None
    is_in_spesifik = False
    
    start_col = 22  # Kolom V
    end_col = 41    # Kolom AO

    for row in range(1, source_sheet.max_row + 1):
        cell_v = source_sheet.cell(row=row, column=22).value 
        cell_w = source_sheet.cell(row=row, column=23).value 
        
        row_vals = []
        for col in range(start_col, end_col + 1):
            try:
                cell = source_sheet.cell(row=row, column=col)
                if isinstance(cell, MergedCell):
                    continue
                if cell.value is not None:
                    v = str(cell.value).strip()
                    if v and v not in row_vals:
                        row_vals.append(v)
            except Exception:
                continue
                    
        line_str = " ".join(row_vals).strip()
        if not line_str:
            continue

        upper_line = line_str.upper()

        if not is_in_spesifik:
            if "SPESIFIK" in upper_line and len(line_str) < 30:
                is_in_spesifik = True
                continue

        if is_in_spesifik:
            if any(upper_line.startswith(kw) for kw in ["PENDIDIKAN", "KETERAMPILAN", "WEWENANG", "HUBUNGAN KERJA", "PERSYARATAN"]):
                break
            if "PENDIDIKAN" in upper_line or "KETERAMPILAN" in upper_line:
                break
                
            v_val = str(cell_v).replace('.', '').strip() if cell_v is not None else ""
            w_val = str(cell_w).replace('.', '').strip().lower() if cell_w is not None else ""

            full_row_text = " ".join(row_vals).strip()

            if v_val.isdigit():
                col_no = int(v_val)
                text_cols = []
                for c in range(24, end_col + 1):
                    val_c = source_sheet.cell(row=row, column=c).value
                    if val_c is not None:
                        text_cols.append(str(val_c).strip())
                text_combined = " ".join(text_cols).strip()
                if not text_combined:
                    text_combined = " ".join(row_vals[1:]).strip()

                current_item = {"no": f"{col_no}.", "text": text_combined, "subitems": []}
                items.append(current_item)
                
                if re.search(r'\b[a-z]\.\s', text_combined, re.IGNORECASE):
                    parts = re.split(r'\s+(?=[a-z]\.\s)', text_combined, flags=re.IGNORECASE)
                    current_item["text"] = parts[0].strip()
                    for p in parts[1:]:
                        sub_match = re.match(r'^([a-z])\.\s*(.*)', p.strip(), re.IGNORECASE)
                        if sub_match:
                            s_char, s_txt = sub_match.groups()
                            current_item["subitems"].append({
                                "no": f"{s_char.lower()}.",
                                "text": s_txt.strip()
                            })

            elif w_val in ['a', 'b', 'c', 'd', 'e', 'f', 'g'] and current_item:
                sub_cols = []
                for c in range(24, end_col + 1):
                    val_c = source_sheet.cell(row=row, column=c).value
                    if val_c is not None:
                        sub_cols.append(str(val_c).strip())
                sub_text = " ".join(sub_cols).strip()
                if not sub_text:
                    sub_text = " ".join(row_vals[1:]).strip()

                if re.search(r'\b[a-z]\.\s', sub_text, re.IGNORECASE):
                    parts = re.split(r'\s+(?=[a-z]\.\s)', sub_text, flags=re.IGNORECASE)
                    for p in parts:
                        sub_match = re.match(r'^([a-z])\.\s*(.*)', p.strip(), re.IGNORECASE)
                        if sub_match:
                            s_char, s_txt = sub_match.groups()
                            current_item["subitems"].append({
                                "no": f"{s_char.lower()}.",
                                "text": s_txt.strip()
                            })
                else:
                    current_item["subitems"].append({
                        "no": f"{w_val}.",
                        "text": sub_text
                    })

            elif current_item:
                if re.search(r'\b[a-z]\.\s', full_row_text, re.IGNORECASE):
                    parts = re.split(r'\s+(?=[a-z]\.\s)', full_row_text, flags=re.IGNORECASE)
                    for p in parts:
                        sub_match = re.match(r'^([a-z])\.\s*(.*)', p.strip(), re.IGNORECASE)
                        if sub_match:
                            s_char, s_txt = sub_match.groups()
                            current_item["subitems"].append({
                                "no": f"{s_char.lower()}.",
                                "text": s_txt.strip()
                            })
                        else:
                            if current_item["subitems"]:
                                current_item["subitems"][-1]["text"] += " " + p.strip()
                            else:
                                current_item["text"] += " " + p.strip()
                else:
                    if current_item["subitems"]:
                        current_item["subitems"][-1]["text"] += " " + full_row_text
                    else:
                        current_item["text"] += " " + full_row_text

    return items


def extract_internal_connection(source_sheet):
    """
    Ekstraksi Internal dari Job Desc Lama:
    - Mengecek Kolom B: Jika ada tanda '-', abaikan/ignore tanda tersebut.
    - Mengambil konten teks murni yang berada di Kolom C.
    """
    internal_items = []
    is_in_internal = False
    
    for row in range(1, source_sheet.max_row + 1):
        try:
            # Ambil nilai dari kolom-kolom penting
            col_b_val = str(source_sheet.cell(row=row, column=2).value or "").strip() # Kolom B (biasanya tempat '-')
            col_c_val = str(source_sheet.cell(row=row, column=3).value or "").strip() # Kolom C (tempat teks)
            
            # Gabungkan teks baris untuk deteksi header
            full_row_text = " ".join([str(source_sheet.cell(row=row, column=col).value or "") for col in range(1, 10)]).upper()
        except Exception:
            continue

        # 1. Deteksi awal bab Internal
        if not is_in_internal:
            if "INTERNAL" in full_row_text or ("KETERKAITAN" in full_row_text and "PIHAK LAIN" in full_row_text):
                is_in_internal = True
                continue

        # 2. Jika sedang di dalam section Internal
        if is_in_internal:
            # STOP jika sudah mencapai bab lain (Eksternal, Tugas Pokok, dll)
            stop_keywords = ["EKSTERNAL", "B. EKSTERNAL", "WEWENANG", "TANGGUNG JAWAB", "TUGAS POKOK", "SPESIFIK"]
            if any(kw in full_row_text for kw in stop_keywords) and not "MENJALIN" in full_row_text:
                break

            # Ambil paragraf pengantar (misal: "Menjalin informasi...")
            if any(k in full_row_text for k in ["MENJALIN", "MENYANGKUT", "KOMUNIKASI", "INFORMASI"]):
                if full_row_text.strip() not in internal_items:
                    # Ambil teks utuh dari baris pengantar
                    intro_text = " ".join([str(source_sheet.cell(row=row, column=c).value) for c in range(2, 10) if source_sheet.cell(row=row, column=c).value is not None]).strip()
                    if intro_text and intro_text not in internal_items:
                        internal_items.append(intro_text)
                continue

            # ATURAN UTAMA: Cek Kolom B. Jika ada tanda '-', ignore Kolom B, ambil teks di Kolom C!
            if col_b_val in ["-", "–", "•", "*"]:
                if col_c_val and col_c_val != "-" and col_c_val not in internal_items:
                    internal_items.append(col_c_val)

    return internal_items

def extract_wewenang(source_sheet):
    """
    Ekstraksi Wewenang (SUDAH PERFECT - JANGAN DIOTAK-ATIK)
    """
    items = []
    current_item = None
    is_in_wewenang = False
    
    start_col = 41  
    end_col = 63

    for row in range(1, source_sheet.max_row + 1):
        row_vals = []
        for col in range(start_col, end_col + 1):
            try:
                cell = source_sheet.cell(row=row, column=col)
                if isinstance(cell, MergedCell):
                    continue
                if cell.value is not None:
                    v = str(cell.value).strip()
                    if v and v not in row_vals:
                        row_vals.append(v)
            except Exception:
                continue
                    
        line_str = " ".join(row_vals).strip()
        if not line_str:
            continue

        upper_line = line_str.upper()
        if not is_in_wewenang:
            if "WEWENANG" in upper_line and len(line_str) < 30:
                is_in_wewenang = True
                continue

        if is_in_wewenang:
            stop_keywords = [
                "TANGGUNG JAWAB", "TANGGUNGJAWAB", "KETERKAITAN", 
                "PERSYARATAN", "HUBUNGAN KERJA", "PENDIDIKAN",
                "PENGALAMAN", "SIKAP KERJA", "SIKAP", "DIVISI", 
                "DEPARTEMEN", "BAGIAN", "SUB DEPT"
            ]
            if any(upper_line.startswith(kw) for kw in stop_keywords):
                break
                
            first_word = row_vals[0].replace('.', '').strip() if row_vals else ""
            
            if first_word.isdigit():
                col_no = int(first_word)
                text_combined = " ".join(row_vals[1:]).strip()
                
                if current_item:
                    items.append(current_item)
                current_item = {"no": col_no, "text": text_combined}
            elif current_item:
                if not any(k in upper_line for k in ["DIVISI :", "DEPARTEMEN :", "BAGIAN :"]):
                    current_item["text"] += " " + line_str

    if current_item:
        items.append(current_item)
        
    return items


def insert_rows_and_preserve_layout(ws, row_idx, amount):
    if amount <= 0:
        return

    old_heights = {}
    for r in range(row_idx, ws.max_row + 1):
        if r in ws.row_dimensions and ws.row_dimensions[r].height is not None:
            old_heights[r] = ws.row_dimensions[r].height

    merged_ranges = list(ws.merged_cells.ranges)
    ranges_to_shift = []
    for rng in merged_ranges:
        if rng.min_row >= row_idx:
            ranges_to_shift.append(rng)
            ws.unmerge_cells(range_string=str(rng))

    ws.insert_rows(row_idx, amount)

    for rng in ranges_to_shift:
        ws.merge_cells(
            start_row=rng.min_row + amount,
            start_column=rng.min_col,
            end_row=rng.max_row + amount,
            end_column=rng.max_col
        )

    for r in range(row_idx, row_idx + amount):
        ws.row_dimensions[r].height = 16.5

    for old_r, h in old_heights.items():
        ws.row_dimensions[old_r + amount].height = h


def generate_new_filename(source_filename):
    base_name = os.path.splitext(source_filename)[0]
    match = re.match(r"^([a-zA-Z]+)(\d+)(.*)", base_name)
    if match:
        prefix_alpha, code_num, rest = match.groups()
        return f"HR{code_num}{rest}.xlsx"
    else:
        return f"HR_{source_filename}"


def process_single_file(source_path, output_dir):
    try:
        template_path = detect_template_and_grading(source_path)
        
        if not template_path or not os.path.exists(template_path):
            print(f"Error, job grading not found: {os.path.basename(source_path)}")
            return

        print(f"🎯 [INFO] {os.path.basename(source_path)} menggunakan template: {os.path.basename(template_path)}")

        wb_src = openpyxl.load_workbook(source_path, data_only=True)
        sheet_src = wb_src.active
        
        wb_tpl = openpyxl.load_workbook(template_path)
        sheet_tpl = wb_tpl.active
        
        data_spesifik = extract_specific_tasks(sheet_src)
        data_internal = extract_internal_connection(sheet_src)
        data_wewenang = extract_wewenang(sheet_src)
        
        # MAIN_CHARS_PER_LINE = 130 dikunci sesuai permintaan
        MAIN_CHARS_PER_LINE = 130
        SUB_CHARS_PER_LINE = 125

        # ==========================================
        # 1. PROSES TUGAS POKOK SPESIFIK
        # ==========================================
        target_row_spesifik = None
        for r in range(1, sheet_tpl.max_row + 1):
            for c in range(1, sheet_tpl.max_column + 1):
                val = sheet_tpl.cell(row=r, column=c).value
                if val and "TUGASPOKOKSPECIFICMOVED" in str(val).upper():
                    target_row_spesifik = r
                    break
            if target_row_spesifik:
                break

        if not target_row_spesifik:
            print(f"❌ [GAGAL] Marker 'TugasPokokSpecificMoved' tidak ditemukan di template {os.path.basename(template_path)}")
            return

        flattened_rows = []
        for item in data_spesifik:
            main_lines = textwrap.wrap(item["text"], width=MAIN_CHARS_PER_LINE, break_long_words=True, break_on_hyphens=False) or [""]
            for l_idx, line in enumerate(main_lines):
                flattened_rows.append({
                    "no": item["no"] if l_idx == 0 else "",
                    "text": line,
                    "is_sub": False
                })
            
            for sub in item["subitems"]:
                sub_lines = textwrap.wrap(sub["text"], width=SUB_CHARS_PER_LINE, break_long_words=True, break_on_hyphens=False) or [""]
                for l_idx, line in enumerate(sub_lines):
                    flattened_rows.append({
                        "no": "",
                        "sub_no": sub["no"] if l_idx == 0 else "",
                        "text": line,
                        "is_sub": True
                    })

        total_needed_rows = len(flattened_rows)
        max_default_rows = 9  
        
        extra_rows = 0
        if total_needed_rows > max_default_rows:
            extra_rows = total_needed_rows - max_default_rows
        
        total_insert_needed = extra_rows + 1 
        if total_insert_needed > 0:
            insert_rows_and_preserve_layout(sheet_tpl, row_idx=target_row_spesifik + max_default_rows, amount=total_insert_needed)

        font_arial_10 = Font(name="Arial", size=10, bold=False)
        double_side = Side(border_style="double", color="000000")
        
        last_item_row = target_row_spesifik + total_needed_rows - 1
        spacer_row = last_item_row + 1  

        for idx, row_data in enumerate(flattened_rows, start=1):
            curr_row = target_row_spesifik + idx - 1
            sheet_tpl.row_dimensions[curr_row].height = 16.5

            if not row_data["is_sub"]:
                cell_no = sheet_tpl.cell(row=curr_row, column=2)
                cell_no.value = row_data["no"]
                cell_no.font = font_arial_10
                cell_no.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                
                cell_text = sheet_tpl.cell(row=curr_row, column=3)
                cell_text.value = row_data["text"]
                cell_text.font = font_arial_10
                cell_text.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            else:
                cell_sub_no = sheet_tpl.cell(row=curr_row, column=3)
                cell_sub_no.value = row_data["sub_no"]
                cell_sub_no.font = font_arial_10
                cell_sub_no.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)

                cell_sub_text = sheet_tpl.cell(row=curr_row, column=4)
                cell_sub_text.value = row_data["text"]
                cell_sub_text.font = font_arial_10
                cell_sub_text.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

            sheet_tpl.cell(row=curr_row, column=2).border = Border(left=double_side, top=None, bottom=None)
            sheet_tpl.cell(row=curr_row, column=41).border = Border(right=double_side, top=None, bottom=None)
            sheet_tpl.cell(row=curr_row, column=42).border = Border(left=double_side, top=None, bottom=None)
            sheet_tpl.cell(row=curr_row, column=63).border = Border(right=double_side, top=None, bottom=None)

        sheet_tpl.row_dimensions[spacer_row].height = 16.5
        sheet_tpl.cell(row=spacer_row, column=2).value = ""
        sheet_tpl.cell(row=spacer_row, column=3).value = ""
        sheet_tpl.cell(row=spacer_row, column=2).border = Border(left=double_side, top=None, bottom=None)
        sheet_tpl.cell(row=spacer_row, column=41).border = Border(right=double_side, top=None, bottom=None)
        sheet_tpl.cell(row=spacer_row, column=42).border = Border(left=double_side, top=None, bottom=None)
        sheet_tpl.cell(row=spacer_row, column=63).border = Border(right=double_side, top=None, bottom=None)

        for c in range(2, 64):
            cell = sheet_tpl.cell(row=spacer_row, column=c)
            curr_b = cell.border
            sheet_tpl.cell(row=spacer_row, column=c).border = Border(
                left=curr_b.left if curr_b else None,
                right=curr_b.right if curr_b else None,
                top=None,
                bottom=double_side
            )

# ==========================================
        # 2. PROSES INTERNAL CONNECTION (DIPERBAIKI)
        # ==========================================
        target_row_internal = None
        target_col_internal = None
        
        for r in range(1, sheet_tpl.max_row + 1):
            for c in range(1, sheet_tpl.max_column + 1):
                val = sheet_tpl.cell(row=r, column=c).value
                if val and "INTERNALCONNECTION" in str(val).upper():
                    target_row_internal = r
                    target_col_internal = c
                    break
            if target_row_internal:
                break

        if target_row_internal:
            INTERNAL_CHARS_PER_LINE = 55
            flattened_internal = []
            
            bullet_counter = 1
            for text_item in data_internal:
                lines = textwrap.wrap(
                    text_item, 
                    width=INTERNAL_CHARS_PER_LINE, 
                    break_long_words=True, 
                    break_on_hyphens=False
                ) or [""]

                # Cek apakah ini teks pengantar (misal mengandung kata menyangkut/menjalin) atau poin bernomor
                is_intro = "menjalin" in text_item.lower() or "menyangkut" in text_item.lower() or "komunikasi" in text_item.lower()
                
                for l_idx, line in enumerate(lines):
                    if is_intro:
                        flattened_internal.append({
                            "no": "",
                            "text": line
                        })
                    elif text_item == "-":
                        flattened_internal.append({
                            "no": "",
                            "text": "-"
                        })
                    else:
                        flattened_internal.append({
                            "no": f"{bullet_counter}." if l_idx == 0 else "",
                            "text": line
                        })
                if not is_intro and text_item != "-":
                    bullet_counter += 1

            # Jika baris internal memerlukan insert row tambahan
            if len(flattened_internal) > 3:
                insert_rows_and_preserve_layout(sheet_tpl, row_idx=target_row_internal + 3, amount=len(flattened_internal) - 3)

            for idx, row_data in enumerate(flattened_internal):
                curr_row = target_row_internal + idx
                sheet_tpl.row_dimensions[curr_row].height = 16.5

                cell_no = sheet_tpl.cell(row=curr_row, column=target_col_internal)
                cell_no.value = row_data["no"]
                cell_no.font = font_arial_10
                cell_no.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                
                cell_text = sheet_tpl.cell(row=curr_row, column=target_col_internal + 1)
                cell_text.value = row_data["text"]
                cell_text.font = font_arial_10
                cell_text.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

        # ==========================================
        # 3. PROSES WEWENANG (TETAP AMAN TIDAK DIUBAH)
        # ==========================================
        if data_wewenang:
            target_row_wewenang = None
            target_col_wewenang = None
            
            for r in range(1, sheet_tpl.max_row + 1):
                for c in range(1, sheet_tpl.max_column + 1):
                    val = sheet_tpl.cell(row=r, column=c).value
                    if val and "KEWENANGANMOVED" in str(val).upper():
                        target_row_wewenang = r
                        target_col_wewenang = c
                        break
                if target_row_wewenang:
                    break

            if target_row_wewenang:
                WEWENANG_CHARS_PER_LINE = 55 
                
                justified_wewenang = []
                for idx, item in enumerate(data_wewenang, start=1):
                    full_text = item["text"]
                    lines = textwrap.wrap(
                        full_text, 
                        width=WEWENANG_CHARS_PER_LINE, 
                        break_long_words=True, 
                        break_on_hyphens=False
                    ) or [""]

                    for l_idx, line in enumerate(lines):
                        justified_wewenang.append({
                            "no": f"{idx}." if l_idx == 0 else "",
                            "text": line
                        })

                for idx, row_data in enumerate(justified_wewenang):
                    curr_row = target_row_wewenang + idx
                    sheet_tpl.row_dimensions[curr_row].height = 16.5

                    cell_no = sheet_tpl.cell(row=curr_row, column=target_col_wewenang)
                    cell_no.value = row_data["no"]
                    cell_no.font = font_arial_10
                    cell_no.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                    
                    cell_text = sheet_tpl.cell(row=curr_row, column=target_col_wewenang + 1)
                    cell_text.value = row_data["text"]
                    cell_text.font = font_arial_10
                    cell_text.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

        out_name = generate_new_filename(os.path.basename(source_path))
        output_path = os.path.join(output_dir, out_name)
        wb_tpl.save(output_path)
        print(f"✅ [BERHASIL] {os.path.basename(source_path)} -> {out_name}")

    except Exception as e:
        print(f"❌ [GAGAL] Error memproses {os.path.basename(source_path)}: {str(e)}")


def batch_process():
    input_dir = "input_jobdesc_lama"
    output_dir = "output_jobdesc_baru"
    
    os.makedirs(output_dir, exist_ok=True)
    excel_files = glob.glob(os.path.join(input_dir, "*.xlsx"))
    
    for file_path in excel_files:
        if os.path.basename(file_path).startswith("~$"):
            continue
        process_single_file(file_path, output_dir)


if __name__ == "__main__":
    batch_process()