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
    
    start_col = 21  # Kolom V
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


import re

def extract_internal_connection(source_sheet):
    """
    Ekstraksi Keterkaitan Pihak Lain (Internal):
    - Mengenali poin dengan angka numerik (1., 2.) maupun simbol minus (-).
    - Mengabaikan baris pengantar 'Menjalin informasi... menyangkut :'.
    - Berhenti total (break) jika menyentuh bab/komponen lain (Eksternal, Wewenang, Job Spec, dll).
    - Menghasilkan daftar teks bersih siap dinomori di template baru.
    """
    internal_items = []
    is_in_internal = False
    
    # Keyword untuk menghentikan ekstraksi agar komponen lain tidak ikut terseret
    stop_keywords = [
        "B. EKSTERNAL", "EKSTERNAL", "WEWENANG", "BATASAN WEWENANG", 
        "TANGGUNG JAWAB", "TANGGUNGJAWAB", "JOB SPECIFICATION", 
        "PERSYARATAN JABATAN", "PENDIDIKAN", "PENGALAMAN"
    ]

    # Regex untuk mendeteksi apakah baris merupakan poin baru (angka atau simbol -/*)
    bullet_or_num_pattern = re.compile(r'^\s*(?:[\-\•\*\–]|\(?\d+[\.\)]?)\s*')

    for row in range(1, source_sheet.max_row + 1):
        try:
            row_texts = []
            col_b_val = str(source_sheet.cell(row=row, column=2).value or "").strip()
            col_c_val = str(source_sheet.cell(row=row, column=3).value or "").strip()

            for col in range(2, 19):  # Membaca Kolom B s.d. R
                val = source_sheet.cell(row=row, column=col).value
                if val is not None:
                    v_str = str(val).strip()
                    if v_str:
                        row_texts.append(v_str)

            full_row_text = " ".join(row_texts).strip()
            upper_row_text = full_row_text.upper()
        except Exception:
            continue

        if not upper_row_text:
            continue

        # 1. DETEKSI PINTU MASUK SECTION INTERNAL
        if not is_in_internal:
            if "A. INTERNAL" in upper_row_text or upper_row_text == "INTERNAL":
                is_in_internal = True
                continue

        # 2. PENGAMAN STOP MUTLAK (Mencegah bab/komponen lain ikut tersedot)
        if is_in_internal:
            # Berhenti jika menemukan judul section/bab penghenti
            if any(kw == upper_row_text or upper_row_text.startswith(kw) for kw in stop_keywords):
                break
                
            # Berhenti jika mendeteksi penomoran 90+ dari section lain
            match_num = re.match(r'^\s*(\d+)[\.\)]?', col_b_val)
            if match_num and int(match_num.group(1)) >= 90:
                break

            # 3. ABAIKAN KALIMAT PENGANTAR
            if "MENJALIN INFORMASI" in upper_row_text or "MENYANGKUT" in upper_row_text:
                continue

            # 4. PENENTUAN TEKS TARGET
            target_text = col_c_val if col_c_val else col_b_val
            if not target_text and row_texts:
                target_text = full_row_text

            if not target_text or target_text == "-":
                continue

            # Cek apakah baris ini poin baru (pakai angka / '-') atau sambungan baris atasnya
            is_new_point = bool(bullet_or_num_pattern.match(col_b_val)) or \
                           bool(bullet_or_num_pattern.match(target_text))

            # Bersihkan angka/simbol minus (-) dari depan teks
            cleaned_text = bullet_or_num_pattern.sub('', target_text).strip()
            while cleaned_text.startswith('-') or cleaned_text.startswith('–'):
                cleaned_text = cleaned_text.lstrip('-–').strip()

            if cleaned_text:
                if is_new_point:
                    internal_items.append(cleaned_text)
                else:
                    # Jika baris lanjutan tanpa nomor/strip (seperti 'wfefwef...'), gabungkan ke poin sebelumnya
                    if internal_items:
                        internal_items[-1] += f" {cleaned_text}"
                    else:
                        internal_items.append(cleaned_text)

    return internal_items


import re

def extract_external_xyz_connection(source_sheet):
    """
    Ekstraksi Hubungan Eksternal XYZ:
    - Menghapus string pengantar deskriptif tertentu.
    - Menghentikan proses (BREAK) seketika jika menemukan 'JOB DESCRIPTION' atau kata kunci stop lainnya.
    - Akurat memisahkan antara baris sambungan (multiline text) dan poin baru (bullet/numerik).
    - Berhenti total jika nomor urut melebihi 90+.
    """
    external_xyz_items = []
    is_in_ext_xyz = False
    
    # Keyword pemicu awal bab eksternal
    trigger_keywords = [
        "BERHUBUNGAN DENGAN SELURUH DEPARTEMEN",
        "BERHUBUNGAN DENGAN DEPARTEMEN USER",
        "BERHUBUNGAN DENGAN DEPARTEMEN",

    ]
    
    # Frasa pengantar spesifik yang ingin dihapus total sebelum ekstraksi poin
    unwanted_phrases = [
        "batasan wewenangnya atau atas persetujuan atasan, perihal:",
        "batasan wewenangnya atau atas persetujuan atasan",
        "sesuai dengan batasan wewenangnya",
        "perihal:"
    ]
    
    # Kata kunci penghentian mutlak
    stop_keywords = [
        "JOB DESCRIPTION", "JOB SPECIFICATION", "POSISI JABATAN", 
        "WEWENANG", "TANGGUNG JAWAB", "TANGGUNGJAWAB", "PERSYARATAN", 
        "HUBUNGAN KERJA", "PENDIDIKAN", "TRAINING", 
        "PENGALAMAN", "SIKAP KERJA", "TUGAS POKOK"
    ]

    # Regex untuk mendeteksi bullet/simbol ATAU angka numerik (1., 1), (1), dst)
    pattern_bullet_or_num = r'^\s*(?:[\-\•\*\–]|\(?\d+[\.\)]?\s*\-?)'

    for row in range(1, source_sheet.max_row + 1):
        try:
            row_texts = []
            col_b_val = ""
            col_c_val = ""
            
            # Scan Kolom B (2) sampai R (18)
            for col in range(2, 19):
                val = source_sheet.cell(row=row, column=col).value
                if val is not None:
                    v_str = str(val).strip()
                    if v_str:
                        row_texts.append(v_str)
                        if col == 2:
                            col_b_val = v_str
                        elif col == 3:
                            col_c_val = v_str
            
            full_row_text = " ".join(row_texts).strip()
            
            # Pembersihan frasa pengantar yang tidak diinginkan
            cleaned_full_row = full_row_text
            for phrase in unwanted_phrases:
                cleaned_full_row = re.sub(re.escape(phrase), '', cleaned_full_row, flags=re.IGNORECASE).strip()

            upper_row_text = cleaned_full_row.upper()
        except Exception:
            continue

        if not upper_row_text:
            continue

        # ----------------------------------------------------
        # 1. PENGAMAN STOP MUTLAK (PRIORITAS UTAMA SEBELUM PROSES)
        # ----------------------------------------------------
        if is_in_ext_xyz:
            # Cek apakah ada kata kunci stop seperti JOB DESCRIPTION
            if any(kw in upper_row_text for kw in stop_keywords):
                break

            # Cek penomoran angka 90+ di Kolom B maupun di teks baris
            match_num_b = re.match(r'^\s*(\d+)[\.\)]?', col_b_val)
            match_num_full = re.match(r'^\s*(\d+)[\.\)]?', full_row_text)
            
            if (match_num_b and int(match_num_b.group(1)) >= 90) or \
               (match_num_full and int(match_num_full.group(1)) >= 90):
                break

        # 2. AKTIFKAN MODE EKSTRAKSI
        if not is_in_ext_xyz:
            full_original_upper = full_row_text.upper()
            if any(kw in full_original_upper for kw in trigger_keywords):
                is_in_ext_xyz = True
                
                # Cek apakah baris pemicu ini sudah langsung berisi bullet/angka
                is_bullet_in_row = bool(re.search(pattern_bullet_or_num, col_b_val)) or \
                                   bool(re.search(pattern_bullet_or_num, col_c_val))
                if not is_bullet_in_row:
                    continue  # Skip baris pengantar

        # 3. PROSES EKSTRAKSI POIN
        if is_in_ext_xyz:
            # Cek keberadaan bullet/numerik
            has_bullet_or_num = bool(re.search(pattern_bullet_or_num, col_b_val)) or \
                                bool(re.search(pattern_bullet_or_num, col_c_val)) or \
                                bool(re.search(pattern_bullet_or_num, full_row_text))

            # Jika TANPA bullet/numerik (Sambungan Teks / Multiline):
            if not has_bullet_or_num:
                if external_xyz_items and not any(kw in upper_row_text for kw in trigger_keywords):
                    target_text = col_c_val if col_c_val else full_row_text
                    for phrase in unwanted_phrases:
                        target_text = re.sub(re.escape(phrase), '', target_text, flags=re.IGNORECASE).strip()
                    if target_text:
                        external_xyz_items[-1] += " " + target_text
                continue

            # Jika ADA bullet/numerik (Poin Baru):
            target_text = col_c_val if col_c_val else col_b_val
            if not target_text or target_text in ["-", "–"]:
                target_text = full_row_text

            if not target_text or target_text in ["-", "–"]:
                continue

            # Hapus frasa pengantar jika ikut terbawa
            for phrase in unwanted_phrases:
                target_text = re.sub(re.escape(phrase), '', target_text, flags=re.IGNORECASE).strip()

            # Bersihkan simbol bullet maupun angka numerik lama di awal teks
            cleaned = re.sub(r'^\s*(?:\(?\d+[\.\)]?\s*\-?|[\-\•\*\–])\s*', '', target_text).strip()

            # Pastikan teks yang dibersihkan tidak mengandung kata kunci stop secara tidak sengaja
            if cleaned and not any(kw in cleaned.upper() for kw in stop_keywords):
                if cleaned not in external_xyz_items:
                    external_xyz_items.append(cleaned)

    return external_xyz_items

import re
import openpyxl

def extract_external_connection(source_sheet):
    """
    Ekstraksi Hubungan Eksternal Utama:
    - Membaca poin angka numerik (1., 2., 3.) maupun tanda strip (- / • / *).
    - Membedakan secara akurat antara poin utama baru dan baris lanjutan (multiline text).
    - BERHENTI TOTAL (BREAK) jika menyentuh 'Berhubungan dengan departemen user' atau bab lain.
    """
    external_items = []
    is_in_ext = False
    
    # Regex pola penomoran numerik berurutan maupun simbol strip/bullet
    pattern_bullet_or_num = r'^\s*(?:\(?\d+[\.\)]?\s*\-?|[\-\•\*\–])'

    # Kata kunci penghentian mutlak (TAMBAHAN: DEPARTEMEN USER)
    stop_keywords = [
        "DEPARTEMEN USER", "BERHUBUNGAN DENGAN DEPARTEMEN USER",
        "WEWENANG", "KEWENANGAN", "TUGAS POKOK SPESIFIK", "SPESIFIK", 
        "TUGAS POKOK", "TANGGUNG JAWAB", "TANGGUNGJAWAB", "PENDIDIKAN", 
        "TRAINING", "PERSYARATAN", "HUBUNGAN KERJA", "KETERKAITAN", 
        "SIKAP KERJA", "PENGALAMAN", "A. POSISI JABATAN", "JOB DESCRIPTION", "JOB SPECIFICATION"
    ]

    for row in range(1, source_sheet.max_row + 1):
        try:
            row_cells_vals = []
            col_b_val = ""
            col_c_val = ""

            # Scan dari Kolom B (2) sampai R (18)
            for col in range(2, 19):
                cell = source_sheet.cell(row=row, column=col)
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                if cell.value is not None:
                    v = str(cell.value).strip()
                    if v:
                        row_cells_vals.append((col, v))
                        if col == 2:
                            col_b_val = v
                        elif col == 3:
                            col_c_val = v
            
            if not row_cells_vals:
                continue

            full_line_str = " ".join([v for _, v in row_cells_vals]).strip()
            upper_line = full_line_str.upper()
        except Exception:
            continue

        # 1. PENGAMAN STOP MUTLAK: Jika sudah di area eksternal lalu mendeteksi DEPARTEMEN USER atau bab lain, STOP!
        if is_in_ext:
            if any(kw in upper_line for kw in stop_keywords):
                break

            # Stop jika menemukan penomoran bab besar (misal 90+)
            match_num = re.match(r'^\s*(\d+)[\.\)]?', col_b_val)
            if match_num and int(match_num.group(1)) >= 90:
                break

        # 2. PINTU MASUK: Deteksi Blok Eksternal Utama
        if not is_in_ext:
            # Pastikan bukan baris 'DEPARTEMEN USER'
            if "DEPARTEMEN USER" not in upper_line:
                if "EKSTERNAL" in upper_line and any(kw in upper_line for kw in ["MELIPUTI", "PERIHAL", "BERHUBUNGAN", "B. EKSTERNAL"]):
                    is_in_ext = True
                    
                    # Cek apakah baris header pemicu ini sudah langsung memuat nomor/bullet/strip
                    has_bullet = bool(re.search(pattern_bullet_or_num, col_b_val)) or bool(re.search(pattern_bullet_or_num, full_line_str))
                    if not has_bullet:
                        continue  # Skip baris header / pengantar awal

        # 3. PROSES EKSTRAKSI POIN
        if is_in_ext:
            # Skip jika baris ini hanyalah teks pengantar tanpa nomor/strip/bullet
            if any(k in upper_line for k in ["BERHUBUNGAN", "EKSTERNAL PERUSAHAAN", "BATASAN PERIHAL"]) and not re.search(pattern_bullet_or_num, full_line_str):
                continue

            # Periksa apakah baris ini diawali ANGKA NUMERIK ATAU TANDA STRIP/BULLET
            has_num_or_bullet = bool(re.search(pattern_bullet_or_num, col_b_val)) or \
                                bool(re.search(pattern_bullet_or_num, col_c_val)) or \
                                bool(re.search(pattern_bullet_or_num, full_line_str))

            # --- A. JIKA BARIS ADALAH SAMBUNGAN KALIMAT (TIDAK MEMILIKI NOMOR / STRIP BARU) ---
            if not has_num_or_bullet:
                if external_items:
                    text_to_append = col_c_val if col_c_val else full_line_str
                    text_to_append = re.sub(r'^\s*[\-\•\*\–]\s*', '', text_to_append).strip()
                    
                    if text_to_append and text_to_append != "-":
                        # Gabungkan ke poin aktif sebelumnya
                        external_items[-1] += " " + text_to_append
                continue

            # --- B. JIKA BARIS ADALAH POIN BARU (DIAWALI ANGKA 1., 2. ATAU STRIP -) ---
            target_text = col_c_val if col_c_val else full_line_str
            
            # Bersihkan angka/strip/bullet lama di awal teks
            cleaned_text = re.sub(r'^\s*(?:\(?\d+[\.\)]?\s*\-?|[\-\•\*\–])\s*', '', target_text).strip()

            if cleaned_text and cleaned_text != "-" and cleaned_text not in external_items:
                external_items.append(cleaned_text)

    return external_items

def get_department_code(department_name):
    dictionary = {
        "Human Capital": "HR55001",
        "Quality Management System": "HR55002",
        "QMS": "HR55002",
        "Areso": "HR55003",
        "Bahan": "HR55004",
        "COR": "HR55005",
        "EG": "HR55006",
        "Elegant Gold": "HR55006",
        "EGV": "HR55007",
        "Elegant Gold Variasi": "HR55007",
        "Finishing Central": "HR55008",
        "FC": "HR55008",
        "Hollow": "HR55009",
        "Kalung": "HR55010",
        "RnD": "HR55011",
        "R&D": "HR55011",
        "Research & Development": "HR55011",
        "Variasi": "HR55012",
        "Marketing Ekspor": "HR55013",
        "Marketing Lokal": "HR55014",
        "ICT": "HR55015",
        "Information Communication Technology": "HR55015",
        "Maintenance": "HR55016",
        "Procurement & Logistic": "HR55017",
        "PPIC": "HR55018",
        "Keamanan Dalam": "HR55019",
        "Kamdal": "HR55019",
        "Chemical": "HR55020",
        "Strategic Management Officer": "HR55021",
        "SMO": "HR55021",
        "Finance": "HR55023",
        "Tax": "HR55024",
        "Workshop": "HR55025",
    }
    
    # Normalisasi teks masukan menjadi huruf besar dan hilangkan spasi ekstra
    dept_upper = str(department_name).upper().strip()
    
    # Normalisasi keys pada dictionary agar pencocokan case-insensitive / fleksibel berhasil
    normalized_dict = {k.upper().strip(): v for k, v in dictionary.items()}
    
    if dept_upper in normalized_dict:
        return normalized_dict[dept_upper]
    else:
        return "000000"


def process_document_code(target_sheet):
    department_name = "QMS"  # Default fallback
    txt_path = "DepartmentNameTitle.txt"
    
    # Membaca nama departemen langsung dari file DepartmentNameTitle.txt
    if os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            dept_name_content = f.read().strip()
            if dept_name_content:
                department_name = dept_name_content

    code = get_department_code(department_name)

    # Mengganti placeholder CodeDocument di dalam sheet template baru
    for row in range(1, target_sheet.max_row + 1):
        for col in range(1, target_sheet.max_column + 1):
            cell = target_sheet.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str):
                if "(CodeDocument)" in cell.value or "CodeDocument" in cell.value:
                    cell.value = cell.value.replace("(CodeDocument)", code).replace("CodeDocument", code)

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


def process_department_name(target_sheet):
    txt_path = "DepartmentNameTitle.txt"
    if not os.path.exists(txt_path):
        return

    with open(txt_path, "r", encoding="utf-8") as f:
        dept_name_content = f.read().strip()  # Pakai .strip() agar bersih

    if not dept_name_content:
        return

    for row in range(1, target_sheet.max_row + 1):
        for col in range(1, target_sheet.max_column + 1):
            cell = target_sheet.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str):
                # Ubah pencarian dari "InputDepartmentName" menjadi "DepartmentName"
                if "DepartmentName" in cell.value:
                    cell.value = cell.value.replace("DepartmentName", dept_name_content)


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


def process_job_title_from_filename(file_path, target_sheet):
    filename = os.path.basename(file_path)
    filename_without_ext = os.path.splitext(filename)[0]
    job_title = re.sub(r'^[Qq]\d+[\s\-_]*', '', filename_without_ext).strip()
    job_title_caps = job_title.upper()
    
    target_cell = target_sheet['X9']
    for merged_range in target_sheet.merged_cells.ranges:
        if 'X9' in merged_range:
            top_left_coord = merged_range.start_cell.coordinate
            target_cell = target_sheet[top_left_coord]
            break
            
    target_cell.value = job_title_caps
    current_font = target_cell.font
    if current_font:
        target_cell.font = Font(
            name=current_font.name,
            size=current_font.size,
            bold=True,
            italic=current_font.italic,
            color=current_font.color
        )
    else:
        target_cell.font = Font(bold=True)
        
    return job_title_caps


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

        print(f" [INFO] {os.path.basename(source_path)} menggunakan template: {os.path.basename(template_path)}")

        wb_src = openpyxl.load_workbook(source_path, data_only=True)
        sheet_src = wb_src.active
        
        wb_tpl = openpyxl.load_workbook(template_path)
        sheet_tpl = wb_tpl.active
        
        # Ekstraksi Data Sumber
        data_spesifik = extract_specific_tasks(sheet_src)
        data_internal = extract_internal_connection(sheet_src)
        data_ext_xyz = extract_external_xyz_connection(sheet_src) # [FITUR BARU] Ekstraksi Eksternal XYZ
        data_wewenang = extract_wewenang(sheet_src)

        # Proses Header & Atribut Dokumen
        job_title_inserted = process_job_title_from_filename(source_path, sheet_tpl)
        print(f" [POSITION] Berhasil mengisi PositionName dengan: '{job_title_inserted}'")
        
        process_document_code(sheet_tpl)
        process_department_name(sheet_tpl)

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
            print(f" [GAGAL] Marker 'TugasPokokSpecificMoved' tidak ditemukan di template {os.path.basename(template_path)}")
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
        thin_side = Side(style='thin', color='000000')
        
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
            sheet_tpl.cell(row=curr_row, column=43).border = Border(left=double_side, top=None, bottom=None)
            sheet_tpl.cell(row=curr_row, column=63).border = Border(right=double_side, top=None, bottom=None)

        sheet_tpl.row_dimensions[spacer_row].height = 16.5
        sheet_tpl.cell(row=spacer_row, column=2).value = ""
        sheet_tpl.cell(row=spacer_row, column=3).value = ""
        sheet_tpl.cell(row=spacer_row, column=2).border = Border(left=double_side, top=None, bottom=None)
        sheet_tpl.cell(row=spacer_row, column=41).border = Border(right=double_side, top=None, bottom=None)
        sheet_tpl.cell(row=spacer_row, column=43).border = Border(left=double_side, top=None, bottom=None)
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
        # 2. PROSES INTERNAL CONNECTION
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
                cleaned_text = re.sub(r'^[\-\•\*\–]\s*', '', str(text_item)).strip()
                while cleaned_text.startswith('-') or cleaned_text.startswith('–'):
                    cleaned_text = cleaned_text.lstrip('-–').strip()

                if not cleaned_text:
                    continue

                lines = textwrap.wrap(
                    cleaned_text, 
                    width=INTERNAL_CHARS_PER_LINE, 
                    break_long_words=True, 
                    break_on_hyphens=False
                ) or [""]
                
                for l_idx, line in enumerate(lines):
                    flattened_internal.append({
                        "no": f"{bullet_counter}." if l_idx == 0 else "",
                        "text": line
                    })
                bullet_counter += 1

            total_internal_rows = len(flattened_internal)

            if total_internal_rows > 4:
                insert_rows_and_preserve_layout(sheet_tpl, row_idx=target_row_internal + 3, amount=total_internal_rows - 3)

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

                # REVISI: Penambalan border menggunakan double_side (garis ganda) secara tegas pada Kolom 1 (A)
                cell_a = sheet_tpl.cell(row=curr_row, column=106)
                cell_a.border = Border(
                    left=cell_a.border.left if cell_a.border else None,
                    right=double_side,
                    top=cell_a.border.top if cell_a.border else None,
                    bottom=cell_a.border.bottom if cell_a.border else None
                )

                # 2. Penambalan border kanan di Kolom BK (Column "BK" / Index 63)
                cell_bk = sheet_tpl.cell(row=curr_row, column=106)  # BK = Kolom ke-63
                cell_bk.border = Border(
                    left=cell_bk.border.left if cell_bk.border else None,
                    right=double_side,
                    top=cell_bk.border.top if cell_bk.border else None,
                    bottom=cell_bk.border.bottom if cell_bk.border else None
                )



# ==========================================
        # PROSES EXTERNAL CONNECTION (DI-UPDATE & DIRAPIHKAN)
        # ==========================================
        data_external = extract_external_connection(sheet_src)
        
        target_row_ext = None
        target_col_ext = None
        
        for r in range(1, sheet_tpl.max_row + 1):
            for c in range(1, sheet_tpl.max_column + 1):
                val = sheet_tpl.cell(row=r, column=c).value
                if val and "EXTERNALCONNECTION" in str(val).upper():
                    target_row_ext = r
                    target_col_ext = c
                    break
            if target_row_ext:
                break

        if target_row_ext:
            EXT_CHARS_PER_LINE = 55
            flattened_ext = []
            
            if data_external:
                bullet_counter = 1
                for text_item in data_external:
                    lines = textwrap.wrap(
                        text_item, 
                        width=EXT_CHARS_PER_LINE, 
                        break_long_words=True, 
                        break_on_hyphens=False
                    ) or [""]
                    
                    for l_idx, line in enumerate(lines):
                        flattened_ext.append({
                            "no": f"{bullet_counter}." if l_idx == 0 else "",
                            "text": line
                        })
                    bullet_counter += 1
            else:
                # Fallback jika tidak terdeteksi: tulis 1 -, 2 -, 3 -
                fallback_items = ["-", "-", "-"]
                for idx, f_item in enumerate(fallback_items, start=1):
                    flattened_ext.append({
                        "no": f"{idx}.",
                        "text": f_item
                    })

            # Perapihan layout rows: Jika lebih dari 3 baris default template, insert baris otomatis
            if len(flattened_ext) > 100:
                insert_rows_and_preserve_layout(sheet_tpl, row_idx=target_row_ext + 3, amount=len(flattened_ext) - 3)

            for idx, row_data in enumerate(flattened_ext):
                curr_row = target_row_ext + idx
                sheet_tpl.row_dimensions[curr_row].height = 16.5

                cell_no = sheet_tpl.cell(row=curr_row, column=target_col_ext)
                cell_no.value = row_data["no"]
                cell_no.font = font_arial_10
                cell_no.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                
                cell_text = sheet_tpl.cell(row=curr_row, column=target_col_ext + 1)
                cell_text.value = row_data["text"]
                cell_text.font = font_arial_10
                cell_text.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

                # Tambal border kiri (kolom 1) secara presisi
                cell_a = sheet_tpl.cell(row=curr_row, column=1)
                cell_a.border = Border(
                    left=cell_a.border.left if cell_a.border else None,
                    right=thin_side,
                    top=cell_a.border.top if cell_a.border else None,
                    bottom=cell_a.border.bottom if cell_a.border else None
                )

        # ==========================================
        # 3. PROSES EXTERNAL XYZ CONNECTION [FITUR BARU]
        # ==========================================
        target_row_ext_xyz = None
        target_col_ext_xyz = None
        
        for r in range(1, sheet_tpl.max_row + 1):
            for c in range(1, sheet_tpl.max_column + 1):
                val = sheet_tpl.cell(row=r, column=c).value
                if val and "EXTERNALXYZCONNECTION" in str(val).upper():
                    target_row_ext_xyz = r
                    target_col_ext_xyz = c
                    break
            if target_row_ext_xyz:
                break

        if target_row_ext_xyz:
            EXT_XYZ_CHARS_PER_LINE = 55
            flattened_ext_xyz = []
            
            bullet_counter = 1
            for text_item in data_ext_xyz:
                cleaned_text = re.sub(r'^[\-\•\*\–]\s*', '', str(text_item)).strip()
                while cleaned_text.startswith('-') or cleaned_text.startswith('–'):
                    cleaned_text = cleaned_text.lstrip('-–').strip()

                if not cleaned_text:
                    continue

                lines = textwrap.wrap(
                    cleaned_text, 
                    width=EXT_XYZ_CHARS_PER_LINE, 
                    break_long_words=True, 
                    break_on_hyphens=False
                ) or [""]
                
                for l_idx, line in enumerate(lines):
                    flattened_ext_xyz.append({
                        "no": f"{bullet_counter}." if l_idx == 0 else "",
                        "text": line
                    })
                bullet_counter += 1

            if len(flattened_ext_xyz) > 100:
                insert_rows_and_preserve_layout(sheet_tpl, row_idx=target_row_ext_xyz + 3, amount=len(flattened_ext_xyz) - 3)

            for idx, row_data in enumerate(flattened_ext_xyz):
                curr_row = target_row_ext_xyz + idx
                sheet_tpl.row_dimensions[curr_row].height = 16.5

                cell_no = sheet_tpl.cell(row=curr_row, column=target_col_ext_xyz)
                cell_no.value = row_data["no"]
                cell_no.font = font_arial_10
                cell_no.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                
                cell_text = sheet_tpl.cell(row=curr_row, column=target_col_ext_xyz + 1)
                cell_text.value = row_data["text"]
                cell_text.font = font_arial_10
                cell_text.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

                # Tambal border kolom A (column=1) secara presisi
                cell_a = sheet_tpl.cell(row=curr_row, column=1)
                cell_a.border = Border(
                    left=cell_a.border.left if cell_a.border else None,
                    right=thin_side,
                    top=cell_a.border.top if cell_a.border else None,
                    bottom=cell_a.border.bottom if cell_a.border else None
                )
            
        # ==========================================
        # 4. PROSES WEWENANG
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
        print(f" [BERHASIL] {os.path.basename(source_path)} -> {out_name}")

    except Exception as e:
        print(f" [GAGAL] Error memproses {os.path.basename(source_path)}: {str(e)}")


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