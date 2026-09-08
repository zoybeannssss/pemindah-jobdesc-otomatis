def extract_wewenang(source_sheet):
    """
    Mengambil poin-poin Kewenangan dari template lama
    tanpa tergantung warna sel (bekerja pada sel 'no fill' maupun berwarna).
    """
    items = []
    current_item = None
    is_in_wewenang = False
    
    # Berdasarkan gambar, area Wewenang berada di kolom sebelah kanan (misal kolom AS ke BZ, atau sesuaikan range kolomnya)
    # Kita bisa memindai dari kolom 43 (AQ/AR) sampai kolom 63 (BK/BL) atau menyesuaikan lebar tabel wewenang.
    start_col = 43  # Sesuaikan dengan posisi kolom Wewenang di template lama
    end_col = 63

    for row in range(1, source_sheet.max_row + 1):
        row_vals = []
        for col in range(start_col, end_col + 1):
            cell = source_sheet.cell(row=row, column=col)
            if cell.value is not None:
                v = str(cell.value).strip()
                if v and v not in row_vals:
                    row_vals.append(v)
                    
        line_str = " ".join(row_vals).strip()
        if not line_str:
            continue

        # 1. DETEKSI MULAI (HEADER WEWENANG)
        if not is_in_wewenang:
            if "WEWENANG" in line_str.upper() and len(line_str) < 30:
                is_in_wewenang = True
                continue

        # 2. PROSES AREA WEWENANG
        if is_in_wewenang:
            stop_keywords = [
                "TANGGUNG JAWAB", "TANGGUNGJAWAB", "KETERKAITAN", 
                "PERSYARATAN", "HUBUNGAN KERJA", "PENDIDIKAN"
            ]
            if any(stop_kw in line_str.upper() for stop_kw in stop_keywords):
                break
                
            first_word = row_vals[0].replace('.', '').strip() if row_vals else ""
            
            if first_word.isdigit():
                col_no = int(first_word)
                text_combined = " ".join(row_vals[1:]).strip()
                
                if current_item:
                    items.append(current_item)
                current_item = {"no": col_no, "text": text_combined}
            elif current_item:
                current_item["text"] += " " + line_str

    if current_item:
        items.append(current_item)
        
    return items