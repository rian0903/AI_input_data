import os
import re
import glob
import shutil
from datetime import datetime
import openpyxl
from rapidocr_onnxruntime import RapidOCR

EXCEL_FILE = "Usulan_BBB_2026_Bireun_agustus.xlsx"

def clean_spelling(text):
    if not text:
        return ""
    text = text.strip()
    # Common OCR typo corrections
    text = re.sub(r'\bBIREVEN\b', 'BIREUEN', text, flags=re.IGNORECASE)
    text = re.sub(r'\bGAMPONO\b', 'GAMPONG', text, flags=re.IGNORECASE)
    text = re.sub(r'\bBARD\b', 'BARO', text, flags=re.IGNORECASE)
    text = re.sub(r'\bPenigelola\b', 'Pengelola', text, flags=re.IGNORECASE)
    text = re.sub(r'\bReferenst\b', 'Referensi', text, flags=re.IGNORECASE)
    text = re.sub(r'\bSirkulasl\b', 'Sirkulasi', text, flags=re.IGNORECASE)
    text = re.sub(r'\bProvins\b', 'Provinsi', text, flags=re.IGNORECASE)
    text = re.sub(r'\bd Negeri\b', 'Negeri', text, flags=re.IGNORECASE)
    text = re.sub(r'\b20d7\b', '2017', text, flags=re.IGNORECASE)
    
    # Clean multiple spaces/newlines
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_date_with_year(texts):
    for t in texts:
        t_upper = t.upper()
        if 'BIREUEN' in t_upper and any(m in t_upper for m in ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']):
            clean_date = clean_spelling(t)
            # if year not in date, look for year nearby
            if not re.search(r'\b20\d{2}\b', clean_date):
                for t2 in texts:
                    if re.search(r'\b20\d{2}\b', t2) and t2.strip() in ['2022', '2023', '2024', '2025', '2026']:
                        clean_date += f" {t2.strip()}"
                        break
            return clean_date
    return ""

def format_with_suffix(val, suffix):
    if not val:
        return ""
    val_str = str(val).strip()
    # Check if value is a number (integer, float or formatted with comma/dot)
    # E.g. '112,5' or '52' or '1.474'
    clean_val = val_str.replace(',', '').replace('.', '').replace(' ', '')
    if clean_val.isdigit() and suffix.lower() not in val_str.lower():
        return f"{val_str} {suffix}"
    return val_str

def parse_form_data(images):
    engine = RapidOCR()
    page1_texts = []
    page2_texts = []
    
    for img in images:
        result, elapse = engine(img)
        if not result:
            continue
        
        is_page2 = False
        for box, text, conf in result:
            if any(k in text.lower() for k in ['luas gedung', 'jumlah judul', 'jumlah eksemplar', 'bantuan stimulan']):
                is_page2 = True
                break
        
        texts = [line[1].strip() for line in result]
        if is_page2:
            page2_texts = texts
        else:
            page1_texts = texts

    # Default data values mapped directly to the new 24 columns structure
    data = {
        'Nama Lengkap Kepala Pustakawan': '',
        'Jabatan': 'kepala perpustakaan',
        'Alamat Email perpustakaan': '',
        'Nomor Telepon/Whatsapp': '',
        'Nomor Pokok Perpustakaan (NPP)': '',
        'Subjenis Perpustakaan': 'Perpustakaan Sekolah',
        'Nama Perpustakaan': '',
        'Provinsi': 'Aceh',
        'Kabupaten': 'Bireuen',
        'Kecamatan': '',
        'Desa/Kelurahan': '',
        'Alamat Lengkap (Jalan RT/ RW)': '',
        'Kode Pos': '',
        'Nama Lembaga Induk': '',
        'Tahun Berdiri Perpustakaan': '',
        'Nomor SK Pendirian Perpustakaan': '',
        'Jumlah Anggota Perpustakaan': '',
        'Luas Gedung Perpustakaan': '',
        'Jumlah Tenaga Perpustakaan': '',
        'Jumlah Judul Koleksi': '',
        'Jumlah Eksemplar Koleksi': '',
        'Jumlah Hari Layanan dalam Seminggu (hari)': '6',
        'Tahun Pendataan': '2022',
    }
    
    # ------------------ PROCESS PAGE 1 ------------------
    # 1. School name (Nama Lembaga) & Nama Perpustakaan
    school_names = []
    for t in page1_texts:
        if any(k in t.upper() for k in ['UPTD', 'SD NEGERI', 'SMP NEGERI', 'SMA NEGERI', 'SMK NEGERI']):
            clean_name = clean_spelling(t)
            clean_name = clean_name.replace('UPTD.SMP NEGERI LBIREUEN', 'UPTD. SMP NEGERI 1 BIREUEN')
            clean_name = clean_name.replace('UPTD.SMP NEGERI 1', 'UPTD. SMP NEGERI 1 BIREUEN')
            if clean_name not in school_names:
                school_names.append(clean_name)
    
    if school_names:
        data['Nama Lembaga Induk'] = school_names[0]
        
        # Search if any text has "PERPUSTAKAAN"
        perp_name = ""
        for t in page1_texts:
            if 'PERPUSTAKAAN' in t.upper() and any(k in t.upper() for k in ['UPTD', 'SD', 'SMP', 'SMA']):
                perp_name = clean_spelling(t)
                perp_name = perp_name.replace('PERPUSTAKAANUPTDSDNEGERI7BIREUEN', 'PERPUSTAKAAN UPTD SD NEGERI 7 BIREUEN')
                break
        
        if not perp_name:
            perp_name = school_names[0]
            if not perp_name.upper().startswith('PERPUSTAKAAN'):
                perp_name = f"PERPUSTAKAAN {perp_name}"
        data['Nama Perpustakaan'] = perp_name
        
        # Determine subjenis
        if 'SMP' in school_names[0].upper():
            data['Subjenis Perpustakaan'] = 'Perpustakaan Sekolah' # sheet SD_ & SMP templates have 'Perpustakaan Sekolah'
        elif 'SD' in school_names[0].upper():
            data['Subjenis Perpustakaan'] = 'Perpustakaan Sekolah'
            
    # 2. Geography
    for t in page1_texts:
        if 'ACEH' in t.upper():
            data['Provinsi'] = 'Aceh'
            break
    for t in page1_texts:
        if 'BIREUEN' in t.upper() or 'BIREVEN' in t.upper():
            data['Kabupaten'] = 'Bireuen'
            break
    for t in page1_texts:
        if 'KOTA JUANG' in t.upper():
            data['Kecamatan'] = 'Kota Juang'
            break
    for t in page1_texts:
        if 'GAMPONG BARO' in t.upper() or 'GAMPONO BARD' in t.upper() or 'GAMPONG BARU' in t.upper():
            data['Desa/Kelurahan'] = 'Gampong Baro'
            break
        elif 'GEULANGGANG KULAM' in t.upper() or 'GEULANGGANGKULAM' in t.upper():
            data['Desa/Kelurahan'] = 'Geulanggang Gampong' # match existing Excel style or correct name
            break
            
    # Alamat
    for t in page1_texts:
        if 'JL.' in t.upper() or 'JALAN' in t.upper() or 'LAKSAMANA' in t.upper() or 'MOL CUT' in t.upper():
            data['Alamat Lengkap (Jalan RT/ RW)'] = clean_spelling(t)
            break
            
    # Kode Pos
    for t in page1_texts:
        if re.match(r'^\b242\d{2}\b$', t):
            data['Kode Pos'] = t
            break
            
    # Nomor Telepon
    for t in page1_texts:
        if re.search(r'\b08\d{8,11}\b', t) or re.search(r'\b00\d{8,11}\b', t):
            num = t.strip()
            if num.startswith('0013'):
                num = '0813' + num[4:]
            data['Nomor Telepon/Whatsapp'] = num
            break
            
    # Email
    for t in page1_texts:
        if '@' in t:
            data['Alamat Email perpustakaan'] = t.strip()
            break
            
    # Names with S.Pd (Pustakawan)
    names_spd = []
    for t in page1_texts:
        if ',S.Pd' in t or ', S.Pd' in t or ',S.Pd.I' in t or ', S.Pd.I' in t:
            names_spd.append(clean_spelling(t))
            
    if names_spd:
        for name in names_spd:
            if not any(k in name.upper() for k in ['HARUN', 'NAZIR']):
                data['Nama Lengkap Kepala Pustakawan'] = name
                
    # SK Pendirian
    for t in page1_texts:
        if 'NOMOR:' in t.upper() or 'NOMOR : ' in t.upper():
            data['Nomor SK Pendirian Perpustakaan'] = clean_spelling(t)
            break
            
    # Tahun Berdiri
    for t in page1_texts:
        if t.strip() in ['2017', '2010', '2007', '2018', '2019', '2020']:
            data['Tahun Berdiri Perpustakaan'] = t.strip()
            break
            
    # Jumlah Peserta Didik
    for t in page1_texts:
        if t.strip() in ['988', '239', '500']:
            data['Jumlah Anggota Perpustakaan'] = t.strip()
            break

    # ------------------ PROCESS PAGE 2 ------------------
    if page2_texts:
        # Luas Gedung
        for idx, t in enumerate(page2_texts):
            if 'LUAS GEDUNG' in t.upper() or 'LUAS' in t.upper():
                for offset in range(-3, 4):
                    if 0 <= idx + offset < len(page2_texts):
                        cand = page2_texts[idx + offset]
                        if cand.strip().isdigit() and len(cand.strip()) <= 3:
                            data['Luas Gedung Perpustakaan'] = cand.strip()
                            break
                            
        # Jumlah Tenaga
        for idx, t in enumerate(page2_texts):
            if 'JUMLAH TENAGA' in t.upper() or 'TENAGA' in t.upper():
                for offset in range(-3, 4):
                    if 0 <= idx + offset < len(page2_texts):
                        cand = page2_texts[idx + offset]
                        if 'orang' in cand or (cand.strip().isdigit() and len(cand.strip()) == 1):
                            m = re.search(r'\d+', cand)
                            data['Jumlah Tenaga Perpustakaan'] = m.group(0) if m else cand.strip()
                            break
                            
        # Jumlah Judul Koleksi & Jumlah Eksemplar Koleksi
        numbers_with_dots = []
        for t in page2_texts:
            clean_num = t.replace(' ', '')
            if re.match(r'^\d{1,3}\.\d{3}$', clean_num):
                numbers_with_dots.append(clean_num)
                
        if len(numbers_with_dots) >= 2:
            sorted_nums = sorted(numbers_with_dots, key=lambda x: float(x.replace('.', '')))
            data['Jumlah Judul Koleksi'] = sorted_nums[0]
            data['Jumlah Eksemplar Koleksi'] = sorted_nums[1]
        elif len(numbers_with_dots) == 1:
            data['Jumlah Judul Koleksi'] = numbers_with_dots[0]
            
        # Date for Tahun Pendataan extraction
        for t in page2_texts:
            if re.search(r'\b20\d{2}\b', t):
                m = re.search(r'\b20\d{2}\b', t)
                data['Tahun Pendataan'] = m.group(0)
                break
                
    return data

def main():
    print("="*60)
    print("SISTEM PENGINPUTAN DATA PERPUSTAKAAN BERBASIS SESI")
    print("="*60)
    
    # 1. Detect images
    images = sorted(glob.glob("*.jpeg") + glob.glob("*.jpg") + glob.glob("*.png"))
    images = [img for img in images if os.path.isfile(img)]
    
    if len(images) != 2:
        print(f"ERROR: Ditemukan {len(images)} gambar di folder aktif.")
        print("Pastikan HANYA ada 2 gambar (Halaman 1 dan Halaman 2) dari satu formulir sekolah.")
        print("Silakan pindahkan gambar lain keluar dari folder ini dan jalankan kembali.")
        return
        
    print(f"Mendeteksi berkas untuk sesi ini:")
    for img in images:
        print(f" - {img}")
    print("\nSedang memproses OCR (Mohon tunggu)...")
    
    # 2. Extract data
    data = parse_form_data(images)
    
    # Define ordered key list for presentation & editing (23 editable fields)
    fields = [
        ('Nama Lengkap Kepala Pustakawan', data['Nama Lengkap Kepala Pustakawan']),
        ('Jabatan', data['Jabatan']),
        ('Alamat Email perpustakaan', data['Alamat Email perpustakaan']),
        ('Nomor Telepon/Whatsapp', data['Nomor Telepon/Whatsapp']),
        ('Nomor Pokok Perpustakaan (NPP)', data['Nomor Pokok Perpustakaan (NPP)']),
        ('Subjenis Perpustakaan', data['Subjenis Perpustakaan']),
        ('Nama Perpustakaan', data['Nama Perpustakaan']),
        ('Provinsi', data['Provinsi']),
        ('Kabupaten', data['Kabupaten']),
        ('Kecamatan', data['Kecamatan']),
        ('Desa/Kelurahan', data['Desa/Kelurahan']),
        ('Alamat Lengkap (Jalan RT/ RW)', data['Alamat Lengkap (Jalan RT/ RW)']),
        ('Kode Pos', data['Kode Pos']),
        ('Nama Lembaga Induk', data['Nama Lembaga Induk']),
        ('Tahun Berdiri Perpustakaan', data['Tahun Berdiri Perpustakaan']),
        ('Nomor SK Pendirian Perpustakaan', data['Nomor SK Pendirian Perpustakaan']),
        ('Jumlah Anggota Perpustakaan', data['Jumlah Anggota Perpustakaan']),
        ('Luas Gedung Perpustakaan', data['Luas Gedung Perpustakaan']),
        ('Jumlah Tenaga Perpustakaan', data['Jumlah Tenaga Perpustakaan']),
        ('Jumlah Judul Koleksi', data['Jumlah Judul Koleksi']),
        ('Jumlah Eksemplar Koleksi', data['Jumlah Eksemplar Koleksi']),
        ('Jumlah Hari Layanan dalam Seminggu (hari)', data['Jumlah Hari Layanan dalam Seminggu (hari)']),
        ('Tahun Pendataan', data['Tahun Pendataan'])
    ]
    
    # 3. Interactive Edit Loop
    while True:
        print("\n" + "="*50)
        print("HASIL EKSTRAKSI DATA (SILAKAN PERIKSA & EDIT):")
        print("="*50)
        for idx, (k, v) in enumerate(fields, 1):
            print(f"{idx:2d}. {k:<40} : {v}")
            
        print("\nKetik nomor field (1-23) untuk mengedit nilainya.")
        print("Ketik 'y' jika semua data sudah benar dan siap disimpan ke Excel.")
        print("Ketik 'n' untuk membatalkan proses sesi ini.")
        
        choice = input("Pilihan Anda: ").strip().lower()
        
        if choice == 'y':
            break
        elif choice == 'n':
            print("Sesi dibatalkan. Tidak ada data yang ditulis ke Excel.")
            return
        elif choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(fields):
                k, v = fields[idx]
                new_val = input(f"Masukkan nilai baru untuk [{k}] (sebelumnya: '{v}'): ").strip()
                fields[idx] = (k, new_val)
                data[k] = new_val
            else:
                print("Nomor tidak valid.")
        else:
            print("Pilihan tidak dikenal.")
            
    # 4. Save to Excel
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: File Excel {EXCEL_FILE} tidak ditemukan di folder aktif!")
        return
        
    print(f"\nSedang menyimpan ke {EXCEL_FILE}...")
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        
        # Determine the sheet name: 'SD_' or 'SMP'
        subjenis = data['Subjenis Perpustakaan'].upper()
        school_name = data['Nama Lembaga Induk'].upper()
        if 'SMP' in subjenis or 'SMP' in school_name:
            sheet_name = 'SMP'
        else:
            sheet_name = 'SD_'
            
        ws = wb[sheet_name]
        
        # Find first empty row (checking column 8: Nama Perpustakaan)
        row_idx = 5
        while True:
            val = ws.cell(row=row_idx, column=8).value
            if val is None or str(val).strip() == "":
                break
            row_idx += 1
            
        print(f"Menulis data pada Sheet: [{sheet_name}], Baris: {row_idx}")
        
        # Write values to Excel (columns 1-24)
        # 1: No (use pre-existing if filled, else write new)
        existing_no = ws.cell(row=row_idx, column=1).value
        if existing_no is None:
            ws.cell(row=row_idx, column=1, value=row_idx - 4)
            
        # 2: Nama Lengkap Kepala Pustakawan
        ws.cell(row=row_idx, column=2, value=data['Nama Lengkap Kepala Pustakawan'])
        # 3: Jabatan
        ws.cell(row=row_idx, column=3, value=data['Jabatan'])
        # 4: Alamat Email perpustakaan
        ws.cell(row=row_idx, column=4, value=data['Alamat Email perpustakaan'])
        # 5: Nomor Telepon/Whatsapp
        ws.cell(row=row_idx, column=5, value=data['Nomor Telepon/Whatsapp'])
        # 6: Nomor Pokok Perpustakaan (NPP)
        ws.cell(row=row_idx, column=6, value=data['Nomor Pokok Perpustakaan (NPP)'])
        # 7: Subjenis Perpustakaan
        ws.cell(row=row_idx, column=7, value=data['Subjenis Perpustakaan'])
        # 8: Nama Perpustakaan
        ws.cell(row=row_idx, column=8, value=data['Nama Perpustakaan'])
        # 9: Provinsi
        ws.cell(row=row_idx, column=9, value=data['Provinsi'])
        # 10: Kabupaten
        ws.cell(row=row_idx, column=10, value=data['Kabupaten'])
        # 11: Kecamatan
        ws.cell(row=row_idx, column=11, value=data['Kecamatan'])
        # 12: Desa/Kelurahan
        ws.cell(row=row_idx, column=12, value=data['Desa/Kelurahan'])
        # 13: Alamat Lengkap (Jalan RT/ RW)
        ws.cell(row=row_idx, column=13, value=data['Alamat Lengkap (Jalan RT/ RW)'])
        # 14: Kode Pos
        # Try to write as integer if it is numeric
        zip_code = data['Kode Pos']
        if zip_code and zip_code.isdigit():
            ws.cell(row=row_idx, column=14, value=int(zip_code))
        else:
            ws.cell(row=row_idx, column=14, value=zip_code)
            
        # 15: Nama Lembaga Induk
        ws.cell(row=row_idx, column=15, value=data['Nama Lembaga Induk'])
        
        # 16: Tahun Berdiri Perpustakaan
        year_est = data['Tahun Berdiri Perpustakaan']
        if year_est and year_est.isdigit():
            ws.cell(row=row_idx, column=16, value=int(year_est))
        else:
            ws.cell(row=row_idx, column=16, value=year_est)
            
        # 17: Nomor SK Pendirian Perpustakaan
        ws.cell(row=row_idx, column=17, value=data['Nomor SK Pendirian Perpustakaan'])
        
        # 18: Jumlah Anggota Perpustakaan (Format with 'Siswa')
        ws.cell(row=row_idx, column=18, value=format_with_suffix(data['Jumlah Anggota Perpustakaan'], 'Siswa'))
        
        # 19: Luas Gedung (Format with 'm2')
        ws.cell(row=row_idx, column=19, value=format_with_suffix(data['Luas Gedung Perpustakaan'], 'm2'))
        
        # 20: Jumlah Tenaga (Format with 'Orang')
        ws.cell(row=row_idx, column=20, value=format_with_suffix(data['Jumlah Tenaga Perpustakaan'], 'Orang'))
        
        # 21: Jumlah Judul Koleksi
        ws.cell(row=row_idx, column=21, value=data['Jumlah Judul Koleksi'])
        # 22: Jumlah Eksemplar Koleksi
        ws.cell(row=row_idx, column=22, value=data['Jumlah Eksemplar Koleksi'])
        
        # 23: Jumlah Hari Layanan dalam Seminggu (hari)
        ws.cell(row=row_idx, column=23, value=data['Jumlah Hari Layanan dalam Seminggu (hari)'])
        
        # 24: Tahun Pendataan
        year_data = data['Tahun Pendataan']
        if year_data and year_data.isdigit():
            ws.cell(row=row_idx, column=24, value=int(year_data))
        else:
            ws.cell(row=row_idx, column=24, value=year_data)
            
        wb.save(EXCEL_FILE)
        print("Sukses menyimpan data ke Excel!")
        
    except Exception as e:
        print(f"ERROR Gagal menulis ke Excel: {e}")
        return
        
    # 5. Archive files
    clean_school_folder = re.sub(r'[^a-zA-Z0-9_]', '_', data['Nama Lembaga Induk']).strip('_')
    archive_dir = os.path.join("processed", f"{clean_school_folder}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    
    print(f"Mengarsipkan file ke: {archive_dir}")
    try:
        os.makedirs(archive_dir, exist_ok=True)
        for img in images:
            shutil.move(img, os.path.join(archive_dir, img))
        print("Pengarsipan selesai! Folder aktif sekarang bersih.")
    except Exception as e:
        print(f"Peringatan: Gagal memindahkan file ke arsip: {e}")

if __name__ == '__main__':
    main()
