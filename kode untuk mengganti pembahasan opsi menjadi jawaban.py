import pandas as pd
import re

# Fungsi untuk menghilangkan karakter kontrol
def sanitize_string(s):
    if isinstance(s, str):  # Periksa apakah s adalah string
        return re.sub(r'[\x00-\x1F]+', '', s)
    else:
        return s  # Jika s bukan string, kembalikan apa adanya

# Pemetaan huruf jawaban ke kolom
MAPPING = {
    "A": "a_opsi",
    "B": "b_opsi",
    "C": "c_opsi",
    "D": "d_opsi",
    "E": "e_opsi"
}

# Membaca file Excel
FILE_EXCEL = 'Export_SKB Guru MAPEL (Kelas).xlsx'  # Ganti dengan nama file Excel Anda
df = pd.read_excel(FILE_EXCEL, engine='openpyxl')

# Mengganti jawaban yang benar pada kolom P
for index, row in df.iterrows():
    # Pastikan bahwa pembahasan adalah string sebelum mencari di dalamnya
    pembahasan = str(row['explanation']) if pd.notnull(row['explanation']) else ''
    
    for huruf, kolom in MAPPING.items():
        jawaban_key = f"<b>Jawaban:{huruf}</b><br>"
        if jawaban_key in pembahasan:
            # Pastikan bahwa teks jawaban adalah string
            teks_jawaban = str(row[kolom]) if pd.notnull(row[kolom]) else ''
            sanitized_answer = sanitize_string(teks_jawaban)
            pembahasan_replaced = pembahasan.replace(jawaban_key, f"<b>Jawaban yang benar: {sanitized_answer}</b><br><br>")
            
            df.at[index, 'explanation'] = pembahasan_replaced
            break

# Menyimpan hasil ke file Excel baru
df.to_excel('123.xlsx', index=False, engine='openpyxl')
