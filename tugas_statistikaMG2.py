# TUGAS STATISTIKA ANALISIS DATA 
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

x = {
    "nama" : "Chris", "kelas" : "2025 TIG", "Nim": 25051204061
}
y = {
    "No": [1,2,3,4,5,6,7,8],
    "Wilayah (Kab/Kota)": [
        "Kab. Banjarnegara",
        "Kab. Wonogiri",
        "Kab. Sragen",
        "Kab. Rembang",
        "Kab. Ciamis",
        "Kab. Kuningan",
        "Kab. Situbondo",
        "Kab. Sampang"
    ],
    "Nilai Upah Minimum (Rp)": [
        2038005,
        2047500,
        2049000,
        2090000,
        2100000,
        2110000,
        2174000,
        2182000
    ],
    "Provinsi": [
        "Jawa Tengah",
        "Jawa Tengah",
        "Jawa Tengah",
        "Jawa Tengah",
        "Jawa Barat",
        "Jawa Barat",
        "Jawa Timur",
        "Jawa Timur"
    ],
    "Sumber Data": [
        "BPS / Pergub Jateng",
        "BPS / Pergub Jateng",
        "BPS / Pergub Jateng",
        "BPS / Pergub Jateng",
        "BPS / Pergub Jabar",
        "BPS / Pergub Jabar",
        "BPS / Pergub Jatim",
        "BPS / Pergub Jatim"
    ]
}

tugas = pd.DataFrame(y)

print("\nTUGAS STATISTIKA ANALIIS DATA\n")
print(f"nama :{x['nama']} \nkelas : {x['kelas']}\nnim : {x['Nim']}\n")
print("----------DATA UPAH MINIMUM KABUPATEN/KOTA----------\n")
print(tugas.to_string(index=False, justify="center"))

folder_script = os.path.dirname(os.path.abspath(__file__))
tugas2 = os.path.join(folder_script, "Chris 25051204061.xlsx")

tugas.to_excel(tugas2, index=False)

# STYLING

wb = load_workbook(tugas2)
ws = wb.active
# Warna
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Biru gelap
body_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")   # Biru muda
# Font
font_body = Font(name="Times New Roman")
font_header = Font(name="Times New Roman", bold=True, color="FFFFFF")
# Border hitam
thin_border = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000")
)
# Header Styling
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = font_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border
# Body Styling
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
        cell.fill = body_fill
        cell.font = font_body
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
# Format Rupiah Otomatis
for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
    for cell in row:
        cell.number_format = '"Rp"#,##0'
# Atur Lebar Kolom Otomatis
for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    ws.column_dimensions[column_letter].width = max_length + 2

wb.save(tugas2)

Penjelassan = """
ANALISIS DATA 
TIPE DATA  :

    1.Kuantitatif, data numerik yang dapat dihitung dan diukur 
    Contoh dalam tabel : Nilai Upah Minimun, No data

    2.Kualitatif , data yang menggmbarkan karakteristik yang tidak dapat diukur dengan angka
    Contoh dalam tabel : Wilayah, provinsi , sumber data

SKALA PENGUKURAN :

    1.Nominal : Hanya menggambarkan karakteristik tanpa ada urutan khusus
    Contoh dalam tabel : Wilayah,provonsi,sumber data

    2.Rasio : Memiliki Interval yang seragam dan memiliki titik 0 absolut
    Contoh dalam tabel :   Nilai upah minimum"""

print(Penjelassan)
os.startfile(tugas2)
