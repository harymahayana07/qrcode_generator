import openpyxl
import barcode
from barcode import Code128
import os

# path file excel
file_path = 'C:/xampp/htdocs/cv2test/kantor/generate_code.xlsx'
# nama sheet yang berisi data
sheet_name = 'nama_sheet'
# kolom yang berisi data barcode
barcode_column = 'B'
# baris pertama yang berisi data (judul kolom)
start_row = 2
# output directory
output_dir = 'C:/xampp/htdocs/cv2test/kantor/'

# membaca file excel
workbook = openpyxl.load_workbook(file_path)
sheet = workbook[sheet_name]

# looping untuk setiap baris pada kolom barcode
for row in sheet.iter_rows(min_row=start_row, min_col=2, values_only=True):
    barcode_value = row[0]
    # generate barcode
    barcode128 = Code128(barcode_value)
    # simpan barcode sebagai file PNG
    file_name = f'{barcode_value}.png'
    file_path = os.path.join(output_dir, file_name)
    barcode128.save(file_path)
    print(
        f'Barcode dengan nilai {barcode_value} berhasil dibuat dan disimpan di {file_path}')
