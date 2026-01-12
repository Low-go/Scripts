import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl import Workbook, load_workbook
import os

# Select File window 
def select_file():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title= "Select Excel Files",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    root.destroy
    return file_path



def audit(file_path):

    try:
        wb = openpyxl.load_workbook(file_path)
        new_wb = Workbook()
        new_ws = new_wb.active
        ws = wb["App export"]

        new_row_counter = 1

        # This section just grabs the student Names and student ids of the Resident excel
        # And moves them to column A and B of the new one
        for row in ws.iter_rows(min_row=2):
            for cell in row:


                col_I = row[8] 
                col_B = row[1]
                col_C = row[2]
                col_N = row[13]
                col_Q = row[16]

                if col_I.value == "CURRENT":

                    if "Hale" in str(col_N.value):
                        address1 = "55-220 Kulanui St"
                    elif "TVA" in str(col_N.value):
                        address1 = "55-550 Naniloa Loop" 
                    else: 
                        address1 = col_N.value # fallbback

                    address2 = col_Q.value

                    # new_row = col_I.row
                    new_ws.cell(row=new_row_counter, column=1).value = col_B.value
                    new_ws.cell(row=new_row_counter, column=2).value = col_C.value
                    new_ws.cell(row=new_row_counter, column=3).value = address1
                    new_ws.cell(row=new_row_counter, column=4).value = address2
                    new_ws.cell(row=new_row_counter, column=5).value = "Laie"
                    new_ws.cell(row=new_row_counter, column=6).value = "Hawaii"

                    new_row_counter +=1

            base, ext = os.path.splitext(file_path)
            new_file_path = f"{base}_updated{ext}"
            new_wb.save(new_file_path)

    except Exception as e:
            print(f"Error processing {file_path}: {e}")


def main():

    file_path = select_file()
    pass


if __name__ == "__main__":
    main()
