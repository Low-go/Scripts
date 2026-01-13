import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl import Workbook, load_workbook
import os

# Select which automated task you wish to perform
def select_task():
    result = {"value": None}  

    def choose(val):
        result["value"] = val
        root.destroy()  

    root = tk.Tk()
    root.title("Select Job Automation")
    root.geometry("400x150")

    button1 = tk.Button(root, text="Combine housing with Main", command=lambda: choose(0))
    button1.pack(side="left", padx=20, pady=20)

    button2 = tk.Button(root, text="Format Housing Sheet", command=lambda: choose(1))
    button2.pack(side="right", padx=20, pady=20)

    root.mainloop()  
    return result["value"]

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
        ws = wb["App Export"]

        new_row_counter = 1

        # This section just grabs the student Names and student ids of the Resident excel
        # And moves them to column A and B of the new one
        for row in ws.iter_rows(min_row=2):

            col_I = row[8] 
            col_B = row[1]
            col_C = row[2]
            col_N = row[13]
            col_Q = row[16]

            if col_I.value == "CURRENT":

                if "Hale" in str(col_N.value):
                    address1 = "55-220 Kulanui St"
                    #remove leading zero if it exists
                    q_value = str(col_Q.value)
                    if q_value.startswith("0"):
                        q_value = q_value[1:]
                    address2 = f"H{q_value}"
                elif "TVA" in str(col_N.value):
                    address1 = "55-550 Naniloa Loop" 
                    address2 = f"TVA {col_Q.value}"
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

        
        return new_file_path

    except Exception as e:
            print(f"Error processing {file_path}: {e}")
            return None

def combine_audit_to_main(file1, file2):
    
    try:
        # Load up and open workesheet of both files
        # Make a new workbook where we will save this info in
        wb1 = openpyxl.load_workbook(file1)
        wb2 = openpyxl.load_workbook(file2)
        ws1 = wb1[1]
        ws2 = wb2.active

        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws(['Student Id', 'Student Name', 'Address 1', 'Address 2'])

        # Build dictionary/Hashmap for Table 2
        # Key = student id, Value ["address 1", "address 2"]
        table2_data = {}
        for row in ws2.iter_rows(min_row=2, values_only=True):
            student_id = row[1]
            address1 = row[2]
            address2 = row[3]

            if student_id:
                table2_data[student_id] = {
                    'address1': address1,
                    'address2': address2
                }
        
        # Loop through table 1 make decisions
        for row in ws1.iter_rows(min_row=2, values_only=True):
            student_id= row[0]
            student_last_name = row[1]
            student_first_name = row[2]

            # NOTE These will probably be changed in the future
            table1_address1 = row[18]
            table1_address2 = row[19]

            if student_id in table2_data:
                table2_info = table2_data[student_id]

                # Check if both addresses filled. NOTE might change this later
                if table2_info['address1'] and table2_info['address2']:
                    #Use table 2 address

                    new_ws.append(student_id, student_last_name, student_first_name, table2_info['address1'], table2_info['address2'])
                else:
                    new_ws.append([student_id, student_last_name, student_first_name, 
                    table1_address1, 
                    table1_address2])
            else:
                    new_ws.append([student_id, student_last_name, student_first_name,
                    table1_address1, 
                    table1_address2])

        base, ext = os.path.splitext(file2)
        new_file_path = f"{base}_finalized_outcome{ext}"
        new_wb.save(new_file_path)

        
        return new_file_path


    except Exception as e:
        print(f"Error processing {file1} or {file2}: {e}")
        return None

def main():

    
    choice = select_task()

    # We format the housing excel sheet
    if choice == 1:
    
        file_path = select_file()

        if not file_path:
            print("No file sleected")
            return
        
        print(f"\nSelected file: {file_path}")
        print("Processing...")

        try:

            new_file_path = audit(file_path)
            print(f"\n Success!")
            print(f" - New file saved as : {new_file_path}")

            #show success dialog
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo(
                "Sucess",
                f"File updated successfully!\n\nNew file: {os.path.basename(new_file_path)}"
            )
            root.destroy()
        
        except Exception as e:
            print(f"\n✗ Error: {str(e)}")
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error", str(e))
            root.destroy()
    
    # We combine it into the main table
    else:
        print(f"\n Please select the main table first")
        print(f"\n ..................") 
        file_path1 = select_file()

        if not file_path1:
            print("No file sleected")
            return
        
        print(f"\n Now select the corrected Housing Audit Workbook")
        print(f"\n ..................")
        file_path2 = select_file()

        if not file_path2:
            print("No file sleected")
            return
        
        try:
            # For security lets create and spit out a new file?
            new_file_path = combine_audit_to_main(file_path1, file_path2)
            print(f"\n Success!")
            print(f" - New file saved as : {new_file_path}")

            #show success dialog
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo(
                "Sucess",
                f"File updated successfully!\n\nNew file: {os.path.basename(new_file_path)}"
            )
            root.destroy()
        except Exception as e:
            print(f"\n✗ Error: {str(e)}")
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error", str(e))
            root.destroy()

    



if __name__ == "__main__":
    main()
