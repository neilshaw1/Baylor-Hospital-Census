import pandas as pd
import xlsxwriter
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def process_excel(file_path, output_folder, year_input):
    try:
        int_year = int(year_input)
        if int_year < 1000 or int_year > 9999:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid Year", "Please enter a valid 4-digit year.")
        return

    excel_file = pd.ExcelFile(file_path)
    results = []

    months_in_order = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]

    feb_days = 29 if int_year % 4 == 0 else 28
    month_days = {
        'January': 31, 'February': feb_days, 'March': 31, 'April': 30, 'May': 31,
        'June': 30, 'July': 31, 'August': 31, 'September': 30,
        'October': 31, 'November': 30, 'December': 31
    }

    for i, sheet in enumerate(excel_file.sheet_names):
        if i >= 12:
            break

        df = pd.read_excel(file_path, sheet_name=sheet, header=None)

        if df.empty or df.shape[1] == 0:
            print(f"Skipping sheet '{sheet}': empty or no columns.")
            continue

        try:
            match_condition = df.iloc[:, 0].astype(str).str.contains(
                r"BSLMC Total Census|Census \(from EPIC\)Total Bed Count",
                case=False, na=False
            )
            row_with_census = df[match_condition]
        except Exception as e:
            print(f"Skipping sheet '{sheet}': error finding target row -> {e}")
            continue

        if row_with_census.empty:
            print(f"Warning: no target census row found in sheet '{sheet}'")
            continue

        row_data = pd.to_numeric(row_with_census.iloc[0, 2:], errors='coerce')
        row_data = row_data.replace(0, pd.NA)
        valid_data = row_data.dropna()
        days_with_data = len(valid_data)

        month = months_in_order[i]
        expected_days = month_days[month]
        total_sum = valid_data.sum()

        if days_with_data < expected_days and days_with_data > 0:
            total_sum = (total_sum / days_with_data) * expected_days

        results.append([int_year, month, round(total_sum, 2)])

    if not results:
        messagebox.showerror("Error", "No valid census data found in any sheets.")
        return

    output_file = os.path.join(output_folder, f"Hospital_Monthly_Sums_{int_year}.xlsx")
    df_results = pd.DataFrame(results, columns=['Year', 'Month', 'Calculated Values'])

    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df_results.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

    for col_num, value in enumerate(df_results.columns.values):
        worksheet.write(0, col_num, value, header_format)
    worksheet.set_column(2, 2, 17)
    worksheet.set_column(1, 1, 11)
    writer.close()

def run_gui():
    global entry_file, entry_folder, entry_year

    def select_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            entry_file.delete(0, tk.END)
            entry_file.insert(0, path)

    def select_folder():
        path = filedialog.askdirectory()
        if path:
            entry_folder.delete(0, tk.END)
            entry_folder.insert(0, path)

    def process():
        file_path = entry_file.get()
        folder_path = entry_folder.get()
        year_input = entry_year.get()

        if not file_path or not folder_path or not year_input:
            messagebox.showerror("Error", "Please select input file, output folder, and enter a year.")
            return

        process_excel(file_path, folder_path, year_input)
        messagebox.showinfo("Success", f"File saved to:\n{folder_path}")

    root = tk.Tk()
    root.title("Hospital Monthly Sum Calculator")

    tk.Label(root, text="Select Input Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    entry_file = tk.Entry(root, width=40)
    entry_file.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Upload", command=select_file).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="Select Output Folder:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    entry_folder = tk.Entry(root, width=40)
    entry_folder.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Upload", command=select_folder).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(root, text="Enter Year (e.g. 2024):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    entry_year = tk.Entry(root, width=10)
    entry_year.grid(row=2, column=1, sticky="w", padx=5, pady=5)

    tk.Button(root, text="Process", command=process, bg="green", fg="white").grid(row=3, column=1, pady=10)

    root.mainloop()

run_gui()
