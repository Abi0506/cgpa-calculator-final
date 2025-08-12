import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os
from datetime import datetime

# Grade mapping
grade_points_map = {
    "O": 10, "A+": 9, "A": 8, "B+": 7, "B": 6, "C": 5,
    "U": 0, "SA": 0, "WD": 0
}

def calculate_sgpa_cgpa(df):
    # Convert grades to points
    df['GradePoints'] = df['GRADE'].map(grade_points_map)
    df['CURRSEMS'] = pd.to_numeric(df['CURRSEMS'], errors='coerce')
    df['Credits'] = pd.to_numeric(df['Credits'], errors='coerce')

    sgpa_data = {}

    # Calculate SGPA per semester
    for sem in range(1, 11):
        sem_df = df[df['CURRSEMS'] == sem]
        if not sem_df.empty:
            sgpa_series = sem_df.groupby('715521YYYYYY').apply(
                lambda x: round((x['Credits'] * x['GradePoints']).sum() / x['Credits'].sum(), 2)
                if x['Credits'].sum() > 0 else "-"
            )
        else:
            sgpa_series = pd.Series("-", index=df['715521YYYYYY'].unique())

        sgpa_data[f"SEM{sem}_SGPA"] = sgpa_series

    # Calculate CGPA
    cgpa_series = df.groupby('715521YYYYYY').apply(
        lambda x: round((x['Credits'] * x['GradePoints']).sum() / x['Credits'].sum(), 2)
        if x['Credits'].sum() > 0 else "-"
    )

    # Prepare result DataFrame
    result = df[['INSTCODE', '715521YYYYYY', 'STUDNAME', 'BRANNAME']].drop_duplicates().copy()
    for sem in range(1, 11):
        result[f"SEM{sem}_SGPA"] = result['715521YYYYYY'].map(lambda r: sgpa_data[f"SEM{sem}_SGPA"].get(r, "-"))
    result["CGPA"] = result['715521YYYYYY'].map(lambda r: cgpa_series.get(r, "-"))

    # Sort by roll number
    result = result.sort_values(by='715521YYYYYY')

    return result

def format_excel_center(file_path):
    wb = load_workbook(file_path)
    ws = wb.active

    # Center align all cells
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Auto adjust column width
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(file_path)

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, dtype=str)  # Load all as strings to avoid dtype issues
        required_cols = ["INSTCODE", "715521YYYYYY", "STUDNAME", "BRANNAME", "CURRSEMS", "GRADE", "Credits"]

        if not all(col in df.columns for col in required_cols):
            messagebox.showerror("Error", f"Excel must have columns: {', '.join(required_cols)}")
            return

        # 🔹 Cross-INSTCODE duplicate check
        roll_to_inst = df.groupby("715521YYYYYY")["INSTCODE"].nunique()
        problem_rolls = roll_to_inst[roll_to_inst > 1].index.tolist()

        if problem_rolls:
            dup_rows = df[df["715521YYYYYY"].isin(problem_rolls)][["INSTCODE", "715521YYYYYY"]].drop_duplicates()
            dup_list = "\n".join(f"Roll No: {rn} found in INSTCODE(s): {', '.join(map(str, df[df['715521YYYYYY'] == rn]['INSTCODE'].unique()))}"
                                 for rn in problem_rolls)
            messagebox.showerror(
                "Cross-INSTCODE Duplicate Error",
                f"The following roll numbers appear in more than one INSTCODE:\n\n{dup_list}"
            )
            return

        # Convert numeric columns where needed
        df["CURRSEMS"] = pd.to_numeric(df["CURRSEMS"], errors="coerce")
        df["Credits"] = pd.to_numeric(df["Credits"], errors="coerce")

        result_df = calculate_sgpa_cgpa(df)

        # Save output Excel with timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_filename = f"SGPA_CGPA_Output_{timestamp}.xlsx"
        output_path = os.path.join(os.path.dirname(file_path), output_filename)
        result_df.to_excel(output_path, index=False)

        # Format Excel
        format_excel_center(output_path)

        # Show in Tkinter table
        show_table(result_df)

        messagebox.showinfo("Success", f"SGPA & CGPA calculated!\nSaved to: {output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def show_table(df):
    for widget in frame_table.winfo_children():
        widget.destroy()

    tree = ttk.Treeview(frame_table, columns=list(df.columns), show='headings')
    tree.pack(fill='both', expand=True)

    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=100)

    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

root = tk.Tk()
root.title("SGPA & CGPA Calculator")
root.geometry("900x600")  # Not full screen

btn_select = tk.Button(root, text="Select Excel File", command=select_file)
btn_select.pack(pady=10)

frame_table = tk.Frame(root)
frame_table.pack(fill='both', expand=True)

root.mainloop()
