import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import locale

structured_path = None

# Try to set Croatian locale for month names
try:
    locale.setlocale(locale.LC_TIME, "hr_HR.utf8")
except:
    # Fallback if Croatian locale is not available
    pass

def upload_structured():
    global structured_path
    structured_path = filedialog.askopenfilename(
        title="Select 10011 structured.xlsx",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if structured_path:
        file_label.config(text=os.path.basename(structured_path))

def process_file():
    if not structured_path:
        messagebox.showerror("Error", "Please upload the structured file first.")
        return

    try:
        # Load structured Excel file
        df = pd.read_excel(structured_path)

        # Column mapping source → target
        col_mapping = {
            "QTY": "Quantity",
            "Part Number": "Part Number",
            "Component Type": "Type",
            "Filename": "Nomenclature",
            "REV": "Revision",
            "Description": "Product Description"
        }
        df = df.rename(columns=col_mapping)

        # Define final columns
        required_cols = [
            "Quantity",
            "Part Number",
            "Type",
            "Nomenclature",
            "Revision",
            "Product Description"
        ]

        # -----------------------------
        # Build output Excel file
        # -----------------------------
        base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "10011.xlsx")

        # "w" mode ensures overwrite if the file exists
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
            row_pos = 0

            # 1) Date and time in Croatian format
            now = datetime.now()
            cro_date = now.strftime("%d. %B %Y. %H:%M")
            pd.DataFrame([[cro_date]]).to_excel(writer, index=False, header=False, startrow=row_pos)
            row_pos += 2

            # 2) Title
            pd.DataFrame([["Bill Of Materials"]]).to_excel(writer, index=False, header=False, startrow=row_pos)
            row_pos += 2

            # 3) Main parent table (only items without dot)
            df_parents = df[df["Item"].astype(str).str.match(r'^\d+$')]
            parent_table = df_parents[required_cols]
            parent_table.to_excel(writer, index=False, startrow=row_pos)
            row_pos += len(parent_table) + 3

            # 4) For each parent assembly → add BOM section
            for _, parent in df_parents.iterrows():
                if str(parent["Type"]).lower() == "assembly":
                    part_no = parent["Part Number"]

                    # Section title
                    section_title = f"Bill Of Materials : {part_no}"
                    pd.DataFrame([[section_title]]).to_excel(
                        writer, index=False, header=False, startrow=row_pos
                    )
                    row_pos += 2

                    # Child items (those starting with e.g. "2.")
                    item_prefix = str(parent["Item"]) + "."
                    df_children = df[df["Item"].astype(str).str.startswith(item_prefix)]
                    if not df_children.empty:
                        child_table = df_children[required_cols]
                        child_table.to_excel(writer, index=False, startrow=row_pos)
                        row_pos += len(child_table) + 3

        messagebox.showinfo("Success", f"File saved (overwritten if existed):\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{e}")

# --- Tkinter GUI ---
root = tk.Tk()
root.title("Structured Excel Processor")
root.geometry("520x280")

instruction = tk.Label(root, text="Upload '10011 structured.xlsx' to generate '10011.xlsx'")
instruction.pack(pady=10)

upload_btn = tk.Button(root, text="Upload Structured File", command=upload_structured)
upload_btn.pack(pady=5)

file_label = tk.Label(root, text="No file selected", fg="gray")
file_label.pack()

process_btn = tk.Button(root, text="Generate 10011.xlsx", command=process_file, bg="green", fg="white")
process_btn.pack(pady=20)

root.mainloop()
