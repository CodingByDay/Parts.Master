import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

structured_path = None

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
        df_structured = pd.read_excel(structured_path)

        # Keep only parent items (no dot in "Item")
        df_parents = df_structured[df_structured["Item"].astype(str).str.match(r'^\d+$')]

        # Column mapping source → target
        col_mapping = {
            "QTY": "Quantity",
            "Part Number": "Part Number",
            "Component Type": "Type",
            "Filename": "Nomenclature",
            "REV": "Revision",
            "Description": "Product Description"
        }

        # Rename
        df_parents = df_parents.rename(columns=col_mapping)

        # Final order
        required_cols = [
            "Quantity",
            "Part Number",
            "Type",
            "Nomenclature",
            "Revision",
            "Product Description"
        ]
        final_df = df_parents[required_cols]

        # --- Add date/time and title rows ---
        # Current time in Croatian-style format
        now = datetime.now()
        croatian_date = now.strftime("%d. %B %Y. %H:%M")

        # Create two header rows
        header_df = pd.DataFrame([[croatian_date], ["Bill Of Materials"]])

        # Write output
        base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "10011.xlsx")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            header_df.to_excel(writer, index=False, header=False, startrow=0)
            final_df.to_excel(writer, index=False, startrow=3)

        messagebox.showinfo("Success", f"File saved as:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{e}")

# --- Tkinter GUI ---
root = tk.Tk()
root.title("Structured Excel Processor")
root.geometry("500x260")

instruction = tk.Label(root, text="Upload '10011 structured.xlsx' to generate '10011.xlsx'")
instruction.pack(pady=10)

upload_btn = tk.Button(root, text="Upload Structured File", command=upload_structured)
upload_btn.pack(pady=5)

file_label = tk.Label(root, text="No file selected", fg="gray")
file_label.pack()

process_btn = tk.Button(root, text="Generate 10011.xlsx", command=process_file, bg="green", fg="white")
process_btn.pack(pady=20)

root.mainloop()
