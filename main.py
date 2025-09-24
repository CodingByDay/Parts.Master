import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import locale

structured_path = None

# Try Croatian locale for month names in date/time
try:
    locale.setlocale(locale.LC_TIME, "hr_HR.utf8")
except:
    pass  # fallback to system locale


def upload_structured():
    global structured_path
    structured_path = filedialog.askopenfilename(
        title="Select 10011 structured.xlsx",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if structured_path:
        file_label.config(text=os.path.basename(structured_path))


def item_key(s: str):
    """Natural sort key for 'Item' like '4', '4.3', '4.10.2'."""
    parts = str(s).split(".")
    key = []
    for p in parts:
        p = p.strip()
        if p.isdigit():
            key.append(int(p))
        else:
            # Fallback if any non-digit sneaks in
            try:
                key.append(int("".join([c for c in p if c.isdigit()])))
            except Exception:
                key.append(0)
    return tuple(key)


def get_direct_children(df: pd.DataFrame, parent_item: str) -> pd.DataFrame:
    """Return only direct children of parent_item (exactly one dot deeper), sorted."""
    p = str(parent_item).strip()
    children = df[df["Item"].apply(
        lambda x: isinstance(x, str) and x.startswith(p + ".") and x.count(".") == p.count(".") + 1
    )].copy()
    if not children.empty:
        children.sort_values(by="Item", key=lambda col: col.map(item_key), inplace=True)
    return children


def write_section(writer, section_title: str, table: pd.DataFrame, row_pos: int, required_cols: list) -> int:
    """Write a section title and a table; return new row position."""
    # Title
    pd.DataFrame([[section_title]]).to_excel(writer, index=False, header=False, startrow=row_pos)
    row_pos += 2
    # Table
    if not table.empty:
        table[required_cols].to_excel(writer, index=False, startrow=row_pos)
        row_pos += len(table) + 3
    else:
        # Still leave some space if no children
        row_pos += 1
    return row_pos


def process_file():
    if not structured_path:
        messagebox.showerror("Error", "Please upload the structured file first.")
        return

    try:
        # Read and normalize
        df = pd.read_excel(structured_path, dtype={"Item": str})
        df["Item"] = df["Item"].astype(str).str.strip()

        # Map source columns -> final names
        col_mapping = {
            "QTY": "Quantity",
            "Part Number": "Part Number",
            "Component Type": "Type",
            "Filename": "Nomenclature",
            "REV": "Revision",
            "Description": "Product Description",
        }
        df = df.rename(columns=col_mapping)

        required_cols = [
            "Quantity",
            "Part Number",
            "Type",
            "Nomenclature",
            "Revision",
            "Product Description",
        ]

        # Ensure required columns exist (fill blanks if missing)
        for c in required_cols:
            if c not in df.columns:
                df[c] = ""

        # Top-level (parents): items without dot
        df_parents = df[df["Item"].str.match(r"^\d+$")].copy()
        if not df_parents.empty:
            df_parents.sort_values(by="Item", key=lambda col: col.map(item_key), inplace=True)

        # Output path
        base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "10011.xlsx")

        with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
            row_pos = 0

            # Header date/time (Croatian if available)
            now = datetime.now()
            cro_date = now.strftime("%d. %B %Y. %H:%M")
            pd.DataFrame([[cro_date]]).to_excel(writer, index=False, header=False, startrow=row_pos)
            row_pos += 2

            # Main title
            pd.DataFrame([["Bill Of Materials"]]).to_excel(writer, index=False, header=False, startrow=row_pos)
            row_pos += 2

            # Main parent table (all top-level items)
            if not df_parents.empty:
                df_parents[required_cols].to_excel(writer, index=False, startrow=row_pos)
                row_pos += len(df_parents) + 3
            else:
                pd.DataFrame([["(No top-level items found)"]]).to_excel(writer, index=False, header=False, startrow=row_pos)
                row_pos += 3

            # ===== Breadth-First BOM sections =====
            # Level 1 assemblies (top-level)
            current_level = []
            for _, r in df_parents.iterrows():
                if str(r["Type"]).lower() == "assembly":
                    current_level.append((str(r["Item"]), r["Part Number"]))

            # Iterate level by level
            while current_level:
                next_level = []

                # Write sections for all assemblies at this level
                for parent_item, part_no in current_level:
                    # Direct children table for this assembly
                    children = get_direct_children(df, parent_item)
                    # Sort children by item for stable order
                    section_title = f"Bill Of Materials : {part_no}"
                    row_pos = write_section(writer, section_title, children, row_pos, required_cols)

                    # Collect child assemblies for next level
                    if not children.empty:
                        asm_children = children[children["Type"].astype(str).str.lower() == "assembly"]
                        for _, cr in asm_children.iterrows():
                            next_level.append((str(cr["Item"]), cr["Part Number"]))

                # Sort next level assemblies by their Item for consistent ordering
                next_level.sort(key=lambda t: item_key(t[0]))
                current_level = next_level

        messagebox.showinfo("Success", f"File saved (overwritten if existed):\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{e}")


# --- Tkinter GUI ---
root = tk.Tk()
root.title("Structured Excel Processor")
root.geometry("560x300")

instruction = tk.Label(root, text="Upload '10011 structured.xlsx' to generate 'output/10011.xlsx'")
instruction.pack(pady=10)

upload_btn = tk.Button(root, text="Upload Structured File", command=upload_structured)
upload_btn.pack(pady=5)

file_label = tk.Label(root, text="No file selected", fg="gray")
file_label.pack()

process_btn = tk.Button(root, text="Generate 10011.xlsx", command=process_file, bg="green", fg="white")
process_btn.pack(pady=20)

root.mainloop()
