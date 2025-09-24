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
    out = []
    for p in parts:
        p = p.strip()
        if p.isdigit():
            out.append(int(p))
        else:
            digits = "".join(c for c in p if c.isdigit())
            out.append(int(digits) if digits else 0)
    return tuple(out)


def get_direct_children(df: pd.DataFrame, parent_item: str) -> pd.DataFrame:
    """Return only direct children of parent_item (exactly one dot deeper), sorted."""
    p = str(parent_item).strip()
    children = df[df["Item"].apply(
        lambda x: isinstance(x, str) and x.startswith(p + ".") and x.count(".") == p.count(".") + 1
    )].copy()
    if not children.empty:
        children.sort_values(by="Item", key=lambda c: c.map(item_key), inplace=True)
    return children


def write_section(writer, title: str, table: pd.DataFrame, row_pos: int, cols: list) -> int:
    """Write a section title and a table; return new row position."""
    pd.DataFrame([[title]]).to_excel(writer, index=False, header=False, startrow=row_pos)
    row_pos += 2
    if not table.empty:
        table[cols].to_excel(writer, index=False, startrow=row_pos)
        row_pos += len(table) + 3
    else:
        row_pos += 1
    return row_pos


def process_file():
    if not structured_path:
        messagebox.showerror("Error", "Please upload the structured file first.")
        return

    try:
        # Load and normalize
        df = pd.read_excel(structured_path, dtype={"Item": str})
        df["Item"] = df["Item"].astype(str).str.strip()

        # Map columns to final names
        col_mapping = {
            "QTY": "Quantity",
            "Part Number": "Part Number",
            "Component Type": "Type",
            "Filename": "Nomenclature",
            "REV": "Revision",
            "Description": "Product Description",
        }
        df = df.rename(columns=col_mapping)

        # Ensure required cols exist and set types
        req_cols = ["Quantity","Part Number","Type","Nomenclature","Revision","Product Description"]
        for c in req_cols:
            if c not in df.columns:
                df[c] = ""
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
        df["Type"] = df["Type"].astype(str).str.strip()

        # Top-level rows (Item with no dot)
        df_parents = df[df["Item"].str.match(r"^\d+$")].copy()
        if not df_parents.empty:
            df_parents.sort_values(by="Item", key=lambda c: c.map(item_key), inplace=True)

        # Output path
        base_dir = os.path.dirname(os.path.abspath(__file__))
        out_dir = os.path.join(base_dir, "output")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, "10011.xlsx")

        with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
            r = 0

            # Header (Croatian if available)
            now = datetime.now()
            cro_dt = now.strftime("%d. %B %Y. %H:%M")
            pd.DataFrame([[cro_dt]]).to_excel(writer, index=False, header=False, startrow=r); r += 2

            # Title
            pd.DataFrame([["Bill Of Materials"]]).to_excel(writer, index=False, header=False, startrow=r); r += 2

            # Main parent table
            if not df_parents.empty:
                df_parents[req_cols].to_excel(writer, index=False, startrow=r)
                r += len(df_parents) + 3
            else:
                pd.DataFrame([["(No top-level items found)"]]).to_excel(writer, index=False, header=False, startrow=r); r += 3

            # ===== Breadth-first BOM sections =====
            level = [(str(row["Item"]), row["Part Number"])
                     for _, row in df_parents.iterrows()
                     if str(row["Type"]).lower() == "assembly"]

            while level:
                next_level = []
                for parent_item, pn in level:
                    children = get_direct_children(df, parent_item)
                    r = write_section(writer, f"Bill Of Materials : {pn}", children, r, req_cols)
                    if not children.empty:
                        asm_children = children[children["Type"].str.lower() == "assembly"]
                        for _, c in asm_children.iterrows():
                            next_level.append((str(c["Item"]), c["Part Number"]))
                next_level.sort(key=lambda t: item_key(t[0]))
                level = next_level

            # ===== Recapitulation =====
            # Different parts (unique PN over all Part rows)
            parts = df[df["Type"].str.lower() == "part"].copy()
            different_parts = int(parts["Part Number"].nunique())

            # Map Item -> Quantity (sum if duplicates)
            qty_by_item = df.groupby("Item")["Quantity"].sum().to_dict()

            # Robust leaf detection: an item is a parent if ANY other item starts with "<item>."
            items_list = df["Item"].dropna().astype(str).tolist()
            is_parent_map = {itm: any(other.startswith(itm + ".") for other in items_list if other != itm)
                             for itm in items_list}
            leaf_mask = df["Item"].map(lambda x: not is_parent_map.get(x, False))

            # Leaf PART rows only
            leaf_parts = df[leaf_mask & (df["Type"].str.lower() == "part")].copy()

            def ancestors(itm: str):
                segs = itm.split(".")
                # prefixes from deepest-1 up to root, e.g. "4.3.1" -> ["4.3", "4"]
                return [".".join(segs[:i]) for i in range(len(segs)-1, 0, -1)]

            # Non-exploded leaf total (for comparison/debug)
            leaf_raw_total = float(leaf_parts["Quantity"].sum())

            # Exploded total = sum(leaf_qty * product(ancestor_qty))
            exploded_total = 0.0
            for _, row in leaf_parts.iterrows():
                itm = str(row["Item"])
                q = float(row["Quantity"])
                mult = 1.0
                for anc in ancestors(itm):
                    mult *= float(qty_by_item.get(anc, 1))
                exploded_total += q * mult

            total_parts = int(round(exploded_total))

            r += 1
            pd.DataFrame([["Recapitulation"]]).to_excel(writer, index=False, header=False, startrow=r); r += 2
            pd.DataFrame([
                ["Different parts:", different_parts],
                ["Total parts:", total_parts],          # <- exploded
            ]).to_excel(writer, index=False, header=False, startrow=r); r += 4

        messagebox.showinfo("Success", f"File saved (overwritten if existed):\n{out_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{e}")


# --- Tkinter GUI ---
root = tk.Tk()
root.title("Structured Excel Processor")
root.geometry("560x340")

tk.Label(root, text="Upload '10011 structured.xlsx' to generate 'output/10011.xlsx'").pack(pady=10)
tk.Button(root, text="Upload Structured File", command=upload_structured).pack(pady=5)
file_label = tk.Label(root, text="No file selected", fg="gray"); file_label.pack()
tk.Button(root, text="Generate 10011.xlsx", command=process_file, bg="green", fg="white").pack(pady=20)

root.mainloop()
