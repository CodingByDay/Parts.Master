import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage, font, simpledialog
import pandas as pd
import os
from datetime import datetime
import locale
import platform
import json
import requests, subprocess, time, sys
from packaging import version

structured_path = None
forklift_number = None  # user-entered value

# --- Updater settings ---
UPDATE_URL = "https://codingbyday.github.io/Parts.Master/latest.json"


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller exe """
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def check_for_updates():
    try:
        url = f"{UPDATE_URL}?t={int(time.time())}"  # cache-buster
        r = requests.get(url, timeout=5, headers={"Cache-Control": "no-cache"})
        r.raise_for_status()
        data = r.json()

        latest = data["version"].strip()
        download_url = data["url"]
        notes = data.get("notes", "")

        if version.parse(latest) > version.parse(APP_VERSION.strip()):
            msg = (f"A new version {latest} is available!\n\n"
                   f"Release notes:\n{notes}\n\n"
                   f"Download and install now?")
            if messagebox.askyesno("Update Available", msg):
                download_and_install(download_url)
        else:
            messagebox.showinfo("Up To Date", "You already have the latest version.")
    except Exception as e:
        messagebox.showerror("Update Error", f"Could not check for updates:\n{e}")


def download_and_install(download_url):
    local_filename = os.path.basename(download_url)
    try:
        with requests.get(download_url, stream=True, timeout=30) as r:
            r.raise_for_status()
            with open(local_filename, "wb") as f:
                for chunk in r.iter_content(8192):
                    if chunk:
                        f.write(chunk)

        messagebox.showinfo("Update", f"Downloaded {local_filename}. Starting installer…")
        subprocess.Popen([local_filename], shell=True)
        os._exit(0)
    except Exception as e:
        messagebox.showerror("Update Failed", str(e))


# Try Croatian locale for month names in date/time
try:
    locale.setlocale(locale.LC_TIME, "hr_HR.utf8")
except Exception:
    pass  # fallback to system locale


def upload_structured():
    global structured_path, forklift_number
    structured_path = filedialog.askopenfilename(
        title="Select structured source file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if structured_path:
        file_label.config(text=os.path.basename(structured_path))
        # Use custom dialog with same app icon
        forklift_number = ask_forklift_number(root, "assets/official-logo.ico")


def ask_forklift_number(parent, app_icon=None):
    dialog = tk.Toplevel(parent)
    dialog.title("Forklift Number")
    dialog.geometry("400x180")  # make it bigger
    dialog.transient(parent)
    dialog.grab_set()

    # --- set same icon ---
    if app_icon:
        try:
            dialog.iconbitmap(app_icon)  # ICO on Windows
        except Exception:
            try:
                dialog.iconphoto(False, PhotoImage(file=app_icon))  # PNG fallback
            except Exception:
                pass

    # --- center relative to parent ---
    parent.update_idletasks()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    w, h = 400, 180
    x = px + (pw - w) // 2
    y = py + (ph - h) // 2
    dialog.geometry(f"{w}x{h}+{x}+{y}")

    tk.Label(dialog, text="Enter Forklift number:", font=("Segoe UI", 11)).pack(pady=15)

    entry = tk.Entry(dialog, font=("Segoe UI", 12))
    entry.pack(pady=5)
    entry.focus()

    value = {"result": None}

    def submit():
        value["result"] = entry.get().strip()
        dialog.destroy()

    tk.Button(
        dialog, text="OK", command=submit,
        bg="#2196F3", fg="white", padx=15, pady=5, cursor="hand2"

    ).pack(pady=15)

    parent.wait_window(dialog)
    return value["result"] or "Unknown"


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


def normalize_str(x):
    s = "" if pd.isna(x) else str(x)
    return s.strip()


def fmt_revision(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # convert floats like "1.0" -> "1"
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return s


def best_title_from_filename(path: str) -> str:
    """Derive a friendly BOM title from the file name."""
    stem = os.path.splitext(os.path.basename(path))[0]
    stem = stem.replace("structured", "").replace("Structured", "").strip("- _")
    return stem or "BOM"


def write_section(writer, title: str, table: pd.DataFrame, row_pos: int, cols: list) -> int:
    """Write a section title and a table; return new row position."""
    pd.DataFrame([[title]]).to_excel(writer, index=False, header=False, startrow=row_pos)
    row_pos += 2
    if not table.empty:
        table[cols].to_excel(writer, index=False, startrow=row_pos)
        row_pos += len(table) + 3
    else:
        pd.DataFrame([["(no rows)"]]).to_excel(writer, index=False, header=False, startrow=row_pos)
        row_pos += 3
    return row_pos


def process_file():
    if not structured_path:
        messagebox.showerror("Error", "Please upload the structured file first.")
        return

    # Ask where to save
    out_path = filedialog.asksaveasfilename(
        title="Save As",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="10011.xlsx"
    )
    if not out_path:
        return  # user cancelled

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

        # Ensure required columns exist and normalize types
        req_cols = ["Quantity", "Part Number", "Type", "Nomenclature", "Revision", "Product Description"]
        for c in req_cols:
            if c not in df.columns:
                df[c] = ""

        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
        df["Type"] = df["Type"].astype(str).str.strip()

        # Normalize for consistent joins
        df["Part Number"] = df["Part Number"].map(normalize_str)
        df["Revision"] = df["Revision"].map(fmt_revision)
        df["Product Description"] = df["Product Description"].map(normalize_str)

        # Top-level rows (Item with no dot)
        df_parents = df[df["Item"].str.match(r"^\d+$")].copy()
        if not df_parents.empty:
            df_parents.sort_values(by="Item", key=lambda c: c.map(item_key), inplace=True)

        with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
            r = 0

            # Header (Croatian, with seconds)
            now = datetime.now()
            cro_dt = now.strftime("%d. %B %Y. %H:%M:%S")
            pd.DataFrame([[cro_dt]]).to_excel(writer, index=False, header=False, startrow=r); r += 2

            # Main title: use Forklift number if provided; else fallback to filename-based title
            title_value = forklift_number if forklift_number else best_title_from_filename(structured_path)
            pd.DataFrame([[f"Bill of Material: {title_value}"]]).to_excel(
                writer, index=False, header=False, startrow=r
            ); r += 2

            # Main parent table
            if not df_parents.empty:
                df_parents[req_cols].to_excel(writer, index=False, startrow=r)
                r += len(df_parents) + 3
            else:
                pd.DataFrame([["(No top-level items found)"]]).to_excel(writer, index=False, header=False, startrow=r); r += 3

            # ===== Breadth-first BOM sections (only direct children) =====
            level = [(str(row["Item"]), row["Part Number"])
                     for _, row in df_parents.iterrows()
                     if str(row["Type"]).lower() == "assembly"]

            while level:
                next_level = []
                for parent_item, pn in level:
                    children = get_direct_children(df, parent_item)
                    # For each assembly section: keep original behavior (show PN)
                    r = write_section(writer, f"Bill of Material: {pn}", children, r, req_cols)
                    if not children.empty:
                        asm_children = children[children["Type"].str.lower() == "assembly"]
                        for _, c in asm_children.iterrows():
                            next_level.append((str(c["Item"]), c["Part Number"]))
                next_level.sort(key=lambda t: item_key(t[0]))
                level = next_level

            # ===== Recapitulation & exploded counts =====
            # Different parts = unique PN among PART rows
            parts_only = df[df["Type"].str.lower() == "part"].copy()
            different_parts = int(parts_only["Part Number"].nunique())

            # Quantities per Item
            qty_by_item = df.groupby("Item")["Quantity"].sum().to_dict()

            # Identify parent vs leaf items
            items_list = df["Item"].dropna().astype(str).tolist()
            is_parent_map = {
                itm: any(other.startswith(itm + ".") for other in items_list if other != itm)
                for itm in items_list
            }
            leaf_mask = df["Item"].map(lambda x: not is_parent_map.get(x, False))

            # Leaf PART rows only
            leaf_parts = df[leaf_mask & (df["Type"].str.lower() == "part")].copy()

            def ancestors(itm: str):
                segs = itm.split(".")
                return [".".join(segs[:i]) for i in range(len(segs)-1, 0, -1)]

            # Exploded total & per-PN aggregations
            exploded_total = 0.0
            per_pn_qty = {}

            for _, row in leaf_parts.iterrows():
                itm = str(row["Item"])
                pn = row["Part Number"]
                q = float(row["Quantity"])
                mult = 1.0
                for anc in ancestors(itm):
                    mult *= float(qty_by_item.get(anc, 1))
                contrib = q * mult
                exploded_total += contrib
                per_pn_qty[pn] = per_pn_qty.get(pn, 0.0) + contrib

            total_parts = int(round(exploded_total))

            # ==== Recapitulation header ====
            r += 1
            pd.DataFrame([[f"Recapitulation of: {title_value}"]]).to_excel(
                writer, index=False, header=False, startrow=r
            ); r += 2
            pd.DataFrame([
                ["Different parts:", different_parts],
                ["Total parts:", total_parts],
            ]).to_excel(writer, index=False, header=False, startrow=r); r += 3

            # ===== Order "Different parts" by first appearance (depth-first) =====
            first_seen = {}
            seq = 0

            def record_part(pn: str):
                nonlocal seq
                if pn not in first_seen:
                    first_seen[pn] = seq
                    seq += 1

            def dfs_from_item(parent_item: str):
                children = get_direct_children(df, parent_item)
                for _, crow in children.iterrows():
                    t = str(crow["Type"]).lower()
                    if t == "part":
                        record_part(crow["Part Number"])
                    elif t == "assembly":
                        dfs_from_item(str(crow["Item"]))

            # Traverse top-level rows (in natural order)
            for _, prow in df_parents.iterrows():
                t = str(prow["Type"]).lower()
                if t == "part":
                    record_part(prow["Part Number"])
                elif t == "assembly":
                    dfs_from_item(str(prow["Item"]))

            # Best available metadata per PN
            meta_parts = parts_only[["Part Number", "Revision", "Product Description"]].copy()
            meta_parts["_desc_ok"] = meta_parts["Product Description"].str.strip().ne("")
            meta_parts["_rev_ok"] = meta_parts["Revision"].str.strip().ne("")
            meta_parts.sort_values(by=["_desc_ok", "_rev_ok"], ascending=[False, False], inplace=True)
            meta = meta_parts.groupby("Part Number", as_index=False).first()[["Part Number", "Revision", "Product Description"]]

            # Prepare list in encounter order
            rows = []
            for pn, qty in per_pn_qty.items():
                rows.append({
                    "Part Number": pn,
                    "Quantity": int(round(qty)),
                    "_order": first_seen.get(pn, 10**9),
                })
            parts_list = pd.DataFrame(rows)

            if not parts_list.empty:
                parts_list = parts_list.merge(meta, on="Part Number", how="left")
                parts_list.sort_values(by=["_order"], inplace=True)
                parts_list.drop(columns=["_order"], inplace=True)

                # Columns exactly as requested (no trailing numbering column)
                parts_list = parts_list[["Quantity", "Part Number", "Revision", "Product Description"]]

                # Write the table directly under Recapitulation counts
                parts_list.to_excel(writer, index=False, startrow=r)
                r += len(parts_list) + 3
            else:
                pd.DataFrame([["(no parts found)"]]).to_excel(writer, index=False, header=False, startrow=r)
                r += 2

        messagebox.showinfo("Success", f"File saved:\n{out_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{e}")


# --- Load info from JSON ---
base_dir = os.path.dirname(os.path.abspath(__file__))
info_path = resource_path("app_info.json")  # important for exe

try:
    with open(info_path, "r", encoding="utf-8") as f:
        app_info = json.load(f)
    APP_NAME = app_info.get("app_name", "My App")
    APP_VERSION = app_info.get("version", "0.0.0")
    APP_AUTHOR = app_info.get("author", "Unknown")
    APP_DATE = app_info.get("release_date", "")
    APP_NOTES = app_info.get("notes", "")
except Exception:
    APP_NAME = "My App"
    APP_VERSION = "0.0.0"
    APP_AUTHOR = "Unknown"
    APP_DATE = ""
    APP_NOTES = ""


# --- Tkinter GUI ---
root = tk.Tk()
root.title(f"{APP_NAME} v{APP_VERSION}")
root.geometry("700x500")
root.configure(bg="#f5f5f5")

# Fonts
title_font = font.Font(family="Segoe UI", size=14, weight="bold")
normal_font = font.Font(family="Segoe UI", size=11)

# --- Center frame with card style ---
center_frame = tk.Frame(root, bg="white", padx=40, pady=30, relief="flat", bd=1)
center_frame.place(relx=0.5, rely=0.5, anchor="center")

# Title label
title_label = tk.Label(
    center_frame,
    text="Upload structured file to generate the result",
    font=title_font,
    bg="white",
    wraplength=500,
    justify="center"
)
title_label.pack(pady=(0, 20))

# Upload button
upload_btn = tk.Button(
    center_frame,
    text="📂 Upload Structured File",
    command=upload_structured,
    font=normal_font,
    bg="#4CAF50",
    fg="white",
    activebackground="#45a049",
    activeforeground="white",
    padx=15,
    pady=8,
    relief="flat",
    cursor="hand2"
)
upload_btn.pack(pady=8, fill="x")

# File label
file_label = tk.Label(center_frame, text="No file selected", fg="gray", font=normal_font, bg="white")
file_label.pack(pady=5)

# Generate button
generate_btn = tk.Button(
    center_frame,
    text="⚙️ Generate output file",
    command=process_file,
    font=normal_font,
    bg="#2196F3",
    fg="white",
    activebackground="#1976D2",
    activeforeground="white",
    padx=15,
    pady=8,
    relief="flat",
    cursor="hand2"
)
generate_btn.pack(pady=12, fill="x")

# Check for Updates button
update_btn = tk.Button(
    center_frame,
    text="⬇️ Check for Updates",
    command=check_for_updates,
    font=normal_font,
    bg="#FF9800",
    fg="white",
    activebackground="#F57C00",
    activeforeground="white",
    padx=15,
    pady=8,
    relief="flat",
    cursor="hand2"
)
update_btn.pack(pady=12, fill="x")

# --- Footer ---
footer_frame = tk.Frame(root, bg="#f5f5f5")
footer_frame.pack(side="bottom", fill="x", pady=5)

footer_label = tk.Label(
    footer_frame,
    text=f"Version {APP_VERSION} | Created by {APP_AUTHOR}",
    fg="gray",
    bg="#f5f5f5",
    font=("Segoe UI", 9)
)
footer_label.pack(side="left", padx=10)

def show_about():
    messagebox.showinfo(
        "About",
        f"{APP_NAME}\n"
        f"Version {APP_VERSION} ({APP_DATE})\n"
        f"Created by {APP_AUTHOR}\n\n"
        f"Notes:\n{APP_NOTES}"
    )

about_btn = tk.Button(
    footer_frame,
    text="About",
    command=show_about,
    font=("Segoe UI", 9),
    bg="#e0e0e0",
    fg="black",
    relief="flat",
    padx=10,
    pady=2,
    cursor="hand2"
)
about_btn.pack(side="right", padx=10)

# --- Icon handling ---
system = platform.system()
try:
    if system == "Windows":
        root.iconbitmap(resource_path("assets/official-logo.ico"))
    else:
        root.iconphoto(False, PhotoImage(file=resource_path("assets/official-logo.png")))
except Exception:
    pass  # no icon fallback

root.mainloop()
