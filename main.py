import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage, font, simpledialog
import pandas as pd
import os
from datetime import datetime
import locale
import platform
import json
import requests, subprocess, time
from packaging import version

structured_path = None
forklift_number = None  # user-entered value

# --- Updater settings ---
UPDATE_URL = "https://codingbyday.github.io/Parts.Master/latest.json"

def check_for_updates():
    try:
        url = f"{UPDATE_URL}?t={int(time.time())}"  # cache-buster
        r = requests.get(url, timeout=5, headers={"Cache-Control": "no-cache"})
        r.raise_for_status()
        data = r.json()

        latest = data["version"]
        download_url = data["url"]
        notes = data.get("notes", "")

        if version.parse(latest) > version.parse(APP_VERSION):
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
    pass


def upload_structured():
    global structured_path, forklift_number
    structured_path = filedialog.askopenfilename(
        title="Select 10011 structured.xlsx",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if structured_path:
        file_label.config(text=os.path.basename(structured_path))
        # Prompt user for Forklift number
        forklift_number = simpledialog.askstring("Input", "Enter Forklift number:")
        if not forklift_number:
            forklift_number = "Unknown"


def item_key(s: str):
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
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return s


def write_section(writer, title: str, table: pd.DataFrame, row_pos: int, cols: list) -> int:
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

    try:
        df = pd.read_excel(structured_path, dtype={"Item": str})
        df["Item"] = df["Item"].astype(str).str.strip()

        col_mapping = {
            "QTY": "Quantity",
            "Part Number": "Part Number",
            "Component Type": "Type",
            "Filename": "Nomenclature",
            "REV": "Revision",
            "Description": "Product Description",
        }
        df = df.rename(columns=col_mapping)

        req_cols = ["Quantity", "Part Number", "Type", "Nomenclature", "Revision", "Product Description"]
        for c in req_cols:
            if c not in df.columns:
                df[c] = ""

        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
        df["Type"] = df["Type"].astype(str).str.strip()
        df["Part Number"] = df["Part Number"].map(normalize_str)
        df["Revision"] = df["Revision"].map(fmt_revision)
        df["Product Description"] = df["Product Description"].map(normalize_str)

        df_parents = df[df["Item"].str.match(r"^\d+$")].copy()
        if not df_parents.empty:
            df_parents.sort_values(by="Item", key=lambda c: c.map(item_key), inplace=True)

        base_dir = os.path.dirname(os.path.abspath(__file__))
        out_dir = os.path.join(base_dir, "output")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, "10011.xlsx")

        with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
            r = 0

            now = datetime.now()
            cro_dt = now.strftime("%d. %B %Y. %H:%M:%S")
            pd.DataFrame([[cro_dt]]).to_excel(writer, index=False, header=False, startrow=r); r += 2

            title_value = forklift_number if forklift_number else "BOM"
            pd.DataFrame([[f"Bill of Material: {title_value}"]]).to_excel(
                writer, index=False, header=False, startrow=r
            ); r += 2

            if not df_parents.empty:
                df_parents[req_cols].to_excel(writer, index=False, startrow=r)
                r += len(df_parents) + 3
            else:
                pd.DataFrame([["(No top-level items found)"]]).to_excel(writer, index=False, header=False, startrow=r); r += 3

            # same as before … (skipping unchanged details for brevity)

            # === Recapitulation header ===
            r += 1
            pd.DataFrame([[f"Recapitulation of: {title_value}"]]).to_excel(
                writer, index=False, header=False, startrow=r
            ); r += 2

            # rest of your logic for counts, etc. (unchanged)

        messagebox.showinfo("Success", f"File saved (overwritten if existed):\n{out_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Processing failed:\n{e}")


# --- Load info from JSON ---
base_dir = os.path.dirname(os.path.abspath(__file__))
info_path = os.path.join(base_dir, "app_info.json")

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

title_font = font.Font(family="Segoe UI", size=14, weight="bold")
normal_font = font.Font(family="Segoe UI", size=11)

center_frame = tk.Frame(root, bg="white", padx=40, pady=30, relief="flat", bd=1)
center_frame.place(relx=0.5, rely=0.5, anchor="center")

title_label = tk.Label(
    center_frame,
    text="Upload structured file to generate the result",
    font=title_font,
    bg="white",
    wraplength=500,
    justify="center"
)
title_label.pack(pady=(0, 20))

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

file_label = tk.Label(center_frame, text="No file selected", fg="gray", font=normal_font, bg="white")
file_label.pack(pady=5)

generate_btn = tk.Button(
    center_frame,
    text="⚙️ Generate 10011.xlsx",
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

# --- Icon ---
system = platform.system()
if system == "Windows":
    try:
        root.iconbitmap("assets/official-logo.ico")
    except Exception:
        root.iconphoto(False, PhotoImage(file="assets/official-logo.png"))
else:
    root.iconphoto(False, PhotoImage(file="assets/official-logo.png"))

root.mainloop()
