import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage, font
import pandas as pd
import os
from datetime import datetime
import locale
import platform
import json
import requests, subprocess, time, sys
from packaging import version
import tempfile

structured_path = None
forklift_number = None  # user-entered value

# --- Updater settings ---
UPDATE_URL = "https://codingbyday.github.io/Parts.Master/latest.json"

# --- Translations ---
translations = {
    "en": {
        "upload_label": "Upload structured file to generate the result!",
        "upload_btn": "📂 Upload Structured File",
        "generate_btn": "⚙️ Generate output file",
        "update_btn": "⬇️ Check for Updates",
        "about": "About",
        "footer": "Version {version} | Created by {author}",
        "forklift_prompt": "Enter Forklift name:",
        "file_select_title": "Select structured source file",
        "save_as": "Save As",
        "success": "Success",
        "success_msg": "File saved:\n{path}",
        "error": "Error",
        "error_msg": "Processing failed:\n{err}",
        'no_file': "No file selected"
    },
    "hr": {
        "upload_label": "Učitaj strukturiranu datoteku za generiranje rezultata!",
        "upload_btn": "📂 Učitaj strukturiranu datoteku",
        "generate_btn": "⚙️ Generiraj izlaznu datoteku",
        "update_btn": "⬇️ Provjeri ažuriranja",
        "about": "O programu",
        "footer": "Verzija {version} | Izradio {author}",
        "forklift_prompt": "Unesite ime viličara:",
        "file_select_title": "Odaberite strukturiranu datoteku",
        "save_as": "Spremi kao",
        "success": "Uspjeh",
        "success_msg": "Datoteka spremljena:\n{path}",
        "error": "Greška",
        "error_msg": "Obrada nije uspjela:\n{err}",
        'no_file': "Nijedna datoteka nije odabrana"
    }
}


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller exe """
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def set_language(lang):
    """Switch UI texts between English and Croatian"""
    t = translations[lang]

    title_label.config(text=t["upload_label"])
    upload_btn.config(text=t["upload_btn"])
    generate_btn.config(text=t["generate_btn"])
    file_label.config(text=translations[lang]["no_file"])
    update_btn.config(text=t["update_btn"])
    about_btn.config(text=t["about"])
    footer_label.config(text=t["footer"].format(version=APP_VERSION, author=APP_AUTHOR))

    current_lang.set(lang)  # update variable
    update_lang_buttons()  



def check_for_updates():
    try:
        url = f"{UPDATE_URL}?t={int(time.time())}"
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
    try:
        # Save to Windows temp folder
        temp_dir = tempfile.gettempdir()
        local_filename = os.path.join(temp_dir, os.path.basename(download_url))

        with requests.get(download_url, stream=True, timeout=30) as r:
            r.raise_for_status()
            with open(local_filename, "wb") as f:
                for chunk in r.iter_content(8192):
                    if chunk:
                        f.write(chunk)

        # Notify user
        messagebox.showinfo("Update", "Installer downloaded. It will now run.")

        # 👉 Detect installer type
        if local_filename.lower().endswith(".msi"):
            subprocess.Popen(['msiexec', '/i', local_filename], shell=True)
        else:
            subprocess.Popen([local_filename], shell=True)

        os._exit(0)  # exit app before installer runs
    except Exception as e:
        messagebox.showerror("Update Failed", str(e))

        
def upload_structured():
    global structured_path, forklift_number
    t = translations[current_lang.get()]

    structured_path = filedialog.askopenfilename(
        title=t["file_select_title"],
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if structured_path:
        file_label.config(text=os.path.basename(structured_path))
        forklift_number = ask_forklift_number(root, "assets/official-logo.ico")


def ask_forklift_number(parent, app_icon=None):
    t = translations[current_lang.get()]
    dialog = tk.Toplevel(parent)
    dialog.title("Forklift Number")
    dialog.geometry("400x180")
    dialog.transient(parent)
    dialog.grab_set()

    if app_icon:
        try:
            dialog.iconbitmap(app_icon)
        except Exception:
            try:
                dialog.iconphoto(False, PhotoImage(file=app_icon))
            except Exception:
                pass

    parent.update_idletasks()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    w, h = 400, 180
    x = px + (pw - w) // 2
    y = py + (ph - h) // 2
    dialog.geometry(f"{w}x{h}+{x}+{y}")

    tk.Label(dialog, text=t["forklift_prompt"], font=("Segoe UI", 11)).pack(pady=15)

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

# Add this near your other helpers
def update_lang_buttons():
    # reset both
    btn_en.config(relief="flat", bd=0, bg="#f5f5f5", activebackground="#f5f5f5")
    btn_hr.config(relief="flat", bd=0, bg="#f5f5f5", activebackground="#f5f5f5")

    # highlight selected
    if current_lang.get() == "en":
        btn_en.config(relief="solid", bd=2, bg="#e8f0ff", activebackground="#e8f0ff")
    else:
        btn_hr.config(relief="solid", bd=2, bg="#e8f0ff", activebackground="#e8f0ff")

def show_about():
    t = translations[current_lang.get()]
    messagebox.showinfo(
        t["about"],
        f"{APP_NAME}\n"
        f"Version {APP_VERSION} ({APP_DATE})\n"
        f"Created by {APP_AUTHOR}\n\n"
        f"Notes:\n{APP_NOTES}"
    )

# --- Load info from JSON ---
base_dir = os.path.dirname(os.path.abspath(__file__))
info_path = resource_path("app_info.json")

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
root.title(f"{APP_NAME}")
root.geometry("700x500")
root.configure(bg="#f5f5f5")
# Prevent window from being resized smaller than the default
root.minsize(700, 500)
# Now safe to create Tkinter variable
current_lang = tk.StringVar(value="en")  # default language

title_font = font.Font(family="Segoe UI", size=14, weight="bold")
normal_font = font.Font(family="Segoe UI", size=11)

# Language selector at top
lang_frame = tk.Frame(root, bg="#f5f5f5")
lang_frame.pack(side="top", pady=10)

flag_en = PhotoImage(file=resource_path("assets/en.png")).subsample(18, 18)
flag_hr = PhotoImage(file=resource_path("assets/hr.png")).subsample(18, 18)

btn_en = tk.Button(
    lang_frame, image=flag_en,
    command=lambda: set_language("en"),
    borderwidth=0, cursor="hand2",
    bg="#f5f5f5", activebackground="#f5f5f5"
)
btn_en.pack(side="left", padx=10)

btn_hr = tk.Button(
    lang_frame, image=flag_hr,
    command=lambda: set_language("hr"),
    borderwidth=0, cursor="hand2",
    bg="#f5f5f5", activebackground="#f5f5f5"
)
btn_hr.pack(side="left", padx=10)

# --- Center frame with card style ---
center_frame = tk.Frame(root, bg="white", padx=40, pady=30, relief="flat", bd=1)
center_frame.place(relx=0.5, rely=0.5, anchor="center")

title_label = tk.Label(center_frame, font=title_font, bg="white", wraplength=500, justify="center")
title_label.pack(pady=(0, 20))

upload_btn = tk.Button(center_frame, command=upload_structured, font=normal_font,
                       bg="#4CAF50", fg="white", activebackground="#45a049",
                       activeforeground="white", padx=15, pady=8,
                       relief="flat", cursor="hand2")
upload_btn.pack(pady=8, fill="x")

file_label = tk.Label(
    center_frame,
    text=translations[current_lang.get()]["no_file"],
    fg="gray", font=normal_font, bg="white"
)

file_label.pack(pady=5)

generate_btn = tk.Button(center_frame, command=lambda: None,  # hook your process_file
                         font=normal_font, bg="#2196F3", fg="white",
                         activebackground="#1976D2", activeforeground="white",
                         padx=15, pady=8, relief="flat", cursor="hand2")
generate_btn.pack(pady=12, fill="x")

update_btn = tk.Button(center_frame, command=check_for_updates, font=normal_font,
                       bg="#FF9800", fg="white", activebackground="#F57C00",
                       activeforeground="white", padx=15, pady=8,
                       relief="flat", cursor="hand2")
update_btn.pack(pady=12, fill="x")

# --- Footer ---
footer_frame = tk.Frame(root, bg="#f5f5f5")
footer_frame.pack(side="bottom", fill="x", pady=5)

footer_label = tk.Label(footer_frame, fg="gray", bg="#f5f5f5", font=("Segoe UI", 9))
footer_label.pack(side="left", padx=10)

about_btn = tk.Button(
    footer_frame,
    text="About",
    command=show_about,   # 👈 now it calls the function
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
try:
    if system == "Windows":
        root.iconbitmap(resource_path("assets/official-logo.ico"))
    else:
        root.iconphoto(False, PhotoImage(file=resource_path("assets/official-logo.png")))
except Exception:
    pass

# Initialize UI with English
set_language("hr")

root.mainloop()
