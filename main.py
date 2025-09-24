import tkinter as tk
from tkinter import filedialog, messagebox

def upload_file1():
    file_path = filedialog.askopenfilename(
        title="Select First Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        file1_label.config(text=file_path)

def upload_file2():
    file_path = filedialog.askopenfilename(
        title="Select Second Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        file2_label.config(text=file_path)

# Create main window
root = tk.Tk()
root.title("Excel File Uploader")
root.geometry("500x200")

# Instructions
instruction = tk.Label(root, text="Upload two Excel files (.xlsx or .xls)", font=("Arial", 12))
instruction.pack(pady=10)

# File 1 upload
file1_btn = tk.Button(root, text="Upload First File", command=upload_file1)
file1_btn.pack(pady=5)
file1_label = tk.Label(root, text="No file selected", fg="gray")
file1_label.pack()

# File 2 upload
file2_btn = tk.Button(root, text="Upload Second File", command=upload_file2)
file2_btn.pack(pady=5)
file2_label = tk.Label(root, text="No file selected", fg="gray")
file2_label.pack()

# Run the app
root.mainloop()
