import tkinter as tk
from tkinter import ttk, filedialog
import os
import win32com.client
import win32api
import win32print
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

def get_printers():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
    return [printer[2] for printer in printers]


def print_word_docs(folder_path, printer_name):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    folder_path = folder_path.replace('/', '\\')  # Normalize backslashes to forward slashes

    # Get the current default printer
    current_printer = win32print.GetDefaultPrinter()

    # Set the default printer to the desired one
    win32print.SetDefaultPrinter(printer_name)

    for filename in os.listdir(folder_path):
        if filename.endswith((".doc", ".docx")):
            file_path = os.path.join(folder_path, filename)
            print(f"Full file path: {file_path}")
            if os.path.exists(file_path):
                try:
                    doc = word.Documents.Open(file_path)
                    doc.PrintOut()
                    doc.Close(SaveChanges=False)
                    print(f"Printed Word document: {filename}")
                except Exception as e:
                    print(f"Failed to print {filename}: {e}")
            else:
                print(f"File does not exist: {file_path}")

    # Revert to the original default printer
    win32print.SetDefaultPrinter(current_printer)
    word.Quit()


def print_excel_files(folder_path, printer_name):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    for filename in os.listdir(folder_path):
        if filename.endswith((".xlsx", ".csv")):
            file_path = os.path.join(folder_path, filename)
            print(f"Attempting to print: {file_path} to {printer_name}")  # Debug print
            try:
                workbook = excel.Workbooks.Open(file_path)
                workbook.PrintOut(PrinterName=printer_name)
                workbook.Close(SaveChanges=False)
                print(f"Printed Excel file: {filename}")
            except Exception as e:
                print(f"Failed to print {filename}: {e}")
    excel.Quit()

def print_pdfs(folder_path, printer_name):
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            print(f"Attempting to print: {file_path} to {printer_name}")  # Debug print
            try:
                win32api.ShellExecute(0, "print", file_path, None, ".", 0)
                print(f"Printed PDF file: {filename}")
            except Exception as e:
                print(f"Failed to print {filename}: {e}")

def print_all_files():
    folder_path = folder_var.get()
    printer_name = printer_var.get()
    if os.path.isdir(folder_path):
        print_word_docs(folder_path, printer_name)
        print_excel_files(folder_path, printer_name)
        print_pdfs(folder_path, printer_name)
    else:
        print("Invalid folder path")

def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_var.set(folder_selected)
    update_file_list(folder_selected)

def update_file_list(folder_path):
    file_text.delete(1.0, tk.END)
    for filename in os.listdir(folder_path):
        if filename.endswith((".doc", ".docx", ".xlsx", ".csv", ".pdf")):
            file_text.insert(tk.END, filename + '\n', 'valid')
        else:
            file_text.insert(tk.END, filename + '\n', 'invalid')

# Initialize Tkinter root window
root = tk.Tk()
root.title("Print Documents")

folder_var = tk.StringVar()
printer_var = tk.StringVar()

# Set up ttkbootstrap styling
style = ttk.Style()
style.configure('TLabel', font=("Helvetica", 12))
style.configure('TButton', font=("Helvetica", 12))

# Widgets
ttk.Label(root, text="Select the folder containing .doc, .docx, .xlsx, .csv, and .pdf files:").grid(row=0, column=0, columnspan=3, pady=10)
ttk.Entry(root, textvariable=folder_var, width=50).grid(row=1, column=0, columnspan=2, padx=10)
ttk.Button(root, text="Browse", command=browse_folder, bootstyle=PRIMARY).grid(row=1, column=2, padx=10)

# Printer dropdown
printers = get_printers()
ttk.Label(root, text="Select Printer:").grid(row=2, column=0, pady=10)
printer_dropdown = ttk.Combobox(root, textvariable=printer_var, values=printers, width=50)
printer_dropdown.grid(row=2, column=1, columnspan=2, padx=10)
printer_dropdown.set(printers[0] if printers else "No printers found")

ttk.Button(root, text="Print All Files", command=print_all_files, bootstyle=SUCCESS).grid(row=3, column=0, columnspan=3, pady=10)

# File Text Widget
file_text = tk.Text(root, width=80, height=10, wrap=tk.NONE)
file_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)
file_text.tag_configure('valid', foreground='black')
file_text.tag_configure('invalid', foreground='gray')

root.mainloop()
