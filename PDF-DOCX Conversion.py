import os
from pdf2docx import Converter
from pathlib import Path
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog
import threading

def select_input_folder():
    folder = filedialog.askdirectory()
    input_folder_var.set(folder)

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".docx")
    output_file_var.set(file_path)

def start_merge_thread():
    thread = threading.Thread(target=merge_docs)
    thread.start()

def merge_docs():
    input_folder = input_folder_var.get()
    output_file = output_file_var.get()

    if input_folder == "" or output_file == "":
        print("Choose input folder/output file!")
        set_status("Choose input folder/output file!", "red")
        return

    try:
        input_folder_path = Path(input_folder)
        pdf_files = sorted(input_folder_path.glob("*.pdf"))

        if not pdf_files:
            print("No PDF files found.")
            set_status("No PDF files found!", "red")
            return

        merger = PdfMerger()

        for pdf in pdf_files:
            merger.append(str(pdf))
            print(f"Added {pdf.name}")
            set_status(f"Added {pdf.name}", "black")

        # Define output PDF path
        output_pdf_path = output_file.replace(".docx", "") + ".pdf"
        print("Merging...")
        set_status("Merging...", "orange")

        merger.write(output_pdf_path)
        merger.close()

        print(f"Merged PDF saved to: {output_pdf_path}")
        set_status("Converting to DOCX... (This may take a few minutes)", "purple")

        # Convert PDF to DOCX
        cv = Converter(output_pdf_path)
        cv.convert(output_file, start=0, end=None)
        cv.close()

        print(f"DOCX saved to: {output_file}")
        set_status("Done!", "green")

    except Exception as e:
        print("Error during merge/conversion:", e)
        set_status(f"Error: {e}", "red")

def set_status(text, color="blue"):
    root.after(0, lambda: status_label.config(text=text, fg=color)) # Allows updating of status label during process

# Main window
root = tk.Tk()
root.title("PDF-DOCX Merger and Converter")
root.geometry("500x275")

# Variables to store user input
input_folder_var = tk.StringVar()
output_file_var = tk.StringVar()

# --- GUI Layout ---

# Input folder selection
tk.Label(root, text="Select Folder of Files to Merge (Must be PDFs):").pack(pady=5)
tk.Entry(root, textvariable=input_folder_var, width=75).pack()
tk.Button(root, text="Browse", command=select_input_folder).pack(pady=5)

# Output file selection
tk.Label(root, text="Select Output File Location (Will save as both PDF and DOCX):").pack(pady=5)
tk.Entry(root, textvariable=output_file_var, width=75).pack()
tk.Button(root, text="Save As", command=select_output_file).pack(pady=5)

# Merge button
tk.Button(root, text="Merge Files", bg="lightblue", command=start_merge_thread).pack(pady=15)

# Report Label
status_label = tk.Label(root, text="")
status_label.pack(pady=5)

# Start the GUI
root.mainloop()


