import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)

def convert_pdf_to_word():
    pdf_file = pdf_entry.get().strip()
    output_name = output_entry.get().strip()

    if not pdf_file:
        messagebox.showerror("Error", "Please select a PDF file.")
        return
    
    if not output_name:
        messagebox.showerror("Error", "Please enter an output file name.")
        return

    docx_file = os.path.join(os.path.dirname(pdf_file), f"{output_name}.docx")
    
    try:
        cv = Converter(pdf_file)
        cv.convert(docx_file)
        cv.close()
        messagebox.showinfo("Success", f"Conversion successful: {docx_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def on_drop(event):
    # This works on Windows, drops come in as a list of paths
    file_path = event.data.strip('{}')  # Remove braces from path
    if file_path.lower().endswith(".pdf"):
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)
    else:
        messagebox.showerror("Error", "Only PDF files are supported.")

# Create the main window
root = tk.Tk()
root.title("PDF to Word Converter")
root.geometry("500x250")

# Enable DnD support using tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    root.destroy()
    root = TkinterDnD.Tk()
    root.title("PDF to Word Converter with Drag-and-Drop")
    root.geometry("500x250")
except ImportError:
    messagebox.showerror("Missing Module", "Please install tkinterdnd2:\npip install tkinterdnd2")
    exit()

# PDF file selection
tk.Label(root, text="Select PDF file (or drag a file into the box):").pack(pady=5)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(pady=5)
pdf_entry.drop_target_register(DND_FILES)
pdf_entry.dnd_bind('<<Drop>>', on_drop)
tk.Button(root, text="Browse", command=select_pdf).pack(pady=5)

# Output file name input
tk.Label(root, text="Enter output Word file name:").pack(pady=5)
output_entry = tk.Entry(root, width=50)
output_entry.pack(pady=5)

# Convert button
tk.Button(root, text="Convert to Word", command=convert_pdf_to_word).pack(pady=10)

# Run the Tkinter main loop
root.mainloop()
