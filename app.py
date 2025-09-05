import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
import pytesseract
from docx import Document

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_and_save():
    files = filedialog.askopenfilenames(
        title="Select Images",
        filetypes=[("Image Files", "*.jpg *.jpeg *.png *.bmp *.tiff")]
    )
    if not files:
        messagebox.showwarning("No files", "No images selected!")
        return
    progress['maximum'] = len(files)
    progress['value'] = 0
    root.update_idletasks()
    doc = Document()
    for i, file_path in enumerate(files):
        img = Image.open(file_path)
        lines = pytesseract.image_to_string(img).strip().split('\n')
        lines = [line for line in lines if line.strip() != '']
        if i > 0:
            doc.add_page_break()
        for line in lines:
            doc.add_paragraph(line)
        progress['value'] = i + 1
        root.update_idletasks()
    save_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")],
        title="Save Word Document"
    )
    if save_path:
        doc.save(save_path)
        messagebox.showinfo("Success", f"Document saved:\n{save_path}")
    else:
        messagebox.showinfo("Cancelled", "Save cancelled.")

root = tk.Tk()
root.title("Batch OCR to Word (With Progress Bar)")
root.geometry("400x200")
btn = tk.Button(root, text="Select Images and Save Text to Word", command=extract_and_save)
btn.pack(pady=20)
progress = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress.pack(pady=10)
root.mainloop()
