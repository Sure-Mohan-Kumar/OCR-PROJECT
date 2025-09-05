## OCR to Word Document
A simple Python application that extracts text from multiple images and compiles the results into a single Word document. Designed for easy use with a graphical interface and real-time progress feedback.
---
## Features
- Extract text from multiple image files using Tesseract OCR  
- Combine extracted text into one Word document  
- Insert page breaks between text from each image  
- Display a progress bar during processing  
- Simple and intuitive Tkinter-based GUI  
- Cross-platform compatible (Windows, macOS, Linux)
---
## Requirements

- Python 3.7 or above  
- [Tesseract OCR](https://tesseract-ocr.github.io/tessdoc/Installation.html) installed and added to system PATH  
- Python packages: `pytesseract`, `Pillow`, `python-docx`, and `tkinter` (usually included with Python)

  ## 🛠️ Installation & Setup

1. **Clone the repository**:
   ```bash
   git https://github.com/Sure-Mohan-Kumar/OCR-PROJECT.git
   ```
2. **(Optional) Create a virtual environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate        # On Windows: venv\Scripts\activate
   ```
3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
## ▶️ Running the App
  ```bash
  python app.py
  ```
## Usage

- Click **Select Images and Save Text to Word**.  
- Choose multiple images to process (JPEG, PNG, BMP, TIFF).  
- Wait for the progress bar to complete.  
- Select the location to save the combined Word document.  
- Open the Word document to view and edit extracted text, each image’s content separated by a page break.

---
## Screenshots
<img width="594" height="341" alt="image" src="https://github.com/user-attachments/assets/7c55d924-131b-403c-a68f-a15171c5914c" />

