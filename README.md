# pdf-converter
# File Converter to PDF

## Overview
This is a **Python-based file converter** that converts various file formats into **PDF**. It supports:

 **Text Files (`.txt`)** → Converts plain text into a formatted PDF.  
 **Images (`.jpg`, `.png`, `.jpeg`)** → Converts images into PDF.  
 **Word Documents (`.docx`)** → Converts Word documents into PDF.  
 **Excel Files (`.xls`, `.xlsx`)** → Converts spreadsheets into PDF.  
 **PowerPoint Presentations (`.pptx`)** → Extracts text from slides and saves it as a PDF.

The project also includes a **GUI interface using Tkinter**, allowing users to select a file and start the conversion process with a button click.

---

## Installation
### **Clone the repository**
```
git clone https://github.com/your-repo/file-converter-to-pdf.git
cd file-converter-to-pdf
```

### **Create a virtual environment**
```
python -m venv venv
source venv/bin/activate  # For macOS/Linux
venv\Scripts\activate  # For Windows
```

### **Install dependencies**
Run the following command to install all required libraries:
```
pip install -r requirements.txt
```

---

## Usage
### **Run the program**
```bash
python convert.py
```
This will open the graphical user interface (GUI).

### **Convert a file**
1. Click **"Choose a File"**.
2. Select the file you want to convert.
3. The application will detect the file type and convert it into a **PDF**.
4. A success message will appear with the path of the generated PDF.

---
## Dependencies
This project requires the following Python libraries:

- `fpdf` → For PDF generation
- `pillow (PIL)` → For handling image files
- `pdfkit` → For converting HTML to PDF
- `python-docx` → For extracting text from Word files
- `pandas` → For handling Excel files
- `python-pptx` → For extracting text from PowerPoint slides
- `tkinter` → For the graphical user interface

All dependencies are listed in `requirements.txt` and can be installed using:
```bash
pip install -r requirements.txt
```

## 📬 Contact
For any questions or suggestions, feel free to reach out!
- **Email:** your-email@example.com
- **GitHub:** [Your GitHub Profile](https://github.com/your-profile)
