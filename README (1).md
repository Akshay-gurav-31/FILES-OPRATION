# 📄 FooterForge — Add Footers to PDF & DOCX

A clean, production-style Flask web app that lets you upload a **PDF** or **DOCX** file, type in custom footer text, and instantly download the processed file with a footer added to **every page**.

---

## ✨ Features

- **PDF**: footer stamped on every page using ReportLab + PyPDF2 (includes page X of Y)
- **DOCX**: footer added to every section using python-docx (includes PAGE and NUMPAGES fields)
- Drag-and-drop or click-to-browse file upload
- Unique filenames prevent file collisions
- Proper error handling for missing files, wrong formats, and processing failures
- No external frontend frameworks — pure HTML, CSS, and JavaScript

---

## 🗂️ Project Structure

```
footerforge/
├── app.py                  # Flask backend
├── templates/
│   └── index.html          # Single-file frontend (HTML + CSS + JS)
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── uploads/                # Created automatically — stores incoming files
└── outputs/                # Created automatically — stores processed files
```

---

## ⚙️ Setup & Installation

### 1. Clone or download the project

```bash
git clone https://github.com/your-username/footerforge.git
cd footerforge
```

### 2. Create and activate a virtual environment (recommended)

```bash
python -m venv venv

# macOS / Linux
source venv/bin/activate

# Windows
venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

---

## 🚀 Run the App

```bash
python app.py
```

Then open your browser and visit:

```
http://127.0.0.1:5000
```

---

## 🧪 Example Usage

1. Open the app in your browser.
2. Drag and drop (or click to browse) a `.pdf` or `.docx` file.
3. Type your footer text in the input box — e.g.:  
   `Confidential — Acme Corp © 2025`
4. Click **Add Footer**.
5. The processed file downloads automatically with the footer on every page.

---

## 📦 Dependencies

| Package        | Purpose                                      |
|----------------|----------------------------------------------|
| Flask          | Web framework / routing                      |
| PyPDF2         | Reading and writing PDF pages                |
| ReportLab      | Rendering the footer overlay onto PDF pages  |
| python-docx    | Adding/replacing footers in DOCX sections    |
| Werkzeug       | Secure filename handling (ships with Flask)  |

---

## 🔧 Configuration Notes

- Uploaded files are stored in `uploads/` with a unique suffix to avoid overwriting.
- Processed files are stored in `outputs/` with `_with_footer` appended to the name.
- Both folders are created automatically on first run.
- For production deployment, consider using Gunicorn and periodically cleaning the `uploads/` and `outputs/` directories.

---

## 🐛 Troubleshooting

| Issue | Fix |
|---|---|
| `ModuleNotFoundError` | Make sure you ran `pip install -r requirements.txt` inside your virtualenv |
| Footer not visible in PDF viewer | Some viewers cache aggressively — try opening in a different PDF reader |
| DOCX page numbers show as field codes | Open in Microsoft Word and press `Ctrl+A` then `F9` to refresh fields |
| Port 5000 already in use | Run with `python app.py` after changing `port=5001` at the bottom of `app.py` |
