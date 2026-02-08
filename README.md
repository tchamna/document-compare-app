# 📄 Document Compare

A Streamlit web app that compares two documents side-by-side and highlights the differences. Supports **PowerPoint (.pptx)**, **Word (.docx)**, and **plain text (.txt)** files.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.31+-FF4B4B?logo=streamlit&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

## ✨ Features

- **Upload & compare** two documents directly in your browser
- **Word-level highlighting** — changed words are marked in bold/yellow
- **Download a Word report** (.docx) with all differences
- **Smart matching:**
  - PPTX: order-independent per-slide comparison (handles swapped lines)
  - DOCX/TXT: sequential alignment using `SequenceMatcher`
- **Unicode-safe** — NFC normalisation handles diacritics correctly (great for African languages like Ewondo, Basaa, etc.)
- **Corrupt DOCX fallback** — gracefully handles files with broken embedded images

## 🚀 Quick Start

### 1. Clone the repo

```bash
git clone https://github.com/<your-username>/document-compare-app.git
cd document-compare-app
```

### 2. Create a virtual environment

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the app

```bash
streamlit run app.py
```

The app opens at [http://localhost:8501](http://localhost:8501).

## 📁 Project Structure

```
document-compare-app/
├── app.py                  # Streamlit web interface
├── core/
│   ├── __init__.py         # Public API
│   ├── helpers.py          # Normalisation, text splitting, data classes
│   ├── extractors.py       # PPTX / DOCX / TXT text extraction
│   ├── comparators.py      # Diff algorithms
│   └── report.py           # Word report generation
├── .streamlit/
│   └── config.toml         # Streamlit theme
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

## 🛠️ Supported File Types

| Format | Extension | Comparison Strategy |
|--------|-----------|-------------------|
| PowerPoint | `.pptx` | Order-independent per-slide matching |
| Word | `.docx` | Sequential paragraph alignment |
| Plain text | `.txt` | Sequential line alignment |

> **Note:** Legacy `.doc` files are not supported. Convert them to `.docx` first.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -m 'Add my feature'`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.
