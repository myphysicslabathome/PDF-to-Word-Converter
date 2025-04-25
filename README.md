# PDF to Word Converter with Drag-and-Drop

A simple Python GUI tool to convert PDF files to Word (`.docx`) format using `pdf2docx` and `tkinter`. Supports file selection through a dialog or drag-and-drop (on supported platforms).

## Features

- 📄 Convert PDF files to Word documents with a single click
- 🖱️ Drag-and-drop support for quick PDF selection
- 🧰 Simple and clean user interface
- 🪄 Automatic file naming and location handling

## Requirements

- Python 3.x
- Dependencies:
  - [`pdf2docx`](https://pypi.org/project/pdf2docx/)
  - [`tkinterdnd2`](https://pypi.org/project/tkinterdnd2/)

## Installation

1. Clone this repository or download the script.
2. Install the required packages:

```bash
pip install pdf2docx tkinterdnd2
```

## Usage
Run the script using Python:
```bash
python Pdf_to_word_converter.py
```

## Instructions:

1. <b>Select a PDF file</b> by clicking "Browse" or by dragging a .pdf file into the entry box.
2. <b>Enter the output file name</b> (without extension).
3. Click <b>"Convert to Word"</b>.
4. The .docx file will be saved in the same directory as the input PDF.


Enjoy converting your PDFs! 😊
