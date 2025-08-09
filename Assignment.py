import tkinter as tk
from tkinter import filedialog
import docx
import fitz  # PyMuPDF
import language_tool_python
import pytesseract
from PIL import Image
import io

# Optional: Set Tesseract path for Windows users
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_text_from_pdf(file_path):
    text = ""
    pdf_document = fitz.open(file_path)

    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        page_text = page.get_text()

        if page_text.strip():  # Normal text extraction
            text += page_text
        else:
            # OCR for scanned PDF pages
            pix = page.get_pixmap()
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            text += pytesseract.image_to_string(img)

    return text

def extract_text_from_word(file_path):
    doc = docx.Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_txt(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()

def check_grammar_and_spelling(text):
    tool = language_tool_python.LanguageTool('en-US')
    matches = tool.check(text)

    if not matches:
        print("No errors found. Everything is alright ✅")
        return

    for match in matches:
        sentence = match.context.strip()
        error_text = match.context[match.offset:match.offset + match.errorLength]
        suggestions = ", ".join(match.replacements) if match.replacements else "No suggestions"

        print(f"Sentence : {sentence}")
        print(f"Error    : {error_text}")
        print(f"Suggestion : {suggestions}")
        print("-" * 50)

def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select PDF, Word, or Text file",
        filetypes=[
            ("All supported files", "*.pdf *.docx *.txt"),
            ("PDF files", "*.pdf"),
            ("Word files", "*.docx"),
            ("Text files", "*.txt")
        ]
    )

    if not file_path:
        print("No file selected.")
        return

    if file_path.lower().endswith(".pdf"):
        text = extract_text_from_pdf(file_path)
    elif file_path.lower().endswith(".docx"):
        text = extract_text_from_word(file_path)
    elif file_path.lower().endswith(".txt"):
        text = extract_text_from_txt(file_path)
    else:
        print("Unsupported file format.")
        return

    check_grammar_and_spelling(text)

if __name__ == "__main__":
    main()
