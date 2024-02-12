import tkinter as tk
from tkinter import filedialog, messagebox
import openai
import os
import pandas as pd
import pdfplumber
from docx import Document

# Set your OpenAI API key
openai.api_key = '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'

# Initialize the main application window
root = tk.Tk()
root.title("Test Plan Doctor")
root.geometry("800x600")

# Variables to keep track of the uploaded file path and its type
file_path = ''
file_type = ''

def gpt3_enhance(content, language="English", max_tokens=500, temperature=0.7, top_p=1):
    """
    Enhances the provided content using GPT-3 in the specified language.
    """
    prompt = f"[Translate this test plan to {language}]\n\n{content}" if language != "English" else content
    try:
        response = openai.Completion.create(
            engine="text-davinci-003",
            prompt=prompt,
            max_tokens=max_tokens,
            temperature=temperature,
            top_p=top_p
        )
        return response.choices[0].text.strip()
    except Exception as e:
        messagebox.showerror("Error", f"Error during GPT-3 enhancement: {e}")
        return None

def read_excel(file_path):
    """
    Reads an Excel file and returns its content.
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")
        return None

def read_pdf(file_path):
    """
    Reads a PDF file and extracts the text.
    """
    text = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + '\n'
    return text

def read_docx(file_path):
    """
    Reads a Word document and extracts the text.
    """
    doc = Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])

def upload_file():
    """
    Handles the file upload process and sets the file type based on the extension.
    """
    global file_path, file_type
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("PDF Files", "*.pdf"), ("Word Files", "*.docx")])
    if file_path:
        file_label.config(text=f"Uploaded File: {os.path.basename(file_path)}")
        if file_path.endswith('.xlsx'):
            file_type = 'Excel'
        elif file_path.endswith('.pdf'):
            file_type = 'PDF'
        elif file_path.endswith('.docx'):
            file_type = 'Word'
        enhance_button.pack()

def start_enhancement():
    """
    Starts the file enhancement process based on the file type.
    """
    global file_type
    try:
        if file_type == "Excel":
            df = read_excel(file_path)
            # Further processing for Excel...
        elif file_type == "PDF":
            pdf_text = read_pdf(file_path)
            enhanced_text = gpt3_enhance(pdf_text)
            print(enhanced_text)  # Example action
        elif file_type == "Word":
            doc_text = read_docx(file_path)
            enhanced_text = gpt3_enhance(doc_text)
            print(enhanced_text)  # Example action
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI setup
file_label = tk.Label(root, text="No file uploaded.")
file_label.pack()
upload_button = tk.Button(root, text="Upload File", command=upload_file)
upload_button.pack()
enhance_button = tk.Button(root, text="Start Enhancement", command=start_enhancement)
# enhance_button is packed in the upload_file function after a file is successfully uploaded

# Start the main application loop
root.mainloop()