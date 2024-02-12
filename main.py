import tkinter as tk
from tkinter import filedialog, messagebox
import openai
import os
import pandas as pd
import pdfplumber
from docx import Document

# Set your OpenAI API key
openai.api_key = 'YOUR_API_KEY_HERE'

# Initialize the main application window
root = tk.Tk()
root.title("Test Plan Doctor")
root.geometry("800x600")

# Global variables to keep track of the uploaded file path and its type
file_path = ''
file_type = ''

# Function to enhance the content using GPT-3
def gpt3_enhance(content, language="English", max_tokens=500, temperature=0.7, top_p=1):
    """
    Enhance the content using GPT-3 in the specified language with custom parameters.
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

# Function to read Excel files
def read_excel(file_path):
    """
    Read an Excel file and return its content.
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")
        return None

# Function to read PDF files
def read_pdf(file_path):
    """
    Read a PDF file and extract the text.
    """
    text = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + '\n'
    return text

# Function to read Word documents
def read_docx(file_path):
    """
    Read a Word document and extract the text.
    """
    doc = Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])

# Define the file_type_var, file_label, and enhance_button
file_type_var = tk.StringVar(value="Excel")
file_label = tk.Label(root, text="No file uploaded.")
enhance_button = tk.Button(root, text="Start Enhancement")

# Function to handle file uploads
def upload_file():
    """
    Handle the file upload process.
    """
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("PDF Files", "*.pdf"), ("Word Files", "*.docx")])
    if file_path:
        file_label.config(text=f"Uploaded File: {os.path.basename(file_path)}")
        enhance_button.pack()

# Initialize GUI components
file_label.pack()
upload_button = tk.Button(root, text="Upload File", command=upload_file)
upload_button.pack()
enhance_button.config(command=lambda: start_enhancement(file_type_var.get()))

# Function to start the file enhancement process
def start_enhancement(file_type):
    """
    Start the file enhancement process based on the file type⬤