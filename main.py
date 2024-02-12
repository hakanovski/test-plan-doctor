import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openai
import os
import pandas as pd
import PyPDF2
import pdfplumber
from docx import Document

# OpenAI API Key
openai.api_key = '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'

def gpt3_enhance(content, language="English", max_tokens=500, temperature=0.7, top_p=1):
    """Enhance the content using GPT-3 in the specified language with custom parameters."""
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

def analyze_content(content):
    """Analyze the content for quality and coverage."""
    word_count, quality = len(content.split()), "Good Quality" if len(content.split()) > 50 else "Needs Improvement"
    return word_count, quality

def generate_report(content):
    """Generate and save a report based on the analysis."""
    word_count, quality = analyze_content(content)
    report = f"Content Analysis Report:\nWord Count: {word_count}\nQuality: {quality}"
    messagebox.showinfo("Report", report)

def read_excel(file_path):
    """Read an Excel file and return its content."""
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")
        return None

def read_pdf(file_path):
    """Read a PDF file and extract the text."""
    text = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + '\n'
    return text

def read_docx(file_path):
    """Read a Word document and extract the text."""
    doc = Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])

def validate_file_type(file_path):
    """Check if the file type is valid."""
    valid_extensions = ['.xlsx', '.pdf', '.docx']
    if not any(file_path.endswith(ext) for ext in valid_extensions):
        raise ValueError("Invalid file type selected.")

def secure_gpt3_prompt(content):
    """Ensure the GPT-3 prompt does not contain sensitive information."""
    # Implement any necessary checks to sanitize the content
    return content

def upload_file():
    """Handle the file upload process with added security checks."""
    global file_path, file_type
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("PDF Files", "*.pdf"), ("Word Files", "*.docx")])
        if file_path:
            validate_file_type(file_path)
            file_label.config(text=f"Uploaded File: {os.path.basename(file_path)}")
            enhance_button.pack()
    except ValueError as e:
        messagebox.showerror("Error", str(e))

def start_enhancement():
    """Start the file enhancement process with robust error handling."""
    try:
        if file_type == "Excel":
            df = read_excel(file_path)
            if df is not None:
                enhanced_df = enhance_excel_content(df)
                if enhanced_df is not None:
                    save_enhanced_excel(enhanced_df, file_path)
        elif file_type == "PDF":
            pdf_text = read_pdf(file_path)
            enhanced_text = gpt3_enhance(pdf_text)
            print(enhanced_text)  # Or handle the enhanced text appropriately
        elif file_type == "Word":
            doc_text = read_docx(file_path)
            enhanced_text = gpt3_enhance(doc_text)
            print(enhanced_text)  # Or handle the enhanced text appropriately
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main window
root = tk.Tk()
root.title("Test Plan Doctor")
root.geometry("800x600")

# Frame Setup
top_frame = tk.Frame(root)
top_frame.pack(pady=20)
middle_frame = tk.Frame(root)
middle_frame.pack(pady=20)
bottom_frame = tk.Frame(root)
bottom_frame.pack(pady=20)

# File Type Options
file_type_options = ["Excel", "PDF", "Word"]
file_type_var = tk.StringVar(value=file_type_options[0])
file_type_menu = tk.OptionMenu(top_frame, file_type_var, *file_type_options)
file_type_menu.grid(row=1, column=1, padx=10)

# Upload Button
upload_button = tk.Button(top_frame, text="Upload File", command=upload_file)
upload_button.grid(row=0, column=0, padx=10)

# File Label
file_label = tk.Label(top_frame, text="No file uploaded.")
file_label.grid(row=0, column=1, padx=10)

# Enhance Button
enhance_button = tk.Button(bottom_frame, text="Start Enhancement", command=start_enhancement)
enhance_button.grid(row=0, column=0, padx=10)

# Report Button
report_button = tk.Button(bottom_frame, text="Generate Report", command=lambda: generate_report(enhanced_content))
report_button.grid(row=1, column=0, padx=10)

# Start the GUI event loop
root.mainloop()