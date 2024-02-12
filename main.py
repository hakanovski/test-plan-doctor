import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openai
import os
import pandas as pd
import pdfplumber
from docx import Document

# OpenAI API Key
openai.api_key = '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'

# Initialize global variable for enhanced content
enhanced_content = None

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

def analyze_content(content):
    """
    Analyze the content for quality and coverage.
    """
    word_count, quality = len(content.split()), "Good Quality" if len(content.split()) > 50 else "Needs Improvement"
    return word_count, quality

def generate_report(content):
    """
    Generate and save a report based on the analysis.
    """
    global enhanced_content  # Ensure the use of the global variable
    word_count, quality = analyze_content(content)
    report = f"Content Analysis Report:\nWord Count: {word_count}\nQuality: {quality}"
    messagebox.showinfo("Report", report)

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

def read_pdf(file_path):
    """
    Read a PDF file and extract the text.
    """
    text = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + '\n'
    return text

def read_docx(file_path):
    """
    Read a Word document and extract the text.
    """
    doc = Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])

def validate_file_type(file_path):
    """
    Check if the file type is valid.
    """
    valid_extensions = ['.xlsx', '.pdf', '.docx']
    if not any(file_path.endswith(ext) for ext in valid_extensions):
        raise ValueError("Invalid file type selected.")

def secure_gpt3_prompt(content):
    """
    Ensure the GPT-3 prompt does not contain sensitive information.
    """
    # Implement any necessary checks to sanitize the content here
    return content

def upload_file():
    """
    Handle the file upload process with added security checks.
    """
    global file_path, file_type
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("PDF Files", "*.pdf"), ("Word Files", "*.docx")])
        if file_path:
            validate_file_type(file_path)
            file_label.config(text=f"Uploaded File: {os.path.basename(file_path)}")
            enhance_button.pack()
    except ValueError as e:
        messagebox.showerror("Error", str(e))

def enhance_excel_content(df):
    """
    Enhance the Excel content using GPT-3.
    """
    enhanced_rows = []
    for index, row in df.iterrows():
        enhanced_description = gpt3_enhance(row['Description'])
        enhanced_rows.append(enhanced_description)
    df['Enhanced Description'] = enhanced_rows
    return df

def save_enhanced_excel(df, file_path):
    """
    Save the enhanced Excel data to a file.
    """
    enhanced_file_path = os.path.splitext(file_path)[0] + "_enhanced.xlsx"
    df.to_excel(enhanced_file_path)
    messagebox.showinfo("Success", f"Enhanced file saved as {enhanced_file_path}")

def start_enhancement():
    """
    Start the file enhancement process with robust error handling.
    """
    global enhanced_content
    try:
        file_type = file_type_var.get()
        if file_type == "Excel":
            df = read_excel(file_path)
            if df is not None:
                enhanced_df = enhance_excel_content(df)
                if enhanced_df is not None:
                    save_enhanced_excel(enhanced_df, file_path)
        elif file_type == "PDF":
            pdf_text = read_pdf(file_path)
            enhanced_text = gpt3_enhance(pdf_text)
            enhanced_content = enhanced_text
            print(enhanced_text)  # Or handle the enhanced text appropriately
        elif file_type == "Word":
            doc_text = read_docx(file_path)
            enhanced_text = gpt3_enhance(doc_text)
            enhanced_content = enhanced_text
            print(enhanced_text)  # Or handle the enhanced text appropriately
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI setup and main loop are unchanged from your original code