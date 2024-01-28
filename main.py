import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openai
import os
import pandas as pd

# OpenAI API Key
openai.api_key = 'YOUR_API_KEY_HERE'

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
    """Generate a report based on the analysis."""
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

def enhance_excel_content(df):
    """Enhance the content of an Excel file using GPT-3."""
    try:
        for col in df.columns:
            for i in range(len(df)):
                enhanced_text = gpt3_enhance(str(df.at[i, col]))
                if enhanced_text:
                    df.at[i, col] = enhanced_text
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error during Excel content enhancement: {e}")
        return None

def save_enhanced_excel(df, file_path):
    """Save the enhanced content to a new Excel file."""
    try:
        enhanced_file_path = os.path.splitext(file_path)[0] + "_enhanced.xlsx"
        df.to_excel(enhanced_file_path, index=False)
        messagebox.showinfo("Success", f"Enhanced file saved as {enhanced_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving enhanced Excel file: {e}")

def upload_file():
    """Handle the file upload process."""
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_label.config(text="Uploaded File: " + os.path.basename(file_path))
        enhance_button.pack()

def start_enhancement():
    """Start the file enhancement process."""
    df = read_excel(file_path)
    if df is not None:
        enhanced_df = enhance_excel_content(df)
        if enhanced_df is not None:
            save_enhanced_excel(enhanced_df, file_path)

# Create the main window
root = tk.Tk()
root.title("Test Plan Doctor")
root.geometry("800x600")

# Create frames for the layout
top_frame = tk.Frame(root)
top_frame.pack(pady=20)
middle_frame = tk.Frame(root)
middle_frame.pack(pady=20)
bottom_frame = tk.Frame(root)
bottom_frame.pack(pady=20)

# Add language options
language_options = ["English", "Spanish", "French", "German", "Chinese"]
language_var = tk.StringVar(value=language_options[0])
language_menu = tk.OptionMenu(top_frame, language_var, *language_options)
language_menu.grid(row=1, column=1, padx=10)

# Upload button
upload_button = tk.Button(top_frame, text="Upload Excel File", command=upload_file)
upload_button.grid(row=0, column=0, padx=10)

# File label
file_label = tk.Label(top_frame, text="No file uploaded.")
file_label.grid(row=0, column=1, padx=10)

# Enhance button
enhance_button = tk.Button(bottom_frame, text="Start Enhancement", command=start_enhancement)
enhance_button.grid(row=0, column=0, padx=10)

# Report button
report_button = tk.Button(bottom_frame, text="Generate Report", command=lambda: generate_report(enhanced_content))
report_button.grid(row=2, column=0, padx=10)

root.mainloop()