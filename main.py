import tkinter as tk
from tkinter import filedialog, messagebox
import openai
import os
import chardet
import pandas as pd

# OpenAI API Key
openai.api_key = 'YOUR_API_KEY_HERE'

def gpt3_enhance(content):
    """Enhance the content using GPT-3."""
    try:
        response = openai.Completion.create(engine="text-davinci-003", prompt=content, max_tokens=500)
        return response.choices[0].text.strip()
    except Exception as e:
        messagebox.showerror("Error", f"Error during GPT-3 enhancement: {e}")
        return None

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

# Create and place the upload button
upload_button = tk.Button(root, text="Upload Excel File", command=upload_file)
upload_button.pack()

# Label to display the name of the uploaded file
file_label = tk.Label(root, text="No file uploaded.")
file_label.pack()

# Button to start the enhancement process, hidden until a file is uploaded
enhance_button = tk.Button(root, text="Start Enhancement", command=start_enhancement)

# Start the GUI event loop
root.mainloop()