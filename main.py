import tkinter as tk
from tkinter import filedialog, messagebox
import openai
import os
import chardet  # chardet kütüphanesini ekliyoruz

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

def detect_encoding(file_path):
    """Detect the character encoding of the file."""
    with open(file_path, 'rb') as file:
        result = chardet.detect(file.read())
        return result['encoding']

def read_and_enhance_file(file_path):
    """Read the file and enhance its content using GPT-3."""
    try:
        # Detect the encoding of the file
        encoding = detect_encoding(file_path)
        # Read the file using the detected encoding
        with open(file_path, 'r', encoding=encoding) as file:
            content = file.read()

        # Enhance the content
        return gpt3_enhance(content)
    except UnicodeDecodeError as e:
        messagebox.showerror("Error", f"Error reading the file: {e}")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Unexpected error: {e}")
        return None

def upload_file():
    """Handle the file upload process."""
    global file_path
    file_path = filedialog.askopenfilename()
    if file_path:
        file_label.config(text="Uploaded File: " + os.path.basename(file_path))
        enhance_button.pack()

def start_enhancement():
    """Start the file enhancement process."""
    enhanced_content = read_and_enhance_file(file_path)
    if enhanced_content:
        # Logic to save enhanced content or display it in the GUI
        messagebox.showinfo("Enhancement Complete", "File has been enhanced.")
        # Example: Display enhanced content (or save it)
        print(enhanced_content)

# Create the main window
root = tk.Tk()
root.title("Test Plan Doctor")
root.geometry("800x600")

# Create and place the upload button
upload_button = tk.Button(root, text="Upload File", command=upload_file)
upload_button.pack()

# Label to display the name of the uploaded file
file_label = tk.Label(root, text="No file uploaded.")
file_label.pack()

# Button to start the enhancement process, hidden until a file is uploaded
enhance_button = tk.Button(root, text="Start to Enhance", command=start_enhancement)

# Start the GUI event loop
root.mainloop()
