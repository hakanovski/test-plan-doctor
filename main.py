import tkinter as tk
from tkinter import filedialog, messagebox

# Uygulama penceresini oluşturma
app = tk.Tk()
app.title("Test Plan Doctor")

# Giriş ekranı için metin ve buton
welcome_label = tk.Label(app, text="Welcome to Test Plan Doctor!\nThis tool helps you enhance your test plans.")
welcome_label.pack()

def open_main_interface():
    # Giriş ekranını kaldır
    welcome_label.destroy()
    start_button.destroy()
    
    # Ana arayüzün içeriği
    upload_label = tk.Label(app, text="Upload your test plan in Word, PDF, Excel, or Google Docs format.")
    upload_label.pack()

    def upload_file():
        file_path = filedialog.askopenfilename()
        # Dosya yükleme işlemleri burada yapılacak
        messagebox.showinfo("Uploaded", f"File uploaded: {file_path}")

    upload_button = tk.Button(app, text="Upload File", command=upload_file)
    upload_button.pack()

start_button = tk.Button(app, text="Start", command=open_main_interface)
start_button.pack()

# Uygulamayı çalıştır
app.mainloop()

