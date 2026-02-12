import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
import pandas as pd
import threading

class ExcelConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to TXT Converter")
        self.root.geometry("600x400")
        
        # Variables
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        
        # UI Setup
        self.bg_color = "#f4f4f4"
        self.root.configure(bg=self.bg_color)
        self.setup_ui()
        
    def setup_ui(self):
        tk.Label(self.root, text="Excel to TXT Converter", font=("Arial", 18, "bold"), bg=self.bg_color).pack(pady=20)
        
        # Input Selection
        tk.Button(self.root, text="Select Input Folder (Excel Files)", command=self.browse_input).pack(pady=10)
        tk.Label(self.root, textvariable=self.input_folder, bg=self.bg_color, fg="blue").pack()

        # Output Selection
        tk.Button(self.root, text="Select Output Folder", command=self.browse_output).pack(pady=10)
        tk.Label(self.root, textvariable=self.output_folder, bg=self.bg_color, fg="blue").pack()

        # Progress
        self.status_label = tk.Label(self.root, text="Ready", bg=self.bg_color)
        self.status_label.pack(pady=20)

        # Start Button
        tk.Button(self.root, text="START CONVERSION", command=self.start_conversion, bg="#4CAF50", fg="white", font=("Arial", 12, "bold")).pack(pady=10)

    def browse_input(self):
        folder = filedialog.askdirectory()
        if folder: self.input_folder.set(folder)
    
    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder: self.output_folder.set(folder)

    def start_conversion(self):
        if not self.input_folder.get() or not self.output_folder.get():
            messagebox.showerror("Error", "Please select both folders!")
            return
        
        thread = threading.Thread(target=self.convert)
        thread.start()

    def convert(self):
        try:
            input_path = Path(self.input_folder.get())
            output_path = Path(self.output_folder.get())
            files = list(input_path.glob("*.xlsx")) + list(input_path.glob("*.xls"))
            
            for excel_file in files:
                self.status_label.config(text=f"Processing: {excel_file.name}")
                df = pd.read_excel(excel_file)
                txt_file = output_path / f"{excel_file.stem}.txt"
                
                # Data ko TXT mein save karna
                df.to_csv(txt_file, sep='\t', index=False)
            
            self.status_label.config(text="Success! Conversion Completed.")
            messagebox.showinfo("Success", "All files converted to TXT!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()
