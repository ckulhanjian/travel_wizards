#!/usr/bin/env python3
"""
PDF Invoice Renamer - Standalone Version
Just double-click to run!
"""

import os
import sys
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

try:
    import fitz  # PyMuPDF
    import regex
except ImportError as e:
    # Show error if dependencies are missing
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Missing Dependencies", 
                        f"Missing required library: {e}\n\n"
                        "Please install:\n"
                        "pip install PyMuPDF regex")
    sys.exit(1)

class PDFRenamerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Invoice Renamer")
        self.root.geometry("800x500")
        
        # Variables
        self.source_folder = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Title
        title = tk.Label(self.root, text="PDF Invoice Renamer", 
                        font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        # Instructions
        instructions = tk.Label(self.root, 
                               text="Select the folder containing your PDF invoices.\n"
                                   "Renamed files will be saved in a 'renamed_invoices' subfolder.",
                               wraplength=500)
        instructions.pack(pady=5)
        
        # Folder selection
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(folder_frame, text="Folder:").pack(side="left")
        
        folder_entry = tk.Entry(folder_frame, textvariable=self.source_folder, 
                               width=50, state="readonly")
        folder_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        browse_btn = tk.Button(folder_frame, text="Browse", 
                              command=self.browse_folder)
        browse_btn.pack(side="right")
        
        # Process button
        self.process_btn = tk.Button(self.root, text="Process PDFs", 
                                    command=self.start_processing,
                                    bg="#FFCCFF", fg="black", 
                                    font=("Arial", 12, "bold"),
                                    height=2)
        self.process_btn.pack(pady=20)
        
        # Progress/Log area
        log_label = tk.Label(self.root, text="Progress Log:")
        log_label.pack(anchor="w", padx=20)
        
        self.log_text = scrolledtext.ScrolledText(self.root, height=15, width=70)
        self.log_text.pack(pady=5, padx=20, fill="both", expand=True)
        
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDF invoices")
        if folder:
            self.source_folder.set(folder)
            self.log("Selected folder: " + folder)
            
    def log(self, message):
        """Add message to log area"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def start_processing(self):
        """Start processing in a separate thread to keep GUI responsive"""
        if not self.source_folder.get():
            messagebox.showerror("Error", "Please select a folder first!")
            return
            
        self.process_btn.config(state="disabled", text="Processing...")
        self.log_text.delete(1.0, tk.END)
        
        # Run in separate thread
        thread = threading.Thread(target=self.process_pdfs)
        thread.daemon = True
        thread.start()
        
    def process_pdfs(self):
        try:
            source_path = self.source_folder.get()
            target_path = os.path.join(source_path, "renamed_invoices")
            
            # Create target directory
            if not os.path.exists(target_path):
                os.makedirs(target_path)
                self.log(f"Created directory: {target_path}")
            
            # Get PDF files
            pdf_files = [f for f in os.listdir(source_path) 
                        if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(source_path, f))]
            
            if not pdf_files:
                self.log("No PDF files found in the selected folder!")
                self.process_btn.config(state="normal", text="Process PDFs")
                return
                
            self.log(f"Found {len(pdf_files)} PDF files")
            
            successful = 0
            failed = 0
            
            for i, file in enumerate(pdf_files, 1):
                self.log(f"\n[{i}/{len(pdf_files)}] Processing: {file}")
                
                try:
                    # Copy original file first
                    src = os.path.join(source_path, file)
                    dest = os.path.join(target_path, file)
                    shutil.copy2(src, dest)
                    
                    # Extract text for renaming
                    doc = fitz.open(src)
                    page = doc[0]
                    text = page.get_text("text")
                    doc.close()
                    
                    # Extract data using regex
                    agent_match = regex.search(r'SALES PERSON:\s*([A-Z0-9]+)', text)
                    invoice_match = regex.search(r'INVOICE NO\.\s+(?:ITIN)?(\d+)', text)
                    lastName_match = regex.search(r'FOR:\s+([A-Z]+)(?=/)', text)
                    
                    if agent_match and invoice_match and lastName_match:
                        agent = agent_match.group(1)
                        invoice = invoice_match.group(1)
                        lastName = lastName_match.group(1)
                        new_name = f"{agent[3:]}_{invoice}_{lastName}.pdf"
                        
                        # Rename the copied file
                        new_path = os.path.join(target_path, new_name)
                        os.rename(dest, new_path)
                        
                        self.log(f"  ✓ Renamed to: {new_name}")
                        successful += 1
                    else:
                        self.log(f"  ✗ Could not extract required fields")
                        failed += 1
                        
                except Exception as e:
                    self.log(f"  ✗ Error: {str(e)}")
                    failed += 1
            
            # Summary
            self.log(f"\n{'='*50}")
            self.log(f"SUMMARY:")
            self.log(f"Successfully processed: {successful}")
            self.log(f"Failed: {failed}")
            self.log(f"Output folder: {target_path}")
            
            if successful > 0:
                messagebox.showinfo("Complete", 
                                  f"Processing complete!\n"
                                  f"Successfully renamed: {successful} files\n"
                                  f"Failed: {failed} files\n\n"
                                  f"Check the 'renamed_invoices' folder in your selected directory.")
            
        except Exception as e:
            self.log(f"Fatal error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        finally:
            self.process_btn.config(state="normal", text="Process PDFs")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = PDFRenamerGUI()
    app.run()