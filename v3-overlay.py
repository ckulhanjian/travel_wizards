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
        self.root.geometry("800x300")
        
        # Variables
        self.source_folder = tk.StringVar()
        self.overlay_path = tk.StringVar()
        self.lastPage = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Title
        title = tk.Label(self.root, text="Modify Invoices", font=("Arial", 16, "bold"))
        title.pack(pady=10)
        # Instructions
        instructions = tk.Label(self.root, text="Select the folder containing your PDF invoices.\n"
            "Renamed files will be saved in a 'renamed_invoices' subfolder.",wraplength=500)
        instructions.pack(pady=5)
        
        # Folder selection
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(pady=10, padx=20, fill="x")
        tk.Label(folder_frame, text="Folder:").pack(side="left")
        folder_entry = tk.Entry(folder_frame, textvariable=self.source_folder, width=40, state="readonly")
        folder_entry.pack(side="left", padx=5, fill="x", expand=True)
        browse_btn = tk.Button(folder_frame, text="Browse", command=self.browse_folder)
        browse_btn.pack(side="right")

        # Overlay PDF selector (no checkbox)
        overlay_frame = tk.Frame(self.root)
        overlay_frame.pack(pady=10, padx=20, fill="x")
        tk.Label(overlay_frame, text="Overlay PDF:").pack(side="left")
        overlay_entry = tk.Entry(overlay_frame, textvariable=self.overlay_path, width=40, state="readonly")
        overlay_entry.pack(side="left", padx=5, fill="x", expand=True)
        overlay_btn = tk.Button(overlay_frame, text="Select Overlay PDF", command=lambda:self.browse_file("Select Overlay PDF",self.overlay_path))
        overlay_btn.pack(side="right")
                
        # last page
        last_frame = tk.Frame(self.root)
        last_frame.pack(pady=10, padx=20, fill="x")
        tk.Label(last_frame, text="Last Page:").pack(side="left")
        lastPage_entry = tk.Entry(last_frame, textvariable=self.lastPage, width=40, state="readonly")
        lastPage_entry.pack(side="left", padx=5, fill="x", expand=True)
        last_btn = tk.Button(last_frame, text="Select Back Page", command=lambda:self.browse_file("Select Back Page",self.lastPage))
        last_btn.pack(side="right")

        # Process button
        self.process_btn = tk.Button(self.root, text="Process PDFs", 
                                    command=self.start_processing,
                                    bg="navy", fg="black", 
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
    
    def browse_file(self, text_title, path):
        file = filedialog.askopenfilename(
            title=f"{text_title}",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file:
            path.set(file)
            self.log(f"{text_title}: " + file)

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
        
        if not self.overlay_path.get():
            messagebox.showerror("Error", "Please select an overlay PDF or uncheck the overlay option!")
            return
            
        self.process_btn.config(state="disabled", text="Processing...")
        self.log_text.delete(1.0, tk.END)
        
        # Run in separate thread
        thread = threading.Thread(target=self.process_pdfs)
        thread.daemon = True
        thread.start()

    def apply_overlay(self, pdf_path, overlay_path, last_path):
        """Apply overlay PDF to all pages of the target PDF"""
        try:
            # Create a temporary file path
            temp_path = pdf_path.replace('.pdf', '_temp.pdf')
            
            # Open both PDFs
            background = fitz.open(pdf_path)
            overlay = fitz.open(overlay_path)
            lp = fitz.open(last_path)
            
            # Ensure both PDFs have at least one page
            if len(background) < 1:
                self.log(f"    Warning: Background PDF has no pages")
                background.close()
                overlay.close()
                return False
                
            if len(overlay) < 1:
                self.log(f"    Warning: Overlay PDF has no pages")
                background.close()
                overlay.close()
                return False
            
            # Process each page by inserting overlay content
            for page_num in range(len(background)):
                page = background[page_num]
                # Insert the overlay page on top of current page
                page.show_pdf_page(page.rect, overlay, 0)

            background.insert_pdf(lp, from_page=0, to_page=0)

            # Save to a new file (not incremental)
            background.save(temp_path)
            
            # Close documents
            background.close()
            overlay.close()
            lp.close()
            
            # Replace the original file
            if os.path.exists(temp_path):
                # Remove original and rename temp
                os.remove(pdf_path)
                os.rename(temp_path, pdf_path)
                return True
            else:
                self.log(f"    Warning: Temp file was not created")
                return False
            
        except Exception as e:
            self.log(f"    ✗ Overlay error: {str(e)}")
            # Clean up temp file if it exists
            temp_path = pdf_path.replace('.pdf', '_temp.pdf')
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass
            return False

    def process_pdfs(self):
        try:
            source_path = self.source_folder.get()
            target_path = os.path.join(source_path, "processed_invoices")
            
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
            # if self.use_overlay.get():
            self.log(f"Overlay mode enabled with: {os.path.basename(self.overlay_path.get())}")
            
            successful = 0
            failed = 0
            
            for i, file in enumerate(pdf_files, 1):
                self.log(f"\n[{i}/{len(pdf_files)}] Processing: {file}")
                
                try:
                    # Copy original file first
                    src = os.path.join(source_path, file)
                    dest = os.path.join(target_path, file)
                    shutil.copy2(src, dest)
                    
                    # Apply overlay if enabled
                    self.log(f"  Applying overlay...")
                    if not self.apply_overlay(dest, self.overlay_path.get(), self.lastPage.get()):
                        self.log(f"  ✗ Overlay failed, but continuing with rename...")
                    else:
                        self.log(f"  ✓ Overlay & last page applied")

                    # Extract text for renaming (from original file to avoid overlay text interference)
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
                        
                        # Rename the processed file
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
                overlay_msg = " with overlay applied"
                messagebox.showinfo("Complete", 
                                  f"Processing complete!\n"
                                  f"Successfully processed{overlay_msg}: {successful} files\n"
                                  f"Failed: {failed} files\n\n"
                                  f"Check the 'processed_invoices' folder in your selected directory.")
            
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