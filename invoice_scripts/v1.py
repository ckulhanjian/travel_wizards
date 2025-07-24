import os
import pymupdf
import fitz
import regex
import shutil

def rename_file_in_directory(original_file, new_name, target_path, source_path):
    '''
    Rename file in target directory if it exists there
    '''
    os.chdir(target_path)
    if original_file in os.listdir():
        # Add .pdf extension if not present
        if not new_name.endswith('.pdf'):
            new_name += '.pdf'
        os.rename(original_file, new_name)
        print(f"Renamed {original_file} to {new_name}")
    else:
        print(f"File {original_file} not found in {target_path}")
    os.chdir(source_path)

def process_and_rename_files(source_path, target_path):
    if os.getcwd() != source_path:
        print("Current directory does not match source path!")
        return
    
    files = [f for f in os.listdir() if f.lower().endswith('.pdf')]
    print(f"Number of PDF files: {len(files)}")

    for file in files:
        agent, invoice, lastName = "", "", ""
        new_name = ""
        
        # Get the full path to the original file
        original_file_path = os.path.join(source_path, file)
    
        try:
            # Open the ORIGINAL file for text extraction (don't touch the copy)
            doc = fitz.open(original_file_path)
            if doc:
                print(f"{file} opened successfully for text extraction!")
            
            page = doc[0]  # Grab the first page
            text = page.get_text("text")  # Extract text
            doc.close()  # Important: close the document immediately after reading
            
            # Extract needed data using regex
            agent_match = regex.search(r'SALES PERSON:\s*([A-Z0-9]+)', text)
            invoice_match = regex.search(r'INVOICE NO\.\s+(?:ITIN)?(\d+)', text)
            lastName_match = regex.search(r'FOR:\s+([A-Z]+)(?=/)', text)
            
            if agent_match and invoice_match and lastName_match:
                agent = agent_match.group(1)
                invoice = invoice_match.group(1) 
                lastName = lastName_match.group(1)
                new_name = f"{agent[3:]}_{invoice}_{lastName}.pdf"
                print(f"Generated new name: {new_name}")
                rename_file_in_directory(file, new_name, target_path, source_path)
            else:
                print(f"Error: couldn't extract all required fields from {file}")
                if not agent_match:
                    print("  - Missing SALES PERSON")
                if not invoice_match:
                    print("  - Missing INVOICE NO")
                if not lastName_match:
                    print("  - Missing FOR field")
            
        except Exception as e:
            print(f"Error processing {file}: {e}")

def main():
    # Define source path
    source_path = '/Users/carakool/Downloads/TW'
    os.chdir(source_path)

    new_dir = "new_invoices"
    target_path = os.path.join(source_path, new_dir)

    # Create target directory if it doesn't exist
    if not os.path.isdir(target_path):
        os.mkdir(target_path)
        print(f"Created directory: {target_path}")

    print("Copying files to new directory...")
    pdf_files = [f for f in os.listdir() if f.lower().endswith('.pdf')]
    
    for file in pdf_files:
        src = os.path.join(source_path, file)
        dest = os.path.join(target_path, file)
        
        if os.path.isfile(src):
            # Use copy2 to preserve metadata and timestamps
            shutil.copy2(src, dest)
            print(f"Copied: {file}")

    print("File copying complete. Starting rename process...")
    process_and_rename_files(source_path, target_path)
    print("Process completed!")

if __name__ == "__main__":
    main()