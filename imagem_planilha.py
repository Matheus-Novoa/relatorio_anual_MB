import os
from pathlib import Path
from pdf2image import convert_from_path
import win32com.client
import time

# Define directories
input_dir = "extracted_tables"
output_dir = "imagens"
temp_dir = "temp_pdf"

# Create output and temp directories if they don't exist
os.makedirs(output_dir, exist_ok=True)
os.makedirs(temp_dir, exist_ok=True)

# Initialize Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

try:
    # Process each Excel file in the input directory
    for excel_file in Path(input_dir).glob("*.xlsx"):
        try:
            # Full paths
            pdf_path = os.path.join(temp_dir, f"{excel_file.stem}.pdf")
            excel_path = str(excel_file.absolute())
            
            # Open workbook
            wb = excel.Workbooks.Open(excel_path)
            
            try:
                # Select and export notas sheet
                ws = wb.Worksheets("notas")
                ws.Select()
                
                # Convert to PDF - 0 is the constant for PDF format
                wb.ExportAsFixedFormat(0, pdf_path)
                
                # Ensure PDF is ready
                time.sleep(1)
                
                # Convert PDF to image
                images = convert_from_path(pdf_path)
                image_path = os.path.join(output_dir, f"{excel_file.stem}.png")
                images[0].save(image_path, 'PNG')
                
                print(f"Successfully converted {excel_file.name} to image")
                
            finally:
                wb.Close(False)  # Close without saving
                
        except Exception as e:
            print(f"Error processing {excel_file.name}: {str(e)}")
        finally:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
finally:
    # Make sure Excel is properly closed
    excel.Quit()
    
# Clean up temp directory
if os.path.exists(temp_dir) and not os.listdir(temp_dir):
    os.rmdir(temp_dir)