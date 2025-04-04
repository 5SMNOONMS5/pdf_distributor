import os
import shutil
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfMerger
import sys

def create_student_folders(excel_path, pdf_source_dir, output_dir):
    """
    Create folders for each student and copy their required PDFs.
    
    Args:
        excel_path (str): Path to the Excel file containing student data
        pdf_source_dir (str): Directory containing the PDF files
        output_dir (str): Directory where student folders will be created
    """
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Read the Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"讀取Excel檔案時發生錯誤: {e}")
        return
    
    # Process each student
    for index, row in df.iterrows():
        student_name = str(row['Student Name']).strip()
        required_pdfs = str(row['Required PDFs']).strip()
        
        # Create student folder
        student_folder = os.path.join(output_dir, student_name)
        os.makedirs(student_folder, exist_ok=True)
        
        # Split the required PDFs (assuming they're comma-separated)
        pdf_list = [pdf.strip() for pdf in required_pdfs.split(',')]
        
        # List to store successfully copied PDFs
        copied_pdfs = []
        
        # Copy each required PDF to student's folder
        for pdf in pdf_list:
            # Add .pdf suffix if not present
            if not pdf.lower().endswith('.pdf'):
                pdf = pdf + '.pdf'
            
            source_pdf = os.path.join(pdf_source_dir, pdf)
            if os.path.exists(source_pdf):
                dest_pdf = os.path.join(student_folder, pdf)
                shutil.copy2(source_pdf, dest_pdf)
                copied_pdfs.append(dest_pdf)
                print(f"已複製 {pdf} 到 {student_name} 的資料夾")
            else:
                print(f"警告: 找不到 {student_name} 需要的 {pdf} 檔案")
        
        # If we have successfully copied PDFs, merge them
        if copied_pdfs:
            try:
                merger = PdfMerger()
                for pdf in copied_pdfs:
                    try:
                        merger.append(pdf)
                    except Exception as e:
                        print(f"警告: 處理檔案 {os.path.basename(pdf)} 時發生錯誤，將跳過此檔案")
                        continue
                
                if len(merger.pages) > 0:  # Only create merged PDF if we have pages
                    # Create merged PDF filename
                    merged_pdf = os.path.join(student_folder, f"{student_name}.pdf")
                    merger.write(merged_pdf)
                    print(f"已合併 {len(merger.pages)} 個PDF檔案為 {student_name}.pdf")
                else:
                    print(f"警告: 無法為 {student_name} 建立合併PDF，因為沒有可用的PDF檔案")
                
                merger.close()
            except Exception as e:
                print(f"警告: 合併PDF檔案時發生錯誤: {e}")

def main():
    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define paths
    excel_path = os.path.join(current_dir, "students.xlsx")
    pdf_source_dir = os.path.join(current_dir, "PDFs")
    output_dir = os.path.join(current_dir, "Student_Folders")
    
    # Check if Excel file exists
    if not os.path.exists(excel_path):
        print("錯誤: 找不到 students.xlsx 檔案!")
        print("請在同一個資料夾中建立名為 'students.xlsx' 的Excel檔案，並包含以下欄位:")
        print("- Student Name (學生姓名)")
        print("- Required PDFs (需要的PDF檔案，以逗號分隔，可省略.pdf後綴)")
        return
    
    # Check if PDFs directory exists
    if not os.path.exists(pdf_source_dir):
        print("錯誤: 找不到 PDFs 資料夾!")
        return
    
    # Create student folders and copy PDFs
    create_student_folders(excel_path, pdf_source_dir, output_dir)
    print("\nPDF檔案分配完成!")

if __name__ == "__main__":
    main() 