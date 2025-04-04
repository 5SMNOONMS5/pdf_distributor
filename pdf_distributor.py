import os
import shutil
import pandas as pd
from pathlib import Path

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
        print(f"🔴 讀取Excel檔案時發生錯誤: {e}")
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
        
        # Copy each required PDF to student's folder
        for pdf in pdf_list:
            source_pdf = os.path.join(pdf_source_dir, pdf)
            if os.path.exists(source_pdf):
                dest_pdf = os.path.join(student_folder, pdf)
                shutil.copy2(source_pdf, dest_pdf)
                print(f"已複製 {pdf} 到 {student_name} 的資料夾")
            else:
                print(f"🔴 ------ 找不到學生 {student_name} 需要的 {pdf} 檔案 ")

def main():
    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define paths
    excel_path = os.path.join(current_dir, "students.xlsx")
    pdf_source_dir = os.path.join(current_dir, "PDFs")
    output_dir = os.path.join(current_dir, "Student_Folders")
    
    # Check if Excel file exists
    if not os.path.exists(excel_path):
        print("🔴 找不到 students.xlsx 檔案!")
        print("🔴 請在同一個資料夾中建立名為 'students.xlsx' 的Excel檔案，並包含以下欄位:")
        print("🔴 - Student Name (學生姓名)")
        print("🔴 - Required PDFs (需要的PDF檔案，以逗號分隔)")
        return
    
    # Check if PDFs directory exists
    if not os.path.exists(pdf_source_dir):
        print("🔴 錯誤: 找不到 PDFs 資料夾!")
        return
    
    # Create student folders and copy PDFs
    create_student_folders(excel_path, pdf_source_dir, output_dir)
    print("\nPDF檔案分配完成!")

if __name__ == "__main__":
    main() 