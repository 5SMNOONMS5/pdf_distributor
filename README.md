# PDF Distributor

This application helps distribute PDF files to students based on their requirements listed in an Excel file.

## Virtual environment

Avoid conflicts between different projects use Python virtual environment

In mac 
```
source venv/bin/activate
```

In Windows
```
venv\Scripts\activate
```

## Requirements

- Python 3.7 or higher
- Required Python packages (install using `pip install -r requirements.txt`):
  - pandas        
  - openpyxl
  - PyPDF2==3.0.1 

## Setup

1. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

2. Prepare your files:
   - Place all PDF files in the `PDFs` folder
   - Create an Excel file named `students.xlsx` with the following columns:
     - `Student Name`: Name of the student
     - `Required PDFs`: Comma-separated list of PDF filenames required by the student

## Excel File Format

Example of how your `students.xlsx` should look:

| Student Name | Required PDFs |
|-------------|---------------|
| A           | A國1.pdf, B數1.pdf |
| B           | C英2.pdf, D社1.pdf |
| C           | E自2.pdf, A國12.pdf |

## Usage

1. Make sure your PDF files are in the `PDFs` folder
2. Create your `students.xlsx` file in the same directory as the script
3. Run the script:
   ```
   python3 pdf_distributor.py
   ```

The script will:
1. Create a `Student_Folders` directory
2. Create a subfolder for each student
3. Copy the required PDFs to each student's folder

## Output

The script will create a `Student_Folders` directory containing:
- A folder for each student
- The required PDFs copied to each student's folder

## Error Handling

The script will:
- Check if the Excel file exists
- Check if the PDFs directory exists
- Verify if each required PDF exists
- Print warnings for missing PDFs
- Create folders automatically if they don't exist 