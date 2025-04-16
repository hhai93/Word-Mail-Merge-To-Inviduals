# 📜 Word Mail Merge To Individuals

![GitHub](https://img.shields.io/badge/license-MIT-blue.svg) ![GitHub last commit](https://img.shields.io/github/last-commit/hhai93/Word-Mail-Merge-to-Individuals)

This VBA script automates the creation of individual files from a Word document using **Mail Merge** with data from an Excel file. Each row in the Excel file generates a separate file (PDF, DOCX, RTF, etc.), named based on a specified column.

---

## ✨ Features
- 🔗 Links a Word document to an Excel data source via Mail Merge.
- 📄 Generates one file per row of data.
- 🖋️ Names each file using a value from a chosen Excel column.
- 🎨 Supports multiple output formats (PDF, DOCX, RTF, HTML, etc.).

## 📋 Prerequisites
- 🖥️ Microsoft Word (2010 or later) with VBA support enabled.
- 📊 An Excel file (`.xlsx` or `.xls`) containing your data.
- 📝 A Word document (`.docx`) with Mail Merge fields configured.

---

## 🚀 How to Use

### 1. Prepare Your Files
- 📝 Create a Word document with Mail Merge fields (e.g., `{Name}`, `{Address}`). Connect it to your Excel file using the **Mailings** tab.
- 📊 Ensure your Excel file has column headers matching the Mail Merge fields (e.g., "Name", "Address").

### 2. Add the VBA Script
- Open your Word document.
- Press `Alt + F11` to launch the VBA editor.
- Go to **Insert** > **Module** and paste the code from [`SaveAsSeparateFiles.vba`](SaveAsSeparateFiles.vba).
- ✏️ Customize the script:
  - Replace `"Name"` with the column name for naming files (e.g., `"ID"`, `"CustomerName"`).
  - Update `"C:\YourFolderPath\"` to your desired output folder (e.g., `"C:\Users\YourName\Documents\Output\"`). Ensure the folder exists.
  - Change the `fileFormat` variable to your desired format:
    - `wdFormatPDF` for `.pdf` (default).
    - `wdFormatDocumentDefault` for `.docx`.
    - `wdFormatRTF` for `.rtf`.
    - `wdFormatHTML` for `.html`.

### 3. Run the Script
- Press `F5` or go to **Run** > **Run Sub/UserForm** to execute.
- 🎉 The script will generate one file per row in your Excel data!

---

## 🛠️ Code Explanation
- **`dataField`**: Defines the Excel column used for naming output files.
- **`fileFormat`**: Sets the output format (e.g., PDF, DOCX). Modify this in the script to change formats.
- **`GetFileExtension`**: Helper function to match file extensions with formats.
- 🔄 The script loops through each Mail Merge record, creates a new document, and saves it in the chosen format.

---

## ⚠️ Notes
- 📂 Keep your Excel and Word files open while running the script.
- 🔍 If column values are duplicated, files may overwrite each other. Modify the script to append a unique identifier (e.g., row number) if needed.
- 💾 Back up your files before running to avoid data loss.
- 🖌️ Not all formats (e.g., HTML) may preserve formatting perfectly—test your output.

## 🌟 Example
### Excel Data
| Name   | Address      | Phone    |
|--------|--------------|----------|
| John   | 123 Main St  | 555-1234 |
| Alice  | 456 Oak Ave  | 555-5678 |

### Output (with `fileFormat = wdFormatPDF`)
- `C:\YourFolderPath\John.pdf`
- `C:\YourFolderPath\Alice.pdf`
