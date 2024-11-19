# PDF Comparer Tool - Setup, Installation, and Usage Guide

## Table of Contents
1. [Introduction](#introduction)
2. [Requirements](#requirements)
3. [Installation](#installation)
4. [Running the Tool](#running-the-tool)
5. [Usage](#usage)
6. [Packaging as an Executable](#packaging-as-an-executable)
7. [Troubleshooting](#troubleshooting)

## Introduction
The PDF Comparer Tool is a graphical user interface (GUI) application designed to perform three types of PDF comparisons:
- **1-to-N Comparison**: Compare a single PDF against multiple PDFs.
- **All PDFs Comparison**: Compare all PDFs within a single folder.
- **N-to-M Comparison**: Compare all PDFs in one folder (Folder 1) against all PDFs in another folder (Folder 2).

The application provides both an HTML and an Excel report for each comparison, highlighting the similarity percentages between the text blocks of the PDF documents.

## Requirements
To set up and use the PDF Comparer Tool, ensure that you have the following prerequisites installed on your system:

- Python 3.8 or higher
- Tkinter (for the GUI)
- PyMuPDF (fitz)
- scikit-learn
- Pandas
- openpyxl
- Other Python packages: re, concurrent.futures

## Installation
Follow these steps to set up the project on your local environment:

1. **Clone the Repository**
   ```
   git clone <repository-url>
   cd <repository-folder>
   ```

2. **Create a Virtual Environment**
   Create a virtual environment to isolate the project dependencies:
   ```
   python -m venv .venv
   ```

3. **Activate the Virtual Environment**
   - On Windows:
     ```
     .venv\Scripts\activate
     ```
   - On macOS/Linux:
     ```
     source .venv/bin/activate
     ```

4. **Install Dependencies**
   Install all required dependencies:
   ```
   pip install -r requirements.txt
   ```
   Ensure that your `requirements.txt` contains the following:
   ```
   PyMuPDF
   scikit-learn
   pandas
   openpyxl
   ```

## Running the Tool
To run the PDF Comparer Tool, follow these steps:

1. **Navigate to the Project Directory**
   ```
   cd <repository-folder>
   ```

2. **Run the Python Script**
   ```
   python main.py
   ```

   This will start the Tkinter-based GUI, allowing you to perform the PDF comparisons.

## Usage

### 1. Selecting Comparison Type
- Upon starting the application, you will be presented with a drop-down menu to select the comparison type:
  - **1-to-N**: Compare a single PDF to multiple PDFs.
  - **All PDFs**: Compare all PDFs in a folder.
  - **N-to-M**: Compare PDFs in Folder 1 to PDFs in Folder 2.

### 2. Selecting Input and Output Folders
- Depending on the comparison type, different folders need to be selected:
  - **Single PDF Folder**: The folder containing the PDF for comparison.
  - **All PDFs Folder**: The folder containing multiple PDFs for comparison.
  - **Folder 1** and **Folder 2**: For N-to-M comparison, select the two folders to compare against each other.
  - **Output Folder**: Select the folder where the generated reports will be saved.

### 3. Adding a Logo
- The application also includes an option to display a logo image. The logo file (`dev-logo.png`) should be placed in the same directory as the script to ensure it is loaded correctly.

### 4. Running the Comparison
- After selecting the folders, click the "Compare PDFs" button.
- Once the comparison is complete, HTML and Excel reports will be generated and saved in the output folder.

## Packaging as an Executable
If you want to package the application as a standalone executable, you can use `PyInstaller`.

1. **Install PyInstaller**
   ```
   pip install pyinstaller
   ```

2. **Generate Executable**
   Run the following command to create a single `.exe` file for the project:
   ```
   pyinstaller --onefile --noconsole --add-data "dev-logo.png;." main.py
   ```
   - `--onefile`: Package everything into a single executable file.
   - `--noconsole`: Do not open a console window when running the executable.
   - `--add-data`: Include any additional files (like images) required by the application.

3. **Locate the Executable**
   After running the above command, the executable can be found in the `dist` folder.

## Troubleshooting
### Common Issues and Solutions

1. **ModuleNotFoundError**: If you encounter an error like `ModuleNotFoundError: No module named 'sklearn'`, ensure that all dependencies are installed in the virtual environment. Activate the environment and run:
   ```
   pip install -r requirements.txt
   ```

2. **Encrypted PDF Files**: If an encrypted PDF is encountered, it will be skipped. The error will be logged in `error_log.txt`.

3. **Large Number of Files**: If you're processing a large number of files (e.g., 50,000 PDFs), the tool has exception handling to prevent crashes. However, ensure that you have sufficient memory and processing power to handle large datasets.

4. **Error Logs**: Errors during processing are logged in `error_log.txt`. Review this file to diagnose any issues that occur during PDF analysis or comparison.

### Tips for Optimal Performance
- Ensure that your system has adequate resources (RAM and CPU) to process large numbers of PDFs.
- If possible, split large batches of PDFs into smaller sets to minimize processing time.

## Conclusion
The PDF Comparer Tool is a robust application for comparing PDFs, providing both graphical and report-based insights into the similarities between different documents. With the option to package it as an executable, it can easily be distributed to users who do not have Python installed.

Feel free to reach out if you need additional features or face any issues during setup and usage.

