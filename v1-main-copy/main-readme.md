# Content Rationalizer Tool Documentation

## Overview
The Content Rationalizer Tool is a simple Python-based application that helps you compare paragraphs from multiple PDF files. It shows which paragraphs are common or different and creates reports in Excel format to make it easier to understand the comparison.

## Features
- **Extract Paragraphs from PDFs**: Reads and splits the text from selected PDFs into paragraphs.
- **Rationalize PDFs**: Creates an Excel report showing which paragraphs are present in each PDF.
- **Percentage Match Comparison**: Compares the similarity between paragraphs from different PDFs and generates a similarity score.

## How It Works
The tool uses Python and several key libraries:
- **PyMuPDF (`fitz`)** to read content from PDF files.
- **Tkinter** to create a simple GUI.
- **Openpyxl** to generate Excel reports.
- **Difflib (`SequenceMatcher`)** to calculate the similarity between paragraphs.

### Simple Steps to Use the Tool
1. **Select Input Folder**: Click the "Browse" button to choose the folder containing PDF files you want to compare.
2. **Select Output Folder**: Click the "Browse" button to choose where you want to save the Excel report.
3. **Rationalize PDFs**: Click the "Rationalise" button to compare paragraphs across all PDFs and generate a report.
4. **Percentage Match**: Click the "Percentage Match" button to create a similarity score report for the paragraphs.

### How the Comparison Works
#### Extract Paragraphs
- The tool reads each PDF file and splits the content into paragraphs.
- It removes any empty paragraphs to clean up the data.

#### Rationalize PDFs (`pdfcompare`)
- The tool collects all unique paragraphs from the PDFs and creates a list.
- It then checks each paragraph to see if it exists in each PDF and marks it with `1` (present) or `0` (not present).
- An Excel report (`Pdf Results Sheet`) is generated with a table showing which paragraphs are in each PDF.

#### Percentage Match (`percentage_match`)
- The tool sorts the words in each paragraph and compares them to see how similar they are.
- It calculates a percentage score for each pair of paragraphs and creates a similarity matrix in the Excel report (`Percentage Match Sheet`).

### Example of How the Reports Work
The tool generates two main Excel sheets:

1. **Pdf Results Sheet**:
   - This sheet shows which paragraphs are present in each PDF.
   - **Example Table**:
     
     | PDF                   | Paragraph 1 | Paragraph 2 | Paragraph 3 |
     |-----------------------|-------------|-------------|-------------|
     | common_text_sample.pdf| 1           | 0           | 1           |
     | unique_text_sample.pdf| 0           | 1           | 1           |
     
     In this table, `1` means the paragraph is present in the PDF, and `0` means it is not. For example, `common_text_sample.pdf` has `Paragraph 1` and `Paragraph 3`, but not `Paragraph 2`.

2. **Percentage Match Sheet**:
   - This sheet shows how similar paragraphs are to each other.
   - **Example Table**:
     
     | Paragraph     | Paragraph 1 | Paragraph 2 | Paragraph 3 |
     |---------------|-------------|-------------|-------------|
     | Paragraph 1   | 100%        | 45%         | 30%         |
     | Paragraph 2   | 45%         | 100%        | 60%         |
     
     In this table, `Paragraph 1` is `100%` similar to itself, `45%` similar to `Paragraph 2`, and `30%` similar to `Paragraph 3`.

### How to Use the Tool Step-by-Step
1. **Launch the Tool**:
   - Run the Python script (`main.py`) to open the GUI.
2. **Select Input Folder**:
   - Click "Browse" next to "Input Folder" to choose the folder with the PDFs you want to compare.
3. **Select Output Folder**:
   - Click "Browse" next to "Output Folder" to choose where you want the report to be saved.
4. **Compare PDFs**:
   - **Rationalise**: Click "Rationalise" to generate the report showing which paragraphs are in each PDF.
   - **Percentage Match**: Click "Percentage Match" to generate the similarity score report.

### Simple Workflow Example
Suppose you have three PDF files: `file1.pdf`, `file2.pdf`, and `file3.pdf`.

1. **Rationalize PDFs**:
   - The tool extracts paragraphs from all three files.
   - Let's say there are four unique paragraphs in total.
   - The `Pdf Results Sheet` will look like this:
     
     | PDF      | Paragraph 1 | Paragraph 2 | Paragraph 3 | Paragraph 4 |
     |----------|-------------|-------------|-------------|-------------|
     | file1.pdf| 1           | 1           | 0           | 1           |
     | file2.pdf| 0           | 1           | 1           | 1           |
     | file3.pdf| 1           | 0           | 1           | 1           |

2. **Percentage Match**:
   - The tool calculates how similar each paragraph is to the others.
   - The `Percentage Match Sheet` will look like this:
     
     | Paragraph     | Paragraph 1 | Paragraph 2 | Paragraph 3 | Paragraph 4 |
     |---------------|-------------|-------------|-------------|-------------|
     | Paragraph 1   | 100%        | 30%         | 50%         | 20%         |
     | Paragraph 2   | 30%         | 100%        | 40%         | 25%         |

### Troubleshooting Tips
- **Missing Report**: Make sure your input folder has valid PDF files and your output folder is writable.
- **Text Extraction Error**: If you see an error like "Error extracting text," the PDF might not contain text data (e.g., it might be an image-based PDF).

## System Requirements
- **Python 3.x**
- Required Libraries: `fitz` (PyMuPDF), `tkinter`, `openpyxl`, `PIL` (Pillow)
- Make sure all libraries are installed in your Python environment.

## Future Enhancements
- Support for extracting text from image-based PDFs using OCR.
- Improved user interface for a better experience.
- Add logging for easier troubleshooting.

This guide should help you understand how to use the Content Rationalizer Tool and interpret the results easily. If you need further assistance, feel free to ask!

