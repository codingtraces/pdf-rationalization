# One to Many PDF Comparer Tool Setup Guide

This documentation will help you set up and run the One to Many PDF Comparer Tool, including installing necessary Python packages, configuring the environment, and packaging the script into an executable file.

## Prerequisites

- Python 3.8 or higher installed on your system.
- A working internet connection for installing dependencies.
- Basic knowledge of Python and command-line usage.

## Step 1: Clone or Download the Code

1. Download the script files from the repository or provided location.
2. Extract the files to a desired folder on your computer.

## Step 2: Install Python Dependencies

Ensure you have Python and `pip` (Python package manager) installed. You can verify the installation by running:

```sh
python --version
pip --version
```

### Installing Required Packages

Navigate to the folder containing the script files and install the necessary packages by running the following command:

```sh
pip install -r requirements.txt
```

Alternatively, you can manually install each package using the following command:

```sh
pip install tkinter PyMuPDF openpyxl pandas scikit-learn Pillow
```

The list of required libraries includes:

- `tkinter`: For creating the graphical user interface (GUI).
- `PyMuPDF` (imported as `fitz`): For PDF text extraction.
- `openpyxl`: For generating Excel reports.
- `pandas`: For data manipulation and handling.
- `scikit-learn`: For text similarity calculations using `TfidfVectorizer`.
- `Pillow`: For handling and displaying images in the GUI.

### Requirements File

Create a `requirements.txt` file in the project directory to facilitate easy package installation:

```
tkinter
PyMuPDF
openpyxl
pandas
scikit-learn
Pillow
```

## Step 3: Run the Script

To run the Python script, navigate to the folder where it is saved and run:

```sh
python one_to_many_pdf_comparer.py
```

This will open the One to Many PDF Comparer Tool's graphical interface.

## Step 4: Using the Tool

1. **Single PDF Folder**: Select the folder containing the single PDF file you want to compare against others.
2. **All PDFs Folder**: Select the folder containing multiple PDF files for comparison.
3. **Output Folder**: Select the folder where the generated reports will be saved.
4. **Generate Reports**:
   - Click **Generate Excel Report** to generate an Excel file with the comparison results.
   - Click **Generate HTML Report** to generate an HTML file with the comparison results.
   - Click **Generate Percentage Match** to generate a report with similarity percentages.

## Step 5: Packaging the Script into an Executable

To make the tool more user-friendly, you can package it into an executable file (.exe) using `PyInstaller`.

### Installing PyInstaller

First, install PyInstaller:

```sh
pip install pyinstaller
```

### Creating an Executable

Run the following command to package the script into an executable:

```sh
pyinstaller --onefile --windowed --add-data "dev-logo.png;." one_to_many_pdf_comparer.py
```

Explanation of flags:

- `--onefile`: Creates a single executable file.
- `--windowed`: Suppresses the command prompt window when running the executable.
- `--add-data "dev-logo.png;."`: Includes the logo file (`dev-logo.png`) in the executable.

After running this command, the executable can be found in the `dist` folder.

### Notes on Packaging

- If the logo image (`dev-logo.png`) is not correctly displayed, ensure it is correctly included in the build using the `--add-data` flag.
- The executable might be large due to the bundled libraries. Consider using `UPX` to compress the executable size if needed.

## Troubleshooting

### Common Errors

1. **Missing DLL Files**: When running the executable, if you encounter missing DLL files, ensure all dependencies are installed properly and try re-running the `PyInstaller` command.

2. **GUI Not Opening**: Ensure you have used the `--windowed` flag to suppress the command prompt. If there are any errors, run the executable from the terminal to see detailed logs.

## Summary

This guide provided instructions for setting up the One to Many PDF Comparer Tool, including installing dependencies, running the Python script, and creating an executable file. With this tool, you can compare the contents of one PDF against multiple PDFs and generate detailed reports in Excel, HTML, or as similarity percentages.

