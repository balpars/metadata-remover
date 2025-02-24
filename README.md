# metadata-remover

A simple tool to remove metadata from PDF, DOCX, XLSX, XLS, PPTX, PPT, and DOC files. It creates a cleaned copy while preserving the folder structure.

## Features

- **Removes metadata** like Author, Company, Last Modified By, Revision, Custom Properties, and more.
- **Preserves original structure** without breaking files.
- **Simple GUI** using PyQt5 for folder selection.
- **Processes all files in a folder** and saves cleaned copies in a `_no_metadata` folder.

## Installation

```sh
pip install PyQt5 PyPDF2 python-docx openpyxl python-pptx lxml olefile
```

## Usage
Run the script:
```sh
python metadata_remover_gui.py
```

Select a folder containing files.
Wait for processing to complete.
Cleaned files will be saved in <selected_folder>_no_metadata.
