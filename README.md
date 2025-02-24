# metadata-remover

A simple tool to remove metadata from PDF, DOCX, XLSX, XLS, PPTX, PPT, and DOC files. It creates a cleaned copy while preserving the folder structure.

## Features

- **Removes metadata** like Author, Company, Last Modified By, Revision, Custom Properties, and more.
- **Preserves original structure** without breaking files.
- **Simple GUI** using PyQt5 for folder selection.
- **Processes all files in a folder** and saves cleaned copies in a `<original_folder_name>_no_metadata` folder.

Metadata-Remover is well tested on PDF and docx files. Other file types are to be tested more in the future. So no guarantees there :)

## Installation

```sh
git clone github.com/balpars/metadata-remover
cd metadata-remover
pip install -r requirements.txt
```

Alternatively, just use the packaged executabe. You can find it in releases section.

## Usage
Run the script (or exe file if you downloaded the zip from releases):
```sh
python metadata_remover_gui.py
```

Select a folder containing files.
Wait for processing to complete.
Cleaned files will be saved in <selected_folder>_no_metadata.
