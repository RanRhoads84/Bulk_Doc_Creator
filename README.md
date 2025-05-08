# Document Creator Bulk

A Python script to quickly generate multiple files of various types (Excel, Word, PowerPoint, CSV, Markdown, YAML) based on a user-defined naming convention and quantity.

## Features

- **Bulk file creation**: Generate any number of files with sequential numbering.
- **Supports multiple formats**: Excel (`.xlsx`), Word (`.docx`), PowerPoint (`.pptx`), CSV (`.csv`), Markdown (`.md`), YAML (`.yaml`).
- **Custom naming conventions**: Insert a number anywhere in your filename pattern.
- **Organized output**: Files are grouped by type in suborders under `Created-Files/`.
- **Colorful prompts and errors**: Uses [colorama](https://pypi.org/project/colorama/) for improved terminal UX.
- **Input validation**: Prevents invalid filenames and ensures positive integer input.

## Requirements

- Python 3.7+
- [colorama](https://pypi.org/project/colorama/)
- [openpyxl](https://pypi.org/project/openpyxl/)
- [python-docx](https://pypi.org/project/python-docx/)
- [python-pptx](https://pypi.org/project/python-pptx/)
- [PyYAML](https://pypi.org/project/PyYAML/)

Install dependencies with:

```sh
pip install colorama openpyxl python-docx python-pptx PyYAML
```

## Usage

Run the script:

```sh
python document_creator_bulk.py
```

### Example Workflow

1. **Enter the naming convention**  
   Example: `Report_0_Draft`  
   The `0` will be replaced with sequential numbers.

2. **Select file type**  
   Choose from the menu (e.g., Markdown, Excel, etc.).

3. **Enter the number of copies**  
   Example: `5`  
   This will create files like `Report_1_Draft.md`, `Report_2_Draft.md`, ..., `Report_5_Draft.md`.

4. **Check your files**  
   Files are saved in `Created-Files/<filetype>/`.

5. **Repeat or exit**  
   You can run the process again or exit.

## Example Output

```
Created-Files/
├── md/
│   ├── Report_1_Draft.md
│   ├── Report_2_Draft.md
│   └── ...
├── xlsx/
│   ├── Report_1_Draft.xlsx
│   └── ...
└── ...
```

## Notes

- The script replaces illegal filename characters with underscores.
- If no number is present in your naming convention, numbering starts at 1 and is appended to the end.
- All prompts and errors are colorized for clarity.

## License

MIT License

---

_Created by Allen_
