#!/usr/bin/env python3
'''The script allows users to generate multiple files of different types based on a naming convention
    and the number of copies specified.
    
    Parameters
    ----------
    name : str
        The script you provided is a Python program that allows users to create multiple files of different
    types (Excel, Word, PowerPoint, CSV, Markdown, YAML) based on a naming convention and the number of
    copies specified by the user. The script sanitizes the input filename, prompts the user for the file
    
    Returns
    -------
        The script is returning an exit code at the end of the `main()` function. The exit code is used to
    indicate the success or failure of the script execution. In this case, the script returns `0` if it
    completes successfully and returns `1` if there are any errors encountered during execution.
    
'''
import sys
import re
import logging
from pathlib import Path
from typing import Dict, Callable

import csv  # for CSV output
import yaml  # third-party: pip install PyYAML
from openpyxl import Workbook  # third-party: pip install openpyxl
from docx import Document  # third-party: pip install python-docx
from pptx import Presentation  # third-party: pip install python-pptx

from colorama import Fore, Style, init  # Add colorama import

init(autoreset=True)  # Initialize colorama

# Compile regex once at module level: digit runs not preceded by '_'
NUM_PATTERN = re.compile(r"(?<!_)(\d+)")

# Supported file types and their descriptions
FILE_TYPES: Dict[str, str] = {
    "md": "Markdown",
    "xlsx": "Excel Workbook",
    "csv": "CSV File",
    "docx": "Word Document",
    "pptx": "PowerPoint Presentation",
    "yaml": "YAML File",
}

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def sanitize(name: str) -> str:
    """
    Sanitize a raw filename string.

    This function replaces illegal filesystem characters with underscores and trims whitespace.
    Args:
        name (str): The raw filename string provided by the user.
    Returns:
        str: A sanitized filename safe for use in filesystem operations.
    """
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


def get_positive_int(prompt: str) -> int:
    """
    Prompt the user until they enter a positive integer.
    Args:
        prompt (str): The input prompt message.
    Returns:
        int: A positive integer entered by the user.
    Raises:
        ValueError: If the input is not a positive integer.
    """
    value = input(prompt).strip()
    if not value.isdigit() or int(value) < 1:
        raise ValueError(
            Fore.RED
            + Style.BRIGHT
            + "Input must be a positive integer."
            + Style.RESET_ALL
        )
    return int(value)


def choose_extension() -> str:
    """
    Display a numbered menu of supported file types and return the chosen extension.
    Returns:
        str: The file extension, starting with a dot (e.g. ".md").
    Raises:
        ValueError: If the selection is invalid.
    """
    print(Fore.YELLOW + Style.BRIGHT + "Select file type:" + Style.RESET_ALL)

    for index, (ext, desc) in enumerate(FILE_TYPES.items(), start=1):
        print(
            Fore.YELLOW
            + Style.BRIGHT
            + (f"  {index}. {desc} ({ext})")
            + Style.RESET_ALL
        )

    choice = input(
        Fore.MAGENTA + "Enter the number of your choice:" + Style.RESET_ALL
    ).strip()
    if not choice.isdigit():
        raise ValueError("Selection must be a number.")

    idx = int(choice)
    if idx < 1 or idx > len(FILE_TYPES):
        raise ValueError(
            f"Please enter a number between 1 and {len(FILE_TYPES)}.")

    ext_key = list(FILE_TYPES)[idx - 1]
    return f".{ext_key}"


# Dispatch table mapping extensions to creation functions
FileHandler = Callable[[Path], None]

# Define functions to create files of various types


def create_xlsx(path: Path) -> None:
    """Create a blank Excel workbook."""
    wb = Workbook()
    wb.active.title = "Sheet1"
    wb.save(path)


def create_docx(path: Path) -> None:
    """Create a blank Word document."""
    doc = Document()
    doc.add_paragraph("")
    doc.save(path)


def create_pptx(path: Path) -> None:
    """Create a blank PowerPoint presentation."""
    prs = Presentation()
    prs.save(path)


def create_csv(path: Path) -> None:
    """Create an empty CSV file (no rows)."""
    with path.open("w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        # Example: writer.writerow(["Header1", "Header2"])


def create_md(path: Path) -> None:
    """Create a Markdown file with a default heading."""
    with path.open("w", encoding="utf-8") as f:
        f.write(f"# {path.stem}\n\n")


def create_yaml(path: Path) -> None:
    """Create a YAML file with an empty dictionary."""
    with path.open("w", encoding="utf-8") as f:
        yaml.safe_dump({}, f)


# Extension mapping
EXTENSION_HANDLERS: Dict[str, FileHandler] = {
    ".xlsx": create_xlsx,
    ".docx": create_docx,
    ".pptx": create_pptx,
    ".csv": create_csv,
    ".md": create_md,
    ".yaml": create_yaml,
}


def main() -> int:
    """
    Main entry point.
    Allows repeating the generation process.
    Returns:
        int: Exit code (0 for success, 1 for errors).
    """
    try:
        while True:
            raw_name = input(
                Fore.CYAN
                + Style.BRIGHT
                + "Enter the naming convention"
                + Style.RESET_ALL
                + "(File_0_Content):"
            ).strip()
            base = sanitize(raw_name)
            if not base:
                logging.error(
                    Fore.RED
                    + Style.BRIGHT
                    + "Naming convention cannot be empty."
                    + Style.RESET_ALL
                )
                return 1

            ext = choose_extension()
            copies = get_positive_int(
                Fore.CYAN
                + Style.BRIGHT
                + "Enter the number of copies to create: "
                + Style.RESET_ALL
            )

            # Create main output folder
            out_dir = Path("Created-Files")
            out_dir.mkdir(parents=True, exist_ok=True)

            # Create subfolder for this file type
            type_dir = out_dir / ext.lstrip(".")
            type_dir.mkdir(parents=True, exist_ok=True)

            matches = list(NUM_PATTERN.finditer(base))
            if matches:
                last = matches[-1]
                start_num = int(last.group(1))
                span_start, span_end = last.span(1)
            else:
                start_num = 0
                span_start = span_end = None

            width = len(str(start_num + copies))

            file_count = 0  # Track number of files created

            for i in range(1, copies + 1):
                new_str = str(start_num + i).zfill(width)
                if span_start is not None:
                    name_body = base[:span_start] + new_str + base[span_end:]
                else:
                    name_body = f"{base}{new_str}"

                filename = f"{name_body}{ext}"
                filepath = type_dir / filename

                handler = EXTENSION_HANDLERS.get(ext)
                if handler:
                    handler(filepath)
                else:
                    filepath.touch()

                file_count += 1  # Increment file count

            print(
                Fore.GREEN
                + Style.BRIGHT
                + f"Total files created: {file_count}"
                + Style.RESET_ALL
            )

            again = (
                input(
                    Fore.YELLOW
                    + Style.BRIGHT
                    + "Do you want to run the program again? (y/n): "
                    + Style.RESET_ALL
                )
                .strip()
                .lower()
            )
            if again != "y":
                break

        return 0

    except ValueError as ve:
        logging.error(Fore.RED + str(ve) + Style.RESET_ALL)
        return 1
    except OSError as oe:
        logging.error(Fore.RED + f"Filesystem error: {oe}" + Style.RESET_ALL)
        return 1


if __name__ == "__main__":
    sys.exit(main())

# End of script
