# HTML to DOCX Converter

This Python application reads an HTML file, removes style tags from the head section, and converts it to DOCX format.

## Requirements

- Python 3.6 or higher
- Required libraries (listed in requirements.txt)

## Installation

1. Ensure Python 3 is installed on your system:
   - Windows: Download and install from [python.org](https://www.python.org/downloads/)
   - Linux: Most distributions come with Python pre-installed
   - macOS: Install via Homebrew (`brew install python`) or from [python.org](https://www.python.org/downloads/)

2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

If you encounter issues with lxml installation on Windows, try installing packages individually:

```bash
pip install beautifulsoup4 python-docx
pip install lxml
pip install html2docx
```

## Usage

Run the script with the HTML file as an argument:

```bash
python html_to_docx_converter.py irish-boy.html
```

If no argument is provided, the script will default to processing `irish-boy.html`.

The output will be saved as a DOCX file with the same base name as the input file (e.g., `irish-boy.docx`).

### Troubleshooting

If you encounter errors:

1. **Missing modules**: Ensure all required packages are installed using the commands in the Installation section.

2. **Conversion errors**: Check that your HTML file is well-formed and valid.

## Features

- Removes all style tags from the HTML head section
- Preserves the content structure
- Converts the modified HTML to DOCX format
