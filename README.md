# HTML to DOCX Converter

This Python script converts HTML files to DOCX format with specific formatting requirements. It was designed to convert Project Gutenberg HTML books to properly formatted DOCX files.

## Features

- Identifies blocks without text content and skips them
- Identifies bookmarks and adds them to the DOCX
- Identifies paragraph types (toc/p/h1/h2/h3/etc) and applies appropriate formatting
- Links bookmarks during the next written content, not immediately
- Adds section page breaks for h2 headings (after TOC is encountered)
- Creates PageRef form fields for TOC entries to display page numbers
- Configures page and styles according to specifications:
  - Mirrored margins (1" top/inside, 0.75" bottom/outside, 0.4" header)
  - Normal style uses Book Antiqua 10pt, full justification, 1.15 spacing, 11.5pt after-spacing

## Requirements

- Python 3.6+
- Required packages:
  - python-docx
  - beautifulsoup4

## Installation

1. Ensure Python 3 is installed on your system:
   - Windows: Download and install from [python.org](https://www.python.org/downloads/)
   - Linux: Most distributions come with Python pre-installed
   - macOS: Install via Homebrew (`brew install python`) or from [python.org](https://www.python.org/downloads/)

2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python html_to_docx_converter.py input.html output.docx
```

Example:

```bash
python html_to_docx_converter.py irish-boy.html irish-boy-new.docx
```

## How It Works

1. **HTML Parsing**: The script uses BeautifulSoup to parse the HTML file and extract the content.

2. **Content Analysis**: It analyzes the HTML content, identifies different types of elements (headings, paragraphs, TOC entries), and filters out elements without meaningful text content.

3. **Bookmark Collection**: It collects all elements with ID attributes as potential bookmarks.

4. **Document Setup**: It creates a new DOCX document with the specified styles and formatting.

5. **Content Processing**: It processes each element based on its type:
   - TOC entries: Creates paragraphs with TOC style and adds PageRef fields to link to bookmarks
   - Headings: Creates headings with appropriate styles and adds section breaks for h2 headings
   - Paragraphs: Creates normal paragraphs with the specified formatting

6. **Bookmark Handling**: It adds bookmarks to the document and links them to the appropriate content.

## Troubleshooting

If you encounter errors:

1. **Missing modules**: Ensure all required packages are installed using the commands in the Installation section.

2. **Conversion errors**: Check that your HTML file is well-formed and valid.

3. **Complex HTML**: For complex HTML files, you may need to adjust the content analysis logic in the script.

## License

This script is provided under the MIT License.
