#!/usr/bin/env python3
"""
HTML to DOCX Converter
This script reads an HTML file, removes style tags from the head section,
and converts the modified HTML to DOCX format.
"""

import os
import sys
from bs4 import BeautifulSoup
from html2docx import html2docx

def read_html_file(file_path):
    """Read HTML file and return its content."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

def remove_styles(html_content):
    """Remove style tags from the HTML head section."""
    soup = BeautifulSoup(html_content, 'lxml')
    
    # Find and remove all style tags in the head
    style_tags = soup.head.find_all('style')
    for style_tag in style_tags:
        style_tag.decompose()
    
    return str(soup)

def convert_to_docx(html_content, output_path):
    """Convert HTML content to DOCX and save to output path."""
    try:
        # Extract title from HTML or use a default title
        soup = BeautifulSoup(html_content, 'lxml')
        title_tag = soup.title
        title = title_tag.string if title_tag else "Converted Document"
        
        # Convert HTML to DOCX with title parameter
        docx_result = html2docx(html_content, title)
        
        # Write to file
        with open(output_path, 'wb') as output_file:
            # If result is BytesIO, read it first
            if hasattr(docx_result, 'read'):
                output_file.write(docx_result.getvalue())
            else:
                output_file.write(docx_result)
        
        print(f"Successfully converted to {output_path}")
    except Exception as e:
        print(f"Error converting to DOCX: {e}")
        sys.exit(1)

def main():
    """Main function to process the HTML file and convert to DOCX."""
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = "irish-boy.html"  # Default input file
    
    # Generate output filename
    base_name = os.path.splitext(input_file)[0]
    output_file = f"{base_name}.docx"
    
    print(f"Processing {input_file}...")
    
    # Read HTML file
    html_content = read_html_file(input_file)
    
    # Remove styles
    print("Removing style tags...")
    modified_html = remove_styles(html_content)
    
    # Convert to DOCX
    print(f"Converting to DOCX...")
    convert_to_docx(modified_html, output_file)

if __name__ == "__main__":
    main()