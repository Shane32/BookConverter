#!/usr/bin/env python3
"""
HTML to JSON Converter

This script converts HTML files to a specific JSON format for book content:
1. Extracts book metadata (title, subtitle, author)
2. Extracts chapters with their numbers, titles, and paragraphs
3. Implements strict mode to validate only allowed tags contain text
4. Converts Roman numerals in chapter headings to integers
5. Outputs formatted JSON file

Usage:
    python html_to_json_converter.py [input_file] [output_file]

Default:
    input_file: irish-boy.html
    output_file: irish-boy.json
"""

import sys
import re
import json
from bs4 import BeautifulSoup, NavigableString

def normalize_html_whitespace(text):
    """
    Normalize whitespace according to HTML rules.
    
    Args:
        text (str): Text to normalize
        
    Returns:
        str: Normalized text
    """
    if not text:
        return text
        
    # Replace all whitespace sequences (including newlines) with a single space
    normalized = re.sub(r'\s+', ' ', text)
    
    # Trim leading and trailing whitespace
    normalized = normalized.strip()
    
    return normalized

def roman_to_int(roman):
    """
    Convert Roman numeral to integer.
    
    Args:
        roman (str): Roman numeral string
        
    Returns:
        int: Integer value
        
    Raises:
        ValueError: If the input is not a valid Roman numeral
    """
    print(f"Converting Roman numeral: {roman}")
    
    # Dictionary of Roman numerals
    roman_dict = {
        'I': 1,
        'V': 5,
        'X': 10,
        'L': 50,
        'C': 100,
        'D': 500,
        'M': 1000
    }
    
    # Validate Roman numeral format
    if not re.match(r'^[IVXLCDM]+$', roman, re.IGNORECASE):
        raise ValueError(f"Invalid Roman numeral: {roman}")
    
    # Convert to uppercase for consistency
    roman = roman.upper()
    
    result = 0
    prev_value = 0
    
    # Process from right to left
    for char in reversed(roman):
        if char not in roman_dict:
            raise ValueError(f"Invalid Roman numeral character: {char}")
        
        current_value = roman_dict[char]
        
        # If current value is greater than or equal to previous value, add it
        # Otherwise, subtract it (for cases like IV, IX, etc.)
        if current_value >= prev_value:
            result += current_value
        else:
            result -= current_value
            
        prev_value = current_value
    
    print(f"Converted {roman} to {result}")
    return result

def parse_chapter_heading(heading_text):
    """
    Parse chapter heading to extract chapter number and title.
    
    Args:
        heading_text (str): Chapter heading text
        
    Returns:
        tuple: (chapter_number, chapter_title)
        
    Raises:
        ValueError: If unable to parse the chapter heading
    """
    print(f"Parsing chapter heading: {heading_text}")
    
    # Try to match "CHAPTER X — TITLE" format
    match = re.match(r'CHAPTER\s+([IVXLCDM]+)\s+—\s+(.+)', heading_text, re.IGNORECASE)
    if match:
        roman_numeral, title = match.groups()
        try:
            chapter_number = roman_to_int(roman_numeral)
            return chapter_number, title.strip()
        except ValueError as e:
            raise ValueError(f"Failed to parse chapter number: {e}")
    
    # If no match, raise an error
    raise ValueError(f"Unable to parse chapter heading: {heading_text}")

def has_text_content(element):
    """
    Check if an element has meaningful text content.
    
    Args:
        element: BeautifulSoup element
        
    Returns:
        bool: True if element has text content, False otherwise
    """
    if element.name in ['br', 'hr']:
        return False
    
    # Get text content, stripping whitespace
    text = element.get_text(strip=True)
    
    # Check if there's any text content
    return bool(text)

def validate_element_in_strict_mode(element):
    """
    Validate element in strict mode - ensure it contains no child elements other than text.
    
    Args:
        element: BeautifulSoup element
        
    Raises:
        ValueError: If the element contains any non-text child elements
    """
    # Skip NavigableString objects
    if isinstance(element, NavigableString):
        return
    
    # Check all children - only allow text nodes (NavigableString)
    for child in element.children:
        if not isinstance(child, NavigableString):
            raise ValueError(f"Unrecognized element <{child.name}> found inside <{element.name}>")

def convert_html_to_json(html_file, json_file):
    """
    Convert HTML file to JSON format.
    
    Args:
        html_file (str): Path to HTML file
        json_file (str): Path to output JSON file
        
    Returns:
        None
    """
    print(f"Converting {html_file} to {json_file}")
    
    # Parse HTML
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Initialize book data structure
    book_data = {
        "book": {
            "title": "",
            "subtitle": "",
            "author": ""
        },
        "chapters": []
    }
    
    # Extract book metadata
    print("Extracting book metadata...")
    
    # Find title and subtitle
    title_tag = soup.find('h1')
    if title_tag:
        book_data["book"]["title"] = normalize_html_whitespace(title_tag.get_text())
        print(f"Found book title: {book_data['book']['title']}")
    
    subtitle_tag = soup.find('h3')
    if subtitle_tag:
        book_data["book"]["subtitle"] = normalize_html_whitespace(subtitle_tag.get_text())
        print(f"Found book subtitle: {book_data['book']['subtitle']}")
    
    # Find author
    author_tag = soup.find('h2', string=lambda text: text and "by" in text.lower())
    if author_tag:
        # Extract author name (remove "by" prefix)
        author_text = author_tag.get_text()
        author_text = normalize_html_whitespace(author_text)
        author_match = re.search(r'by\s+(.+)', author_text, re.IGNORECASE)
        if author_match:
            book_data["book"]["author"] = author_match.group(1).strip()
            print(f"Found book author: {book_data['book']['author']}")
    
    # Find all chapter headings (h2 elements)
    chapter_headings = soup.find_all('h2')
    
    # Filter out non-chapter headings (like the book title)
    chapter_headings = [h for h in chapter_headings if re.search(r'CHAPTER\s+[IVXLCDM]+', h.get_text(strip=True), re.IGNORECASE)]
    
    print(f"Found {len(chapter_headings)} chapter headings")
    
    # Process each chapter
    strict_mode = False
    for i, heading in enumerate(chapter_headings):
        try:
            # Parse chapter heading
            chapter_number, chapter_title = parse_chapter_heading(heading.get_text(strip=True))
            
            # Create chapter object
            chapter = {
                "number": chapter_number,
                "title": chapter_title,
                "paragraphs": []
            }
            
            # Find paragraphs for this chapter
            paragraphs = []
            
            # Get all paragraph elements between this heading and the next heading
            current_element = heading.next_sibling
            
            # If this is the first chapter, enable strict mode
            if i == 0:
                print("Enabling strict mode after first chapter heading")
                strict_mode = True
            
            # Collect paragraphs until next heading or end of document
            while current_element:
                # If we hit the next chapter heading, stop
                if current_element.name == 'h2' and re.search(r'CHAPTER\s+[IVXLCDM]+', current_element.get_text(strip=True), re.IGNORECASE):
                    break
                
                # In strict mode, reject any unrecognized element that contains text
                if strict_mode and current_element.name and current_element.name not in ['p', 'pre']:
                    # Special case: Allow <h3> with "THE END" at the end of the book
                    if current_element.name == 'h3' and current_element.get_text(strip=True) == "THE END":
                        print("Found 'THE END' marker, skipping...")
                        current_element = current_element.next_sibling
                        continue
                    
                    # Special case: Skip Project Gutenberg footer section
                    if current_element.name == 'section' and "END OF THE PROJECT GUTENBERG" in current_element.get_text(strip=True):
                        print("Found Project Gutenberg footer, skipping...")
                        current_element = current_element.next_sibling
                        continue
                    
                    # Check if it contains any readable text
                    if has_text_content(current_element):
                        error_msg = f"Unrecognized element <{current_element.name}> contains text: {current_element.get_text(strip=True)[:30]}..."
                        print(f"Error in strict mode: {error_msg}")
                        raise ValueError(error_msg)
                
                # If it's a paragraph with content, add it as a string
                if current_element.name == 'p' and has_text_content(current_element):
                    # In strict mode, validate that only allowed tags contain text
                    if strict_mode:
                        try:
                            validate_element_in_strict_mode(current_element)
                        except ValueError as e:
                            print(f"Error in strict mode: {e}")
                            raise
                    
                    # Add paragraph text with normalized whitespace
                    paragraph_text = current_element.get_text()
                    paragraph_text = normalize_html_whitespace(paragraph_text)
                    paragraphs.append(paragraph_text)
                
                # If it's a pre element with content, add it as an object
                elif current_element.name == 'pre' and has_text_content(current_element):
                    # In strict mode, validate that only allowed tags contain text
                    if strict_mode:
                        try:
                            validate_element_in_strict_mode(current_element)
                        except ValueError as e:
                            print(f"Error in strict mode: {e}")
                            raise
                    
                    # Add pre text as a quote object with normalized whitespace
                    quote_text = current_element.get_text()
                    quote_text = normalize_html_whitespace(quote_text)
                    paragraphs.append({
                        "type": "quote",
                        "content": quote_text
                    })
                
                # Move to next element
                current_element = current_element.next_sibling
            
            # Add paragraphs to chapter
            chapter["paragraphs"] = paragraphs
            print(f"Chapter {chapter_number}: '{chapter_title}' - {len(paragraphs)} paragraphs")
            
            # Add chapter to book data
            book_data["chapters"].append(chapter)
            
        except ValueError as e:
            print(f"Error processing chapter heading: {e}")
            raise
    
    # Write JSON output
    print(f"Writing JSON output to {json_file}")
    with open(json_file, 'w', encoding='utf-8') as file:
        json.dump(book_data, file, indent=2, ensure_ascii=False)
    
    print(f"Conversion complete. Book has {len(book_data['chapters'])} chapters with a total of {sum(len(chapter['paragraphs']) for chapter in book_data['chapters'])} paragraphs.")

def main():
    """
    Main function to handle command line arguments and run the converter.
    """
    # Default filenames
    input_file = "irish-boy.html"
    output_file = "irish-boy.json"
    
    # Parse command line arguments
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    
    try:
        # Run the converter
        convert_html_to_json(input_file, output_file)
        print("Conversion successful!")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()