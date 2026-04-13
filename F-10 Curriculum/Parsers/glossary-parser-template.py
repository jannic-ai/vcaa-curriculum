#!/usr/bin/env python3
"""
Victorian Curriculum Glossary Parser v3.0
==========================================

Parses curriculum glossary documents (Word format) and converts them to CSV format
suitable for loading into Neo4j or other graph databases.

FEATURES:
- Handles diagrams and images by replacing them with customizable reference text
- Detects BOTH image markdown AND text-based diagram references
- Removes URL citations and attributions
- Band-based organisation: Foundation to Year 10 vs Years 11 and 12
- Subject-specific VCE glossary extraction

VERSION HISTORY:
v2.2 - Text-based diagram detection for embedded charts/tables
v3.0 - Added VCE/Senior Years glossary support with band differentiation
       Separate sections for F-10 and VCE
       Direct table extraction for VCE study designs

Requirements:
- pandoc (for converting .docx to markdown)
- Python 3.6+
- python-docx (for VCE table extraction)

==============================================================================
SECTION 1: F-10 GLOSSARIES (Foundation to Year 10)
==============================================================================

Usage for F-10:
    python victorian_glossary_parser_v3.0.py <input_file> <subject_name> [output_file] [--diagram-ref TEXT]

Examples:
    python victorian_glossary_parser_v3.0.py Maths_glossary.docx "Mathematics" Maths_glossary.csv
    
    # Custom diagram reference
    python glossary_parser.py Science.docx "Science" --diagram-ref "See VCAA website"
    
    # Keep original image references (no replacement or URL cleaning)
    python glossary_parser.py History.docx "History" --diagram-ref "none"

Diagram & URL Handling:
    By default, replaces TWO types of diagram references:
    
    1. IMAGE MARKDOWN (from pandoc conversion):
       ![](media/image1.png) → "(See diagram in {Subject} Glossary...)"
    
    2. TEXT PATTERNS (embedded diagrams not converted by pandoc):
       "shown below" → "shown below (See diagram in {Subject} Glossary...)"
       "diagram below" → "diagram below (See diagram in {Subject} Glossary...)"
       "table below" → "table below (See diagram in {Subject} Glossary...)"
    
    AND removes all URL citations:
    - "Image:" or "Images:" citation lines (e.g., "Image: Pierce, Rod...")
    - "Data:" source URLs (e.g., "Data: <http://example.com>")
    - "Map data:" attributions (e.g., "Map data: © 2018 Google")
    - Standalone URLs in angle brackets
    - Markdown links in citations
    
    You can customize diagram text with --diagram-ref flag.
    Use "none" to keep original image markdown and URLs (not recommended).

Author: Created for Victorian Curriculum V2.0 GraphRAG Project
License: MIT
Version: 2.2 (with comprehensive diagram detection - image markdown + text patterns)
"""

import re
import csv
import sys
import os
import subprocess
from pathlib import Path


def convert_docx_to_markdown(docx_path, md_path):
    """
    Convert Word document to markdown using pandoc.
    
    Args:
        docx_path: Path to input .docx file
        md_path: Path to output .md file
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    try:
        subprocess.run(
            ['pandoc', docx_path, '-o', md_path],
            check=True,
            capture_output=True,
            text=True
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error converting document: {e.stderr}")
        return False
    except FileNotFoundError:
        print("Error: pandoc not found. Please install pandoc:")
        print("  Ubuntu/Debian: sudo apt-get install pandoc")
        print("  macOS: brew install pandoc")
        print("  Windows: Download from https://pandoc.org/installing.html")
        return False


def add_text_based_diagram_references(text, subject, diagram_ref_text=None):
    """
    Add diagram references for text patterns that indicate a diagram/table/figure.
    
    This catches cases where text says "shown below" or "diagram below" but there's
    no image markdown (e.g., embedded charts, tables that pandoc doesn't convert).
    
    Patterns detected:
    - "shown below", "is shown below", "are shown below"
    - "diagram below", "in the diagram below"
    - "table below", "in the table below"
    - "figure below"
    - "as shown below"
    
    Args:
        text: The text to scan for diagram references
        subject: Subject name for default reference text
        diagram_ref_text: Custom reference text, or None for default
        
    Returns:
        tuple: (text with added references, count of references added)
    """
    if diagram_ref_text == "none":
        return text, 0
    
    # Default reference text if none provided
    if diagram_ref_text is None:
        vcaa_ref = f" (See diagram in {subject} Glossary document on VCAA website)"
    else:
        vcaa_ref = f" {diagram_ref_text}" if not diagram_ref_text.startswith("(") else f" {diagram_ref_text}"
    
    # Don't add if already present
    if vcaa_ref.strip() in text:
        return text, 0
    
    # Patterns to match and where to insert the reference
    # Format: (pattern_to_find, replacement_with_reference)
    # Order matters - more specific patterns first
    patterns = [
        # Pattern 1: "shown below" with various endings
        (r'(,\s+as shown below)(,)', r'\1' + vcaa_ref + r'\2'),  # "as shown below,"
        (r'(,\s+as shown below)(\.)', r'\1' + vcaa_ref + r'\2'),
        (r'(is shown below)(\s+\()', r'\1' + vcaa_ref + r'\2'),  # "is shown below (something)"
        (r'(is shown below)(\.)', r'\1' + vcaa_ref + r'\2'),
        (r'(shown below)(\.)', r'\1' + vcaa_ref + r'\2'),
        (r'(are shown below)(\.)', r'\1' + vcaa_ref + r'\2'),
        
        # Pattern 2: "diagram below" variations
        (r'(as indicated in the diagram below)(\s+)', r'\1' + vcaa_ref + r'\2'),
        (r'(diagram below)(\s+are)', r'\1' + vcaa_ref + r'\2'),  # "diagram below are said"
        (r'(diagram below)(:)', r'\1' + vcaa_ref + r'\2'),
        (r'(in the diagram below,)(\s+)', r'\1' + vcaa_ref + r'\2'),
        
        # Pattern 3: table/figure references
        (r'(in the frequency table below)(:)', r'\1' + vcaa_ref + r'\2'),
        (r'(table below)(:)', r'\1' + vcaa_ref + r'\2'),  # Generic "table below:"
        (r'(table below)(\.)', r'\1' + vcaa_ref + r'\2'),
        (r'(figure below)(\.)', r'\1' + vcaa_ref + r'\2'),
    ]
    
    added = 0
    for pattern, replacement in patterns:
        if re.search(pattern, text, re.IGNORECASE):
            text = re.sub(pattern, replacement, text, count=1, flags=re.IGNORECASE)
            added = 1
            break  # Only apply one pattern per term
    
    return text, added


def replace_diagram_references(text, subject, diagram_ref_text=None):
    """
    Replace diagram/image references with customizable reference text.
    Also removes URL citations and attributions.
    
    Handles TWO types of diagram references:
    1. IMAGE MARKDOWN: ![](media/image1.png){width="..." height="..."}
    2. TEXT PATTERNS: "shown below", "diagram below", "table below", etc.
    
    This comprehensive approach catches:
    - Explicit image markdown from pandoc conversion
    - Embedded charts/diagrams that pandoc doesn't convert to markdown
    - Tables and figures referenced in text
    
    Also removes:
    - "Image:" or "Images:" citation lines
    - "Data:" citation lines with URLs
    - "Map data:" attributions
    - Standalone URLs in angle brackets
    - Markdown-style links in citations
    
    Args:
        text: The text containing potential image references
        subject: Subject name for default reference text
        diagram_ref_text: Custom reference text, or None for default
        
    Returns:
        tuple: (cleaned text, total reference count)
    """
    if diagram_ref_text == "none":
        return text, 0
    
    # Default reference text if none provided
    if diagram_ref_text is None:
        diagram_ref_text = f"(See diagram in {subject} Glossary document on VCAA website)"
    
    # STEP 1: Replace markdown image references
    image_pattern = r'!\[([^\]]*)\]\([^\)]+\.(?:png|jpg|jpeg|gif|svg|emf|wmf|bmp)\)(?:\{[^\}]+\})?'
    image_count = len(re.findall(image_pattern, text, re.IGNORECASE | re.MULTILINE))
    text = re.sub(image_pattern, diagram_ref_text, text, flags=re.IGNORECASE | re.MULTILINE)
    
    # STEP 2: Add references for text-based diagram indicators
    text, text_ref_count = add_text_based_diagram_references(text, subject, diagram_ref_text)
    
    # STEP 3: Remove URL citations and attributions
    # Remove "Image:" or "Images:" citation lines (entire line)
    text = re.sub(r'\s*Images?:\s*[^\n]+', '', text)
    
    # Remove "Data:" citation lines with URLs
    text = re.sub(r'\s*Data:\s*<[^>]+>', '', text)
    
    # Remove "Map data:" attributions
    text = re.sub(r'\s*Map data:\s*©[^\n]+', '', text)
    
    # Remove any remaining standalone URLs in angle brackets
    text = re.sub(r'\s*<https?://[^>]+>', '', text)
    
    # Remove markdown-style links [text](url) in citations
    text = re.sub(r'\[([^\]]+)\]\(https?://[^\)]+\)', r'\1', text)
    
    # Clean up multiple spaces
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    total_count = image_count + text_ref_count
    
    return text, total_count


def parse_glossary_markdown(filepath, subject, diagram_ref_text=None):
    """
    Parse a glossary markdown file and extract terms and definitions.
    
    This parser handles:
    - Terms with and without markdown IDs (e.g., "term {#id .class}")
    - Alphabetical section headers (single letters A-Z)
    - Multiple paragraph definitions
    - Diagram/image references (replaced with reference text)
    - Various whitespace and formatting issues
    
    Args:
        filepath: Path to markdown file
        subject: Subject name for the glossary
        diagram_ref_text: Custom text for diagram references, or None for default
        
    Returns:
        tuple: (glossary list, total_images_replaced)
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print(f"Error: File not found: {filepath}")
        return [], 0
    except Exception as e:
        print(f"Error reading file: {e}")
        return [], 0
    
    glossary = []
    total_images = 0
    
    # Remove the title (first # heading) - usually the document title
    content = re.sub(r'^#[^#].*?\n', '', content, count=1)
    
    # Split content into sections by ## or ### headings
    # Some glossaries use ## and some use ###
    sections = re.split(r'\n###?\s+', content)
    
    for section in sections:
        if not section.strip():
            continue
        
        # Extract term and definition
        # Handle both formats: "term {#id .class}" and just "term"
        match = re.match(r'([^{#\n]+?)(?:\s+\{[^}]+\})?\s*\n\n(.*)', section, re.DOTALL)
        
        if not match:
            continue
        
        term = match.group(1).strip()
        definition = match.group(2).strip()
        
        # Skip single letter section headers (A, B, C, etc.)
        # These are common in glossaries for alphabetical organization
        if len(term) == 1 and term.isupper():
            continue
        
        # Clean up the definition
        # Remove any trailing ## or ### sections that might have been captured
        definition = re.sub(r'\n###?\s+.*', '', definition).strip()
        
        # Replace diagram references BEFORE cleaning whitespace
        definition, image_count = replace_diagram_references(definition, subject, diagram_ref_text)
        total_images += image_count
        
        # Remove extra whitespace and newlines within the definition
        # Converts multiple spaces/newlines to single spaces
        definition = re.sub(r'\s+', ' ', definition)
        
        # Capitalize first letter of term for consistency
        if term:
            term = term[0].upper() + term[1:] if len(term) > 1 else term.upper()
        
        # Only add if we have both term and definition
        if term and definition:
            glossary.append({
                'Subject': subject,
                'Term': term,
                'Definition': definition
            })
    
    return glossary, total_images


def write_csv(glossary, output_file):
    """
    Write glossary entries to CSV file.
    
    Args:
        glossary: List of dictionaries with Subject, Term, Definition
        output_file: Path to output CSV file
        
    Returns:
        bool: True if write successful, False otherwise
    """
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['Subject', 'Term', 'Definition'])
            writer.writeheader()
            writer.writerows(glossary)
        return True
    except Exception as e:
        print(f"Error writing CSV file: {e}")
        return False


def validate_input_file(filepath):
    """
    Validate that input file exists and is a Word document.
    
    Args:
        filepath: Path to input file
        
    Returns:
        bool: True if valid, False otherwise
    """
    if not os.path.exists(filepath):
        print(f"Error: Input file not found: {filepath}")
        return False
    
    if not filepath.lower().endswith(('.docx', '.doc')):
        print(f"Warning: Input file does not appear to be a Word document: {filepath}")
        print("Continuing anyway, but conversion may fail...")
    
    return True


def parse_arguments():
    """
    Parse command line arguments.
    
    Returns:
        tuple: (input_file, subject_name, output_file, diagram_ref_text)
    """
    # Remove script name
    args = sys.argv[1:]
    
    if len(args) < 2:
        return None, None, None, None, False
    
    input_file = args[0]
    subject_name = args[1]
    output_file = None
    diagram_ref_text = None
    is_vce = False
    
    # Parse remaining arguments
    i = 2
    while i < len(args):
        if args[i] == '--diagram-ref' and i + 1 < len(args):
            diagram_ref_text = args[i + 1]
            i += 2
        elif args[i] == '--vce':
            is_vce = True
            i += 1
        else:
            # Assume it's the output file
            if output_file is None:
                output_file = args[i]
            i += 1
    
    # Generate output filename if not provided
    if output_file is None:
        if is_vce:
            subject_slug = subject_name.lower().replace(' ', '_')
            output_file = f"vcaa_vce_{subject_slug}_glossary.csv"
        else:
            output_file = Path(input_file).stem + '_glossary.csv'
    
    return input_file, subject_name, output_file, diagram_ref_text, is_vce


def main():
    """Main entry point for the glossary parser."""
    
    # Parse arguments
    input_file, subject_name, output_file, diagram_ref_text, is_vce = parse_arguments()
    
    # Check if we have minimum required arguments
    if input_file is None or subject_name is None:
        print("Usage:")
        print("  F-10:  python glossary_parser.py <input_file> <subject_name> [output_file] [--diagram-ref TEXT]")
        print("  VCE:   python glossary_parser.py <input_file> <subject_name> --vce [output_file]")
        print("")
        print("Examples:")
        print("  F-10:")
        print('    python glossary_parser.py Maths.docx "Mathematics" Maths_glossary.csv')
        print('    python glossary_parser.py Science.docx "Science" --diagram-ref "See VCAA website"')
        print("")
        print("  VCE:")
        print('    python glossary_parser.py English_SD.docx "English" --vce')
        print('    python glossary_parser.py Psychology_SD.docx "Psychology" --vce vcaa_vce_psych_glossary.csv')
        print("")
        print("Options:")
        print('  --vce                        : Parse VCE study design glossary (Years 11 and 12)')
        print('  --diagram-ref "custom text"  : Use custom text for diagram references (F-10 only)')
        print('  --diagram-ref "none"         : Keep original image markdown (F-10 only)')
        print("")
        print("For more information, see the script documentation at the top of this file.")
        sys.exit(1)
    
    # Route to appropriate parser
    if is_vce:
        # VCE glossary parsing (direct table extraction)
        parse_vce_glossary(input_file, subject_name, output_file)
    else:
        # F-10 glossary parsing (pandoc markdown conversion)
        
        # Validate input file
        if not validate_input_file(input_file):
            sys.exit(1)
        
        print(f"Parsing glossary for: {subject_name}")
        print(f"Input file: {input_file}")
        print(f"Output file: {output_file}")
        if diagram_ref_text:
            if diagram_ref_text == "none":
                print(f"Diagram handling: Keep original image references")
            else:
                print(f"Diagram reference: {diagram_ref_text}")
        else:
            print(f"Diagram reference: (See diagram in {subject_name} Glossary document on VCAA website)")
        print("")
        
        # Step 1: Convert Word document to Markdown
        temp_md_file = Path(input_file).stem + '_temp.md'
        print(f"Step 1: Converting Word document to Markdown...")
        
        if not convert_docx_to_markdown(input_file, temp_md_file):
            sys.exit(1)
        
        print(f"  ✓ Converted to {temp_md_file}")
        
        # Step 2: Parse the Markdown file
        print(f"Step 2: Parsing glossary terms...")
        
        glossary, total_images = parse_glossary_markdown(temp_md_file, subject_name, diagram_ref_text)
        
        if not glossary:
            print("  ✗ No glossary terms found!")
            print("    Check that your document has ## or ### headings for terms")
            os.remove(temp_md_file)
            sys.exit(1)
        
        print(f"  ✓ Found {len(glossary)} terms")
        if total_images > 0:
            print(f"  ✓ Replaced {total_images} diagram references")
        
        # Step 3: Write to CSV
        print(f"Step 3: Writing to CSV...")
        
        if not write_csv(glossary, output_file):
            os.remove(temp_md_file)
            sys.exit(1)
        
        print(f"  ✓ Wrote {output_file}")
        
        # Clean up temporary markdown file
        os.remove(temp_md_file)
        
        # Show sample of parsed terms
        print("")
        print("Sample terms (first 5):")
        for entry in glossary[:5]:
            term_preview = entry['Term']
            def_preview = entry['Definition'][:60] + "..." if len(entry['Definition']) > 60 else entry['Definition']
            print(f"  - {term_preview}")
            if "diagram" in def_preview.lower():
                print(f"    └─ Contains diagram reference")
        
        print("")
        print(f"✓ SUCCESS: Glossary parsing complete!")
        print(f"  Total terms: {len(glossary)}")
        print(f"  Diagrams replaced: {total_images}")
        print(f"  Output: {output_file}")


# ==============================================================================
# SECTION 2: VCE/SENIOR YEARS GLOSSARIES (Years 11 and 12)
# ==============================================================================

def extract_vce_glossary_table(docx_path, subject_name):
    """
    Extract glossary table from VCE study design document.
    
    Searches for "Terms used in this study" heading and extracts the table
    that follows it.
    
    Args:
        docx_path: Path to VCE study design Word document
        subject_name: Subject name (e.g., "English", "Psychology")
        
    Returns:
        List of dicts with 'Term', 'Definition', 'Subject', 'Band' keys
    """
    from docx import Document
    
    doc = Document(docx_path)
    glossary_terms = []
    
    print(f"Searching for 'Terms used in this study' table in {docx_path}...")
    
    # Find tables in the document
    for table in doc.tables:
        if len(table.rows) >= 2 and len(table.columns) >= 2:
            # Check if this is a glossary table
            first_row = table.rows[0]
            if len(first_row.cells) >= 2:
                col1_header = first_row.cells[0].text.strip().lower()
                col2_header = first_row.cells[1].text.strip().lower()
                
                # Look for "Term" and "Definition" headers
                if 'term' in col1_header and 'definition' in col2_header:
                    print(f"  ✓ Found glossary table with {len(table.rows)-1} terms")
                    
                    # Extract terms (skip header row)
                    for row_idx in range(1, len(table.rows)):
                        row = table.rows[row_idx]
                        if len(row.cells) >= 2:
                            term = row.cells[0].text.strip()
                            definition = row.cells[1].text.strip()
                            
                            if term and definition:
                                glossary_terms.append({
                                    'Term': term,
                                    'Definition': definition,
                                    'Subject': subject_name,
                                    'Band': 'Years 11 and 12'
                                })
                    
                    return glossary_terms
    
    print(f"  ✗ No glossary table found in document")
    return []


def write_vce_glossary_csv(glossary_terms, output_file):
    """
    Write VCE glossary terms to CSV file.
    
    CSV format: Term,Definition,Subject,Band
    """
    import csv
    
    if not glossary_terms:
        print("  ✗ No terms to write!")
        return False
    
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            fieldnames = ['Term', 'Definition', 'Subject', 'Band']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            
            writer.writeheader()
            writer.writerows(glossary_terms)
        
        return True
    except Exception as e:
        print(f"  ✗ Error writing CSV: {e}")
        return False


def parse_vce_glossary(docx_path, subject_name, output_file=None):
    """
    Main function to parse VCE glossary from study design document.
    
    Args:
        docx_path: Path to VCE study design Word document
        subject_name: Subject name (e.g., "English", "Psychology")
        output_file: Optional output CSV filename
        
    Returns:
        List of glossary terms
    """
    # Generate output filename if not provided
    if not output_file:
        subject_slug = subject_name.lower().replace(' ', '_')
        output_file = f"vcaa_vce_{subject_slug}_glossary.csv"
    
    print("")
    print("="*70)
    print(f"VCE GLOSSARY PARSER - {subject_name}")
    print("="*70)
    print(f"Input file:  {docx_path}")
    print(f"Subject:     {subject_name}")
    print(f"Band:        Years 11 and 12")
    print(f"Output file: {output_file}")
    print("")
    
    # Step 1: Extract glossary table from Word document
    print("Step 1: Extracting glossary table from study design...")
    
    glossary_terms = extract_vce_glossary_table(docx_path, subject_name)
    
    if not glossary_terms:
        print("  ✗ No glossary terms found!")
        print("    Check that the document has a 'Terms used in this study' table")
        print("    with 'Term' and 'Definition' column headers")
        return []
    
    print(f"  ✓ Found {len(glossary_terms)} terms")
    
    # Step 2: Write to CSV
    print("Step 2: Writing to CSV...")
    
    if not write_vce_glossary_csv(glossary_terms, output_file):
        return []
    
    print(f"  ✓ Wrote {output_file}")
    
    # Show sample of parsed terms
    print("")
    print("Sample terms (first 5):")
    for entry in glossary_terms[:5]:
        term_preview = entry['Term']
        def_preview = entry['Definition'][:60] + "..." if len(entry['Definition']) > 60 else entry['Definition']
        print(f"  - {term_preview}: {def_preview}")
    
    print("")
    print(f"✓ SUCCESS: VCE glossary parsing complete!")
    print(f"  Total terms: {len(glossary_terms)}")
    print(f"  Output: {output_file}")
    
    return glossary_terms


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

if __name__ == '__main__':
    main()
