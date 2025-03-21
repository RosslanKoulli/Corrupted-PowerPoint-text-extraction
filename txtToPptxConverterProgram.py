#!/usr/bin/env python3
# PPTX Text Preprocessor
# A program to clean and process text extracted from corrupted PowerPoint files

import sys
import os
import re
import argparse
import logging
from pathlib import Path

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def setup_argparse():
    """Set up command line argument parsing."""
    parser = argparse.ArgumentParser(description='Clean and process text extracted from corrupted PPTX files.')
    parser.add_argument('input_file', type=str, help='Path to the text file to process')
    parser.add_argument('--output', '-o', type=str, help='Path to the output file (default: input_file_cleaned.txt)')
    parser.add_argument('--aggressive', '-a', action='store_true', help='Use more aggressive cleaning')
    parser.add_argument('--preserve-bullets', '-b', action='store_true', help='Preserve bullet points and numbering')
    parser.add_argument('--use-nlp', '-n', action='store_true', help='Use NLP for advanced content analysis')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose output')
    return parser.parse_args()

def check_dependencies(use_nlp=False):
    """Check if required libraries are installed."""
    if use_nlp:
        try:
            import nltk
            import spacy
            return True
        except ImportError as e:
            logger.error(f"NLP libraries not installed. Error: {e}")
            logger.error("Please install required packages with: pip install nltk spacy")
            logger.error("And download the required models with: python -m spacy download en_core_web_sm")
            return False
    return True

def read_text_file(file_path):
    """Read the content from a text file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except UnicodeDecodeError:
        # Try with different encodings if UTF-8 fails
        encodings = ['latin-1', 'cp1252', 'iso-8859-1']
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    logger.info(f"Successfully read file using {encoding} encoding")
                    return f.read()
            except UnicodeDecodeError:
                continue
        
        logger.error("Failed to read text file with any encoding")
        return None
    except Exception as e:
        logger.error(f"Failed to read text file: {str(e)}")
        return None

def write_text_file(file_path, content):
    """Write content to a text file."""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        logger.info(f"Successfully wrote processed text to {file_path}")
        return True
    except Exception as e:
        logger.error(f"Failed to write to file: {str(e)}")
        return False

def is_bullet_point(line):
    """Check if a line is a bullet point."""
    bullet_patterns = [
        r'^\s*[-•*]\s+',         # Dash, bullet, or asterisk
        r'^\s*\d+\.\s+',          # Numbered item
        r'^\s*[a-zA-Z]\)\s+',     # Letter followed by parenthesis
        r'^\s*\(\d+\)\s+',        # Number in parentheses
        r'^\s*[ivxIVX]+\.\s+',    # Roman numerals
        r'^\s*[□■◆▪▫●○]\s+'      # Various bullet symbols
    ]
    
    for pattern in bullet_patterns:
        if re.match(pattern, line):
            return True
    
    return False

def is_slide_header(line, prev_line=""):
    """Check if a line is likely a slide header/title."""
    # Skip empty lines
    if not line.strip():
        return False
    
    # If all uppercase or starts with "Slide" followed by a number
    if line.isupper() or re.match(r'^Slide\s+\d+', line, re.IGNORECASE):
        return True
    
    # If it's less than 60 chars and ends with a colon
    if len(line) < 60 and line.rstrip().endswith(':'):
        return True
    
    # If it's relatively short, preceded by an empty line, and followed by content
    if len(line) < 80 and not prev_line.strip():
        return True
    
    # If it starts with a section number pattern
    if re.match(r'^\s*\d+(\.\d+)*\s+\w+', line):
        return True
    
    return False

def is_footer_or_header(line):
    """Check if a line is likely a slide footer or header."""
    footer_patterns = [
        r'confidential',
        r'proprietary',
        r'all\s+rights\s+reserved',
        r'copyright',
        r'page\s+\d+\s+of\s+\d+',
        r'^\s*\d+\s*$',           # Just a page number
        r'www\.',                  # Website
        r'@\w+\.\w+'               # Email domain
    ]
    
    for pattern in footer_patterns:
        if re.search(pattern, line, re.IGNORECASE):
            return True
    
    return False

def is_xml_or_formatting(line):
    """Check if a line contains XML/HTML tags or formatting codes."""
    # Check for XML tags
    if re.search(r'<[^>]+>', line):
        return True
    
    # Check for formatting codes common in PPT extraction
    if re.search(r'\[\w+\]|\{\w+\}|\[\[\w+\]\]', line):
        return True
    
    return False

def basic_clean(text, preserve_bullets=True):
    """Perform basic text cleaning."""
    if not text:
        return ""
    
    # Split into lines for processing
    lines = text.split('\n')
    cleaned_lines = []
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # Skip empty lines
        if not line:
            continue
        
        # Skip XML or formatting lines
        if is_xml_or_formatting(line):
            continue
        
        # Skip footer or header lines
        if is_footer_or_header(line):
            continue
        
        # Keep bullet points if specified
        if preserve_bullets and is_bullet_point(line):
            cleaned_lines.append(line)
            continue
        
        # Keep slide headers
        prev_line = lines[i-1] if i > 0 else ""
        if is_slide_header(line, prev_line):
            # Add a blank line before headers unless it's the first line
            if cleaned_lines:
                cleaned_lines.append("")
            cleaned_lines.append(line)
            continue
            
        # Handle regular text
        cleaned_lines.append(line)
    
    # Join the lines back together
    cleaned_text = '\n'.join(cleaned_lines)
    
    # Remove multiple consecutive newlines
    cleaned_text = re.sub(r'\n\s*\n', '\n\n', cleaned_text)
    
    return cleaned_text

def advanced_clean_with_nlp(text, preserve_bullets=True):
    """Clean text using NLP techniques."""
    if not text:
        return ""
    
    try:
        import spacy
        import nltk
        from nltk.tokenize import sent_tokenize
        
        # Ensure NLTK resources are available
        try:
            nltk.data.find('tokenizers/punkt')
        except LookupError:
            nltk.download('punkt', quiet=True)
        
        # Load spaCy model
        try:
            nlp = spacy.load("en_core_web_sm")
        except OSError:
            logger.info("Downloading spaCy model: en_core_web_sm")
            os.system("python -m spacy download en_core_web_sm")
            nlp = spacy.load("en_core_web_sm")
        
        # First, do basic cleaning
        text = basic_clean(text, preserve_bullets)
        
        # Split into paragraphs
        paragraphs = text.split('\n\n')
        processed_paragraphs = []
        
        for paragraph in paragraphs:
            # Skip empty paragraphs
            if not paragraph.strip():
                continue
            
            # Keep bullet points as is
            if preserve_bullets and is_bullet_point(paragraph.strip()):
                processed_paragraphs.append(paragraph)
                continue
            
            # Keep slide headers as is
            if is_slide_header(paragraph.strip()):
                processed_paragraphs.append(paragraph)
                continue
            
            # Process through spaCy for non-bullet, non-header text
            doc = nlp(paragraph)
            
            # Extract meaningful sentences
            meaningful_sentences = []
            for sent in doc.sents:
                sent_text = sent.text.strip()
                
                # Skip very short sentences that are likely fragments
                if len(sent_text.split()) < 3 and not re.search(r'\w+\s+\w+', sent_text):
                    continue
                
                # Skip sentences with no real content
                if not re.search(r'[a-zA-Z]{3,}', sent_text):
                    continue
                    
                meaningful_sentences.append(sent_text)
            
            if meaningful_sentences:
                processed_paragraphs.append(' '.join(meaningful_sentences))
        
        # Join everything back together
        processed_text = '\n\n'.join(processed_paragraphs)
        
        # Clean up any remaining issues
        processed_text = re.sub(r'\s{2,}', ' ', processed_text)  # Remove multiple spaces
        processed_text = re.sub(r'\n\s*\n', '\n\n', processed_text)  # Fix newlines
        
        return processed_text
        
    except ImportError:
        logger.warning("NLP libraries not available. Falling back to basic cleaning.")
        return basic_clean(text, preserve_bullets)
    except Exception as e:
        logger.error(f"Error during NLP processing: {str(e)}")
        return basic_clean(text, preserve_bullets)

def aggressive_clean(text):
    """Perform aggressive text cleaning to get only essential content."""
    if not text:
        return ""
    
    # First apply basic cleaning
    text = basic_clean(text, preserve_bullets=True)
    
    # Split into lines
    lines = text.split('\n')
    essential_lines = []
    
    current_section = None
    
    for line in lines:
        line = line.strip()
        
        # Skip empty lines
        if not line:
            continue
        
        # Capture section headers
        if is_slide_header(line):
            # Add a blank line before new sections unless it's the first
            if essential_lines:
                essential_lines.append("")
            essential_lines.append(line)
            current_section = line
            continue
        
        # Capture bullet points
        if is_bullet_point(line):
            # If we're in a section, indent the bullet point
            if current_section:
                essential_lines.append("  " + line)
            else:
                essential_lines.append(line)
            continue
            
        # Filter out common noise
        if len(line) < 4 or line.isdigit():
            continue
            
        # Keep lines with actual textual content
        if re.search(r'[a-zA-Z]{3,}', line):
            essential_lines.append(line)
    
    # Join everything back together
    essential_text = '\n'.join(essential_lines)
    
    # Remove any double spacing
    essential_text = re.sub(r'\s{2,}', ' ', essential_text)
    
    # Fix newlines
    essential_text = re.sub(r'\n\s*\n', '\n\n', essential_text)
    
    return essential_text

def extract_structured_content(text):
    """Extract and structure meaningful content from the text."""
    if not text:
        return ""
    
    # First, do basic cleaning
    text = basic_clean(text, preserve_bullets=True)
    
    # Split into lines for processing
    lines = text.split('\n')
    
    # Initialize structure
    structured_content = []
    current_section = {"title": None, "content": []}
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # Skip empty lines
        if not line:
            continue
            
        # Check if it's a section header
        prev_line = lines[i-1].strip() if i > 0 else ""
        if is_slide_header(line, prev_line):
            # Save previous section if it has content
            if current_section["content"]:
                if current_section["title"]:
                    structured_content.append(f"## {current_section['title']}")
                structured_content.extend(current_section["content"])
                structured_content.append("")  # Add blank line between sections
            
            # Start new section
            current_section = {"title": line, "content": []}
            continue
            
        # Add content to current section
        if is_bullet_point(line):
            # Format bullet points consistently
            bullet_match = re.match(r'^\s*([-•*]|\d+\.|[a-zA-Z]\)|\(\d+\))\s+(.+)$', line)
            if bullet_match:
                bullet, content = bullet_match.groups()
                current_section["content"].append(f"* {content}")
            else:
                current_section["content"].append(f"* {line}")
        else:
            current_section["content"].append(line)
    
    # Add the last section
    if current_section["content"]:
        if current_section["title"]:
            structured_content.append(f"## {current_section['title']}")
        structured_content.extend(current_section["content"])
    
    # Join everything back together
    return '\n'.join(structured_content)

def main():
    """Main function to run the preprocessing."""
    args = setup_argparse()
    
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    # Check NLP dependencies if needed
    if args.use_nlp and not check_dependencies(use_nlp=True):
        return 1
    
    # Read the input file
    logger.info(f"Reading input file: {args.input_file}")
    text = read_text_file(args.input_file)
    if text is None:
        return 1
    
    # Process the text
    logger.info("Processing text...")
    
    if args.aggressive:
        processed_text = aggressive_clean(text)
    elif args.use_nlp:
        processed_text = advanced_clean_with_nlp(text, args.preserve_bullets)
    else:
        processed_text = basic_clean(text, args.preserve_bullets)
    
    # Determine output file path
    if args.output:
        output_path = args.output
    else:
        input_path = Path(args.input_file)
        output_path = str(input_path.with_stem(input_path.stem + "_cleaned").with_suffix(input_path.suffix))
    
    # Write the processed text to the output file
    if write_text_file(output_path, processed_text):
        print(f"Successfully processed text and saved to: {output_path}")
        return 0
    else:
        print("Failed to save processed text")
        return 1

if __name__ == "__main__":
    sys.exit(main())
