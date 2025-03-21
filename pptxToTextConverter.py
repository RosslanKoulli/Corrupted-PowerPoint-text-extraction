#!/usr/bin/env python3
# Corrupted PDF Text Extractor
# This script attempts to extract text from potentially corrupted PDF files
# using multiple methods for better recovery chances.

import sys
import os
import argparse
import traceback
from pathlib import Path
import logging

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def setup_argparse():
    """Setup command line argument parsing."""
    parser = argparse.ArgumentParser(description='Extract text from corrupted PDF files.')
    parser.add_argument('input_pdf', type=str, help='Path to the potentially corrupted PDF file')
    parser.add_argument('--output', '-o', type=str, help='Path to the output text file (default: same name as PDF with .txt extension)')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose output')
    parser.add_argument('--work-dir', '-w', type=str, help='Directory to use for temporary files (default: current directory)')
    return parser.parse_args()

def extract_with_pdfminer(pdf_path):
    """Attempt extraction using PDFMiner."""
    try:
        from pdfminer.high_level import extract_text
        logger.info("Attempting extraction with PDFMiner...")
        text = extract_text(pdf_path)
        if text.strip():
            logger.info("Successfully extracted text with PDFMiner")
            return text
        else:
            logger.warning("PDFMiner returned empty text")
            return None
    except Exception as e:
        logger.warning(f"PDFMiner extraction failed: {str(e)}")
        return None

def extract_with_pypdf(pdf_path):
    """Attempt extraction using PyPDF."""
    try:
        import pypdf
        logger.info("Attempting extraction with PyPDF...")
        text = ""
        with open(pdf_path, 'rb') as file:
            try:
                pdf_reader = pypdf.PdfReader(file)
                for page_num in range(len(pdf_reader.pages)):
                    try:
                        page = pdf_reader.pages[page_num]
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n\n"
                    except Exception as e:
                        logger.warning(f"Error extracting page {page_num}: {str(e)}")
                        continue
            except Exception as e:
                logger.warning(f"PyPDF reader initialization failed: {str(e)}")
                return None
                
        if text.strip():
            logger.info("Successfully extracted text with PyPDF")
            return text
        else:
            logger.warning("PyPDF returned empty text")
            return None
    except Exception as e:
        logger.warning(f"PyPDF extraction failed: {str(e)}")
        return None

def extract_with_tika(pdf_path):
    """Attempt extraction using Apache Tika."""
    try:
        from tika import parser
        logger.info("Attempting extraction with Apache Tika...")
        parsed = parser.from_file(pdf_path)
        if parsed.get("content"):
            logger.info("Successfully extracted text with Apache Tika")
            return parsed.get("content")
        else:
            logger.warning("Tika returned empty content")
            return None
    except Exception as e:
        logger.warning(f"Tika extraction failed: {str(e)}")
        return None

def extract_with_ocrmypdf(pdf_path, work_dir):
    """Attempt repair and extraction using OCRmyPDF."""
    try:
        import ocrmypdf
        import shutil
        
        logger.info("Attempting repair with OCRmyPDF...")
        temp_output = os.path.join(work_dir, "repaired.pdf")
        
        # Try to repair the PDF with OCRmyPDF
        try:
            ocrmypdf.ocr(
                pdf_path, 
                temp_output,
                force_ocr=True,
                skip_text=True,
                redo_ocr=True,
                use_threads=True,
                clean=True
            )
            loggr.info("PDF successfully repaired with OCRmyPDF")
            
            # Now extract text from the repaired PDF
            result = extract_with_pypdf(temp_output) or extract_with_pdfminer(temp_output)
            
            # Clean up
            try:
                os.remove(temp_output)
            except:
                pass
                
            return result
        
        except Exception as e:
            logger.warning(f"OCRmyPDF repair failed: {str(e)}")
            return None
            
    except ImportError:
        logger.warning("OCRmyPDF not installed, skipping this method")
        return None

def extract_with_tesseract_pdf(pdf_path, work_dir):
    """Convert PDF to images and use Tesseract OCR."""
    try:
        import pdf2image
        import pytesseract
        from PIL import Image
        
        logger.info("Attempting extraction via PDF-to-Image conversion and Tesseract OCR...")
        
        # Convert PDF to images
        try:
            images = pdf2image.convert_from_path(pdf_path)
        except Exception as e:
            logger.warning(f"Failed to convert PDF to images: {str(e)}")
            return None
            
        if not images:
            logger.warning("No images extracted from PDF")
            return None
            
        full_text = ""
        for i, img in enumerate(images):
            try:
                # Save the image temporarily
                img_path = os.path.join(work_dir, f"page_{i+1}.png")
                img.save(img_path)
                
                logger.info(f"Processing page {i+1}/{len(images)} with Tesseract...")
                page_text = pytesseract.image_to_string(img_path)
                full_text += page_text + "\n\n"
                
                # Clean up
                try:
                    os.remove(img_path)
                except:
                    pass
                    
            except Exception as e:
                logger.warning(f"Failed OCR on page {i+1}: {str(e)}")
                continue
                
        if full_text.strip():
            logger.info("Successfully extracted text with Tesseract OCR")
            return full_text
        else:
            logger.warning("Tesseract OCR returned empty text")
            return None
            
    except ImportError as e:
        logger.warning(f"Required library not installed for Tesseract method: {str(e)}")
        return None

def check_dependencies():
    """Check which extraction libraries are available."""
    available_methods = []
    
    try:
        import pdfminer
        available_methods.append("pdfminer")
    except ImportError:
        logger.warning("PDFMiner not installed")
    
    try:
        import pypdf
        available_methods.append("pypdf")
    except ImportError:
        logger.warning("PyPDF not installed")
    
    try:
        import tika
        available_methods.append("tika")
    except ImportError:
        logger.warning("Apache Tika not installed")
    
    try:
        import ocrmypdf
        available_methods.append("ocrmypdf")
    except ImportError:
        logger.warning("OCRmyPDF not installed")
    
    try:
        import pdf2image
        import pytesseract
        available_methods.append("tesseract")
    except ImportError:
        logger.warning("PDF2Image or Pytesseract not installed")
    
    if not available_methods:
        logger.error("No extraction libraries available. Please install at least one of: "
                    "pdfminer.six, pypdf, tika-python, ocrmypdf, or pdf2image+pytesseract")
        sys.exit(1)
    
    return available_methods

def extract_text_from_pdf(pdf_path, verbose=False, work_dir=None):
    """Extract text from a potentially corrupted PDF using multiple methods."""
    if verbose:
        logger.setLevel(logging.DEBUG)
    
    if not os.path.exists(pdf_path):
        logger.error(f"File not found: {pdf_path}")
        return None
    
    # Use current directory if no working directory specified
    if work_dir is None:
        work_dir = os.getcwd()
    
    # Ensure work_dir exists
    os.makedirs(work_dir, exist_ok=True)
    
    available_methods = check_dependencies()
    logger.info(f"Available extraction methods: {', '.join(available_methods)}")
    
    # Try each method in sequence until one succeeds
    extracted_text = None
    
    if "pypdf" in available_methods:
        extracted_text = extract_with_pypdf(pdf_path)
        if extracted_text:
            return extracted_text
    
    if "pdfminer" in available_methods:
        extracted_text = extract_with_pdfminer(pdf_path)
        if extracted_text:
            return extracted_text
    
    if "tika" in available_methods:
        extracted_text = extract_with_tika(pdf_path)
        if extracted_text:
            return extracted_text
    
    if "ocrmypdf" in available_methods:
        extracted_text = extract_with_ocrmypdf(pdf_path, work_dir)
        if extracted_text:
            return extracted_text
    
    if "tesseract" in available_methods:
        extracted_text = extract_with_tesseract_pdf(pdf_path, work_dir)
        if extracted_text:
            return extracted_text
    
    logger.error("All extraction methods failed")
    return None

def save_text_to_file(text, output_path):
    """Save extracted text to a file."""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        logger.info(f"Text successfully saved to: {output_path}")
        return True
    except Exception as e:
        logger.error(f"Failed to save text: {str(e)}")
        return False

def main():
    """Main function to run the extraction process."""
    args = setup_argparse()
    
    # Determine output file path
    if args.output:
        output_path = args.output
    else:
        pdf_path = Path(args.input_pdf)
        output_path = str(pdf_path.with_suffix('.txt'))
    
    # Extract and save the text
    logger.info(f"Processing PDF: {args.input_pdf}")
    text = extract_text_from_pdf(args.input_pdf, args.verbose, args.work_dir)
    
    if text:
        if save_text_to_file(text, output_path):
            print(f"Text successfully extracted and saved to: {output_path}")
            return 0
        else:
            print("Failed to save extracted text")
            return 1
    else:
        print("Failed to extract text from the PDF")
        return 1

if __name__ == "__main__":
    sys.exit(main())
