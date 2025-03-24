#!/usr/bin/env python3
# Enhanced PowerPoint Rebuilder with Content Awareness
# Builds new PowerPoint files with proper content and layout from extracted text

import os
import sys
import argparse
import logging
import zipfile
import tempfile
import shutil
from PIL import Image, ImageDraw, ImageFont
import io
import re
import xml.etree.ElementTree as ET
from pathlib import Path

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Register namespaces for PowerPoint XML
namespaces = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

for prefix, uri in namespaces.items():
    ET.register_namespace(prefix, uri)

def setup_argparse():
    """Setup command line argument parsing."""
    parser = argparse.ArgumentParser(description='Rebuild PowerPoint files with proper content and layout.')
    parser.add_argument('input_file', type=str, help='Path to the corrupted PPTX file')
    parser.add_argument('--output-dir', '-o', type=str, default='rebuilt_pptx', 
                        help='Directory to save output files (default: rebuilt_pptx)')
    parser.add_argument('--max-slides', '-m', type=int, default=100, 
                        help='Maximum number of slides total (default: 100)')
    parser.add_argument('--slides-per-file', '-s', type=int, default=15, 
                        help='Maximum number of slides per file (default: 15)')
    parser.add_argument('--max-files', '-f', type=int, default=10, 
                        help='Maximum number of files to create (default: 10)')
    parser.add_argument('--extract-dir', '-e', type=str, default=None,
                        help='Directory to extract content (default: temporary directory)')
    parser.add_argument('--text-file', '-t', type=str, default=None,
                        help='Path to text file with extracted content (if already available)')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose output')
    return parser.parse_args()

def extract_images_from_binary(file_path, extract_dir):
    """Extract images from binary PPTX file."""
    logger.info(f"Extracting images from {file_path}")
    
    # Create media directory
    media_dir = os.path.join(extract_dir, 'ppt', 'media')
    os.makedirs(media_dir, exist_ok=True)
    
    with open(file_path, 'rb') as f:
        data = f.read()
    
    # Extract PNG images
    png_header = b'\x89PNG\r\n\x1a\n'
    pos = 0
    png_count = 0
    
    while True:
        pos = data.find(png_header, pos)
        if pos == -1:
            break
        
        # Find IEND chunk
        iend_marker = b'IEND\xaeB\x60\x82'
        end_pos = data.find(iend_marker, pos)
        
        if end_pos != -1:
            end_pos += len(iend_marker)
            image_data = data[pos:end_pos]
            
            # Validate it's a real PNG
            try:
                Image.open(io.BytesIO(image_data))
                with open(os.path.join(media_dir, f"image_{png_count+1}.png"), 'wb') as img_file:
                    img_file.write(image_data)
                png_count += 1
            except:
                pass
            
        pos += 1
    
    # Extract JPEG images
    jpg_header = b'\xff\xd8\xff'
    pos = 0
    jpg_count = 0
    
    while True:
        pos = data.find(jpg_header, pos)
        if pos == -1:
            break
        
        # Find EOI marker
        eoi_marker = b'\xff\xd9'
        end_pos = data.find(eoi_marker, pos)
        
        if end_pos != -1:
            end_pos += len(eoi_marker)
            image_data = data[pos:end_pos]
            
            # Validate it's a real JPEG
            try:
                Image.open(io.BytesIO(image_data))
                with open(os.path.join(media_dir, f"image_{png_count+jpg_count+1}.jpg"), 'wb') as img_file:
                    img_file.write(image_data)
                jpg_count += 1
            except:
                pass
            
        pos += 1
    
    logger.info(f"Extracted {png_count} PNG and {jpg_count} JPEG images")
    return png_count + jpg_count

def extract_text_from_binary(file_path):
    """Extract text segments from binary PPTX file."""
    logger.info(f"Extracting text from {file_path}")
    
    with open(file_path, 'rb') as f:
        data = f.read()
    
    # Convert to text, ignoring errors
    text_data = data.decode('utf-8', errors='ignore')
    
    # Look for text in PowerPoint XML
    text_content = []
    
    # Pattern for text inside a:t tags
    pattern = r'<a:t[^>]*>(.*?)</a:t>'
    matches = re.findall(pattern, text_data, re.DOTALL)
    
    for match in matches:
        # Clean up the text
        text = match.strip()
        
        # Skip if too short or looks like code
        if len(text) < 3 or text.startswith('<?xml'):
            continue
            
        text_content.append(text)
    
    # Also look for readable text chunks
    plain_text_pattern = r'[A-Za-z0-9\s.,;:?!\'\"\-_&]{10,}'
    plain_matches = re.findall(plain_text_pattern, text_data)
    
    for match in plain_matches:
        if len(match.strip()) >= 10:
            text_content.append(match.strip())
    
    # Remove duplicates
    unique_text = []
    seen = set()
    
    for text in text_content:
        if text not in seen:
            seen.add(text)
            unique_text.append(text)
    
    logger.info(f"Extracted {len(unique_text)} text segments")
    return unique_text

def load_extracted_text_file(text_file_path):
    """Load text content from a previously extracted text file."""
    logger.info(f"Loading text from {text_file_path}")
    
    try:
        with open(text_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        # Try different encodings if UTF-8 fails
        encodings = ['latin-1', 'cp1252', 'iso-8859-1']
        for encoding in encodings:
            try:
                with open(text_file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                logger.info(f"Successfully read file using {encoding} encoding")
                break
            except UnicodeDecodeError:
                continue
    
    # Split content into paragraphs
    paragraphs = re.split(r'\n\s*\n', content)
    
    # Filter out empty paragraphs
    paragraphs = [p.strip() for p in paragraphs if p.strip()]
    
    logger.info(f"Loaded {len(paragraphs)} paragraphs from text file")
    return paragraphs

def is_likely_title(text):
    """Check if text is likely a slide title."""
    # Title patterns:
    # - Starts with "Slide" followed by a number
    # - Short text (less than 60 chars)
    # - All uppercase or first letter of each word is uppercase
    # - Ends with a colon
    
    if re.match(r'^Slide\s+\d+', text, re.IGNORECASE):
        return True
        
    if len(text) < 60:
        # Ends with colon
        if text.rstrip().endswith(':'):
            return True
            
        # All uppercase
        if text.isupper():
            return True
            
        # Title case (most words start with uppercase)
        words = text.split()
        if len(words) >= 2:
            uppercase_words = sum(1 for word in words if word and word[0].isupper())
            if uppercase_words / len(words) >= 0.7:
                return True
                
    return False

def organize_text_into_slides(paragraphs, max_slides=100):
    """Organize text paragraphs into slides with logical grouping."""
    slides_content = {}
    current_slide = 1
    
    # Filter out very short paragraphs and duplicates
    filtered_paragraphs = []
    seen = set()
    
    for para in paragraphs:
        # Skip very short content (likely noise)
        if len(para.strip()) < 5:
            continue
            
        # Skip duplicates
        if para in seen:
            continue
            
        seen.add(para)
        filtered_paragraphs.append(para)
    
    # Limit total number of paragraphs to avoid too many slides
    max_paragraphs = max_slides * 8  # Assume max 8 paragraphs per slide
    if len(filtered_paragraphs) > max_paragraphs:
        filtered_paragraphs = filtered_paragraphs[:max_paragraphs]
        
    paragraphs = filtered_paragraphs
    
    # Initial pass to identify potential titles
    title_indices = []
    for i, para in enumerate(paragraphs):
        if is_likely_title(para) or (i > 0 and len(para) < 50 and len(paragraphs[i-1]) < 3):
            title_indices.append(i)
    
    # If very few titles found, add more potential break points
    if len(title_indices) < max_slides // 2:
        for i, para in enumerate(paragraphs):
            # Add more potential break points based on content patterns
            if para.strip().endswith(':') or para.strip().endswith('.'):
                if i not in title_indices:
                    title_indices.append(i)
    
    # Sort indices to ensure correct order
    title_indices.sort()
    
    # If still not enough break points, use a fixed number of paragraphs per slide
    if not title_indices or len(title_indices) < 5:
        # Evenly distribute paragraphs
        paras_per_slide = max(1, min(8, len(paragraphs) // max(1, min(max_slides, 15))))
        
        for i in range(0, len(paragraphs), paras_per_slide):
            slide_paras = paragraphs[i:i+paras_per_slide]
            
            if not slide_paras:  # Skip if no content
                continue
                
            # Try to identify a title in the first paragraph
            title = slide_paras[0] if slide_paras and len(slide_paras[0]) < 60 else f"Slide {current_slide}"
            content = slide_paras[1:] if slide_paras[0] == title else slide_paras
            
            # Skip if no meaningful content
            if not content and title == f"Slide {current_slide}":
                continue
                
            slides_content[current_slide] = {
                'title': title,
                'content': content,
                'images': []
            }
            current_slide += 1
    else:
        # Use identified titles to separate slides
        for i in range(len(title_indices)):
            start_idx = title_indices[i]
            end_idx = title_indices[i+1] if i+1 < len(title_indices) else len(paragraphs)
            
            title = paragraphs[start_idx]
            content = paragraphs[start_idx+1:end_idx]
            
            # Skip slides with no content
            if not content:
                continue
                
            slides_content[current_slide] = {
                'title': title,
                'content': content,
                'images': []
            }
            current_slide += 1
    
    # Limit to max_slides
    if len(slides_content) > max_slides:
        # Keep only the first max_slides
        slides_to_keep = sorted(slides_content.keys())[:max_slides]
        slides_content = {k: slides_content[k] for k in slides_to_keep}
    
    logger.info(f"Organized content into {len(slides_content)} slides")
    return slides_content

def distribute_images(slides_content, image_count):
    """Distribute images across slides."""
    logger.info(f"Distributing {image_count} images across {len(slides_content)} slides")
    
    # Get list of slide numbers
    slide_numbers = sorted(slides_content.keys())
    
    # Distribute images evenly across slides
    for i in range(image_count):
        img_num = i + 1
        
        # Calculate which slide gets this image (distribute evenly)
        slide_idx = i % len(slide_numbers)
        slide_num = slide_numbers[slide_idx]
        
        # Determine image type based on the pattern
        img_file = f"image_{img_num}.png"  # Default to PNG
        
        # Add to slide content
        slides_content[slide_num]['images'].append(img_file)
    
    return slides_content

def create_placeholder_slide(slide_num, title, content_paragraphs, images):
    """Create a PowerPoint slide XML with text and image placeholders."""
    # Sanitize the title and limit its length
    title = title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    if len(title) > 100:
        title = title[:97] + '...'
    
    # Create the slide XML
    slide_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1" name=""/>
            <p:cNvGrpSpPr/>
            <p:nvPr/>
          </p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="0" cy="0"/>
              <a:chOff x="0" y="0"/>
              <a:chExt cx="0" cy="0"/>
            </a:xfrm>
          </p:grpSpPr>
          
          <!-- Slide title -->
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title"/>
              <p:cNvSpPr/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="1388800" y="571500"/>
                <a:ext cx="8636000" cy="762000"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" dirty="0" smtClean="0"/>
                  <a:t>{title}</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>"""
    
    # Add text shapes
    max_content_paragraphs = min(8, len(content_paragraphs))  # Limit to 8 paragraphs per slide
    for i, para in enumerate(content_paragraphs[:max_content_paragraphs]):
        y_pos = 1500000 + (i * 600000)  # Position text boxes vertically
        
        # Sanitize text
        text = para.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        if len(text) > 1000:
            text = text[:997] + '...'  # Limit text length
            
        slide_xml += f"""
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="{i+3}" name="Text {i+1}"/>
              <p:cNvSpPr/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="1388800" y="{y_pos}"/>
                <a:ext cx="8636000" cy="600000"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" dirty="0" smtClean="0"/>
                  <a:t>{text}</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>"""
    
    # Add image placeholders
    max_images = min(4, len(images))  # Limit to 4 images per slide
    for i, img_file in enumerate(images[:max_images]):
        # Position images in a grid (2x2)
        row = i // 2
        col = i % 2
        
        x_pos = 1388800 + (col * 4000000)  # X position based on column
        y_pos = 5000000 + (row * 2000000)  # Y position based on row
        
        rid = i + 1  # Relationship ID
        
        slide_xml += f"""
          <p:pic>
            <p:nvPicPr>
              <p:cNvPr id="{i+20}" name="Picture {i+1}"/>
              <p:cNvPicPr/>
              <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="rId{rid}"/>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="{x_pos}" y="{y_pos}"/>
                <a:ext cx="3000000" cy="2000000"/>
              </a:xfrm>
            </p:spPr>
          </p:pic>"""
    
    # Close the slide XML
    slide_xml += """
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sld>"""
    
    # Create the slide relationship XML
    slide_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""
    
    # Add image relationships
    for i, img_file in enumerate(images[:max_images]):
        rid = i + 1
        
        # Determine file extension
        ext = "png"  # Default
        if img_file.lower().endswith(".jpg") or img_file.lower().endswith(".jpeg"):
            ext = "jpg"
            
        slide_rels_xml += f"""
        <Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{img_file}"/>"""
    
    # Close the relationships XML
    slide_rels_xml += """
    </Relationships>"""
    
    return slide_xml, slide_rels_xml

def create_pptx_structure(extract_dir, slides_content, part_num, start_slide, end_slide):
    """Create a minimal PPTX structure with slides."""
    logger.info(f"Creating PPTX part {part_num} with slides {start_slide}-{end_slide}")
    
    # Create basic directory structure
    pptx_dir = os.path.join(extract_dir, f"part_{part_num}")
    os.makedirs(os.path.join(pptx_dir, "ppt", "slides", "_rels"), exist_ok=True)
    os.makedirs(os.path.join(pptx_dir, "ppt", "media"), exist_ok=True)
    os.makedirs(os.path.join(pptx_dir, "ppt", "_rels"), exist_ok=True)
    os.makedirs(os.path.join(pptx_dir, "_rels"), exist_ok=True)
    
    # Create content types XML
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="jpeg" ContentType="image/jpeg"/>
        <Default Extension="jpg" ContentType="image/jpeg"/>
        <Default Extension="png" ContentType="image/png"/>
        <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"""
    
    # Add slide content types
    for i, slide_num in enumerate(range(start_slide, end_slide + 1)):
        content_types += f"""
        <Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"""
    
    content_types += """
    </Types>"""
    
    with open(os.path.join(pptx_dir, "[Content_Types].xml"), "w", encoding="utf-8") as f:
        f.write(content_types)
    
    # Create presentation.xml
    presentation_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:sldIdLst>"""
    
    # Add slide references
    for i, slide_num in enumerate(range(start_slide, end_slide + 1)):
        presentation_xml += f"""
            <p:sldId id="{256 + i}" r:id="rId{1000 + i}"/>"""
    
    presentation_xml += """
        </p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000" type="screen4x3"/>
        <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>"""
    
    with open(os.path.join(pptx_dir, "ppt", "presentation.xml"), "w", encoding="utf-8") as f:
        f.write(presentation_xml)
    
    # Create presentation relationships
    pres_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""
    
    # Add slide relationships
    for i, slide_num in enumerate(range(start_slide, end_slide + 1)):
        pres_rels += f"""
        <Relationship Id="rId{1000 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i + 1}.xml"/>"""
    
    pres_rels += """
    </Relationships>"""
    
    with open(os.path.join(pptx_dir, "ppt", "_rels", "presentation.xml.rels"), "w", encoding="utf-8") as f:
        f.write(pres_rels)
    
    # Create root relationships
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    </Relationships>"""
    
    with open(os.path.join(pptx_dir, "_rels", ".rels"), "w", encoding="utf-8") as f:
        f.write(root_rels)
    
    # Create slides and their relationships
    all_used_images = []
    
    # Get the subset of slides we're working with
    current_slides = {num: slides_content[num] for num in range(start_slide, end_slide + 1) if num in slides_content}
    
    for i, slide_num in enumerate(range(start_slide, end_slide + 1)):
        if slide_num not in current_slides:
            # Create an empty slide if we don't have content
            title = f"Recovered Slide {slide_num}"
            content = []
            images = []
        else:
            slide_data = current_slides[slide_num]
            title = slide_data['title']
            content = slide_data['content']
            images = slide_data['images']
        
        # Create slide XML and relationships
        slide_xml, slide_rels_xml = create_placeholder_slide(
            slide_num, 
            title, 
            content, 
            images
        )
        
        # Save slide XML
        with open(os.path.join(pptx_dir, "ppt", "slides", f"slide{i + 1}.xml"), "w", encoding="utf-8") as f:
            f.write(slide_xml)
        
        # Save slide relationships
        with open(os.path.join(pptx_dir, "ppt", "slides", "_rels", f"slide{i + 1}.xml.rels"), "w", encoding="utf-8") as f:
            f.write(slide_rels_xml)
        
        # Track images used
        all_used_images.extend(images)
    
    # Copy used images
    for img_file in set(all_used_images):
        src_path = os.path.join(extract_dir, "ppt", "media", img_file)
        dst_path = os.path.join(pptx_dir, "ppt", "media", img_file)
        
        # Only copy if source exists
        if os.path.exists(src_path):
            shutil.copy2(src_path, dst_path)
        else:
            # Create a placeholder image
            create_placeholder_image(dst_path, img_file)
    
    logger.info(f"Created PPTX structure for part {part_num}")
    return pptx_dir

def create_placeholder_image(path, img_name):
    """Create a placeholder image when the original is missing."""
    # Create a simple colored placeholder
    img = Image.new('RGB', (800, 600), color=(220, 220, 240))
    draw = ImageDraw.Draw(img)
    
    # Add text about the missing image
    try:
        # Try to use a default font
        font = ImageFont.load_default()
        draw.text((50, 50), f"Placeholder for: {img_name}", fill=(0, 0, 0), font=font)
        draw.text((50, 100), "Original image could not be recovered", fill=(0, 0, 0), font=font)
        draw.text((50, 150), "This is a generated placeholder", fill=(0, 0, 0), font=font)
    except:
        # Fallback if font issues
        draw.text((50, 50), f"Placeholder for: {img_name}", fill=(0, 0, 0))
    
    # Draw border
    draw.rectangle([(0, 0), (799, 599)], outline=(0, 0, 0), width=2)
    
    # Save the image
    img.save(path)
    logger.info(f"Created placeholder image for {img_name}")

def create_pptx(src_dir, output_path):
    """Create a PPTX file from a directory structure."""
    logger.info(f"Creating PPTX file: {output_path}")
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for root, dirs, files in os.walk(src_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arc_name = os.path.relpath(file_path, src_dir)
                zip_file.write(file_path, arc_name)
    
    logger.info(f"Created PPTX file: {output_path}")
    return True

def main():
    args = setup_argparse()
    
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    # Create output directory
    os.makedirs(args.output_dir, exist_ok=True)
    
    # Create or use extract directory
    if args.extract_dir:
        extract_dir = args.extract_dir
        os.makedirs(extract_dir, exist_ok=True)
    else:
        extract_dir = tempfile.mkdtemp()
    
    try:
        # Extract images from the PPTX
        image_count = extract_images_from_binary(args.input_file, extract_dir)
        
        # Get text content
        if args.text_file and os.path.exists(args.text_file):
            # Load from provided text file
            text_paragraphs = load_extracted_text_file(args.text_file)
        else:
            # Extract text from the PPTX
            extracted_text = extract_text_from_binary(args.input_file)
            text_paragraphs = extracted_text
        
        # Organize text into slides with absolute limit
        max_slides = min(100, args.max_slides)
        slides_content = organize_text_into_slides(text_paragraphs, max_slides)
        
        # Distribute images across slides
        slides_content = distribute_images(slides_content, image_count)
        
        # Limit to a reasonable number of slides and files
        slide_numbers = sorted(slides_content.keys())
        total_slides = len(slide_numbers)
        
        # Ensure we have content for slides
        if total_slides == 0:
            logger.error("No slide content could be extracted. Try providing a text file with --text-file.")
            return 1
            
        # Calculate optimal distribution
        slides_per_file = min(args.slides_per_file, total_slides)
        desired_files = min(args.max_files, (total_slides + slides_per_file - 1) // slides_per_file)
        
        # Recalculate based on desired number of files
        slides_per_file = (total_slides + desired_files - 1) // desired_files
        num_files = (total_slides + slides_per_file - 1) // slides_per_file
        
        logger.info(f"Distributing {total_slides} slides across {num_files} files (approx. {slides_per_file} per file)")
        
        logger.info(f"Creating {num_files} PowerPoint files with {total_slides} total slides")
        
        # Create separate PPTX files
        for i in range(num_files):
            file_slide_numbers = slide_numbers[i * slides_per_file:min((i+1) * slides_per_file, total_slides)]
            
            if not file_slide_numbers:
                continue
                
            start_slide = file_slide_numbers[0]
            end_slide = file_slide_numbers[-1]
            
            # Create mapping for continuous slide numbers in the output
            slide_mapping = {original: i+1 for i, original in enumerate(file_slide_numbers)}
            
            # Create filtered content for this file
            file_content = {slide_mapping[num]: slides_content[num] for num in file_slide_numbers}
            
            # Create PPTX structure with remapped slide numbers
            pptx_dir = create_pptx_structure(
                extract_dir, 
                file_content, 
                i + 1, 
                1,  # Always start at 1 in each file
                len(file_slide_numbers)
            )
            
            # Create PPTX file
            output_file = os.path.join(args.output_dir, f"rebuilt_part_{i+1}.pptx")
            if not create_pptx(pptx_dir, output_file):
                logger.error(f"Failed to create file: {output_file}")
                continue
            
            # Cleanup directory used for this part
            if not args.extract_dir:  # Only cleanup if using temp dir
                shutil.rmtree(pptx_dir, ignore_errors=True)
        
        logger.info(f"Created {num_files} PowerPoint files in {args.output_dir}")
        print(f"Successfully created {num_files} PowerPoint files in {args.output_dir}")
        
        return 0
    finally:
        # Cleanup if using temp dir
        if not args.extract_dir and os.path.exists(extract_dir):
            shutil.rmtree(extract_dir, ignore_errors=True)

if __name__ == "__main__":
    sys.exit(main())