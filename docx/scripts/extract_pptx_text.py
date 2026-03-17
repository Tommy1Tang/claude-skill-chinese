#!/usr/bin/env python3
"""
Extract text content from .pptx files with proper Chinese/Unicode support.

This script handles encoding detection for presentations created on different systems
(Windows GBK/GB18030, UTF-8, etc.) and outputs clean UTF-8 text.

Usage:
    python extract_pptx_text.py presentation.pptx
    python extract_pptx_text.py presentation.pptx -o output.txt
    python extract_pptx_text.py presentation.pptx --json  # Output as JSON with slide structure
"""

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# XML namespaces for PowerPoint presentations
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def detect_xml_encoding(xml_bytes: bytes) -> str:
    """
    Detect XML encoding from declaration or content analysis.

    Tries multiple encodings in order of likelihood for Chinese documents.
    """
    # First check for explicit encoding declaration
    header = xml_bytes[:200]
    try:
        header_str = header.decode('ascii', errors='ignore')
        if 'encoding="' in header_str:
            match = re.search(r'encoding="([^"]+)"', header_str)
            if match:
                declared = match.group(1)
                # Verify it works
                xml_bytes.decode(declared)
                return declared
    except:
        pass

    # Try encodings in order of likelihood
    encodings_to_try = ['utf-8', 'gb18030', 'gbk', 'gb2312', 'utf-16']
    for encoding in encodings_to_try:
        try:
            xml_bytes.decode(encoding)
            return encoding
        except:
            continue

    # Fallback to UTF-8 with error handling
    return 'utf-8'


def extract_text_from_slide(slide_xml: str) -> list:
    """
    Extract text from a single slide's XML content.

    Returns list of text strings found in the slide.
    """
    texts = []

    try:
        root = ET.fromstring(slide_xml)

        # Find all text elements (<a:t> is the text element in OOXML)
        for t_elem in root.findall('.//a:t', NAMESPACES):
            if t_elem.text and t_elem.text.strip():
                texts.append(t_elem.text)
    except ET.ParseError as e:
        # If parsing fails, try regex fallback
        texts = re.findall(r'<a:t[^>]*>([^<]+)</a:t>', slide_xml)
        texts = [t for t in texts if t.strip()]

    return texts


def extract_text_from_pptx(pptx_path: Path) -> dict:
    """
    Extract text from a .pptx file, organized by slide.

    Returns dict with slide numbers as keys and lists of text as values.
    """
    slides_content = {}

    with zipfile.ZipFile(pptx_path, 'r') as pptx:
        # Get all slide files and sort them
        slide_files = [f for f in pptx.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
        slide_files.sort()

        for slide_file in slide_files:
            # Extract slide number from filename
            slide_num = int(re.search(r'slide(\d+)\.xml', slide_file).group(1))

            # Read and decode slide XML
            xml_bytes = pptx.read(slide_file)
            encoding = detect_xml_encoding(xml_bytes)
            xml_content = xml_bytes.decode(encoding, errors='replace')

            # Extract text
            texts = extract_text_from_slide(xml_content)
            if texts:
                slides_content[slide_num] = texts

    return slides_content


def extract_with_metadata(pptx_path: Path) -> dict:
    """
    Extract text with metadata from a .pptx file.

    Returns a dict with slides and presentation properties.
    """
    result = {
        'file': str(pptx_path),
        'slides': [],
        'properties': {}
    }

    with zipfile.ZipFile(pptx_path, 'r') as pptx:
        # Extract slides
        slide_files = [f for f in pptx.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
        slide_files.sort()

        for slide_file in slide_files:
            slide_num = int(re.search(r'slide(\d+)\.xml', slide_file).group(1))

            xml_bytes = pptx.read(slide_file)
            encoding = detect_xml_encoding(xml_bytes)
            xml_content = xml_bytes.decode(encoding, errors='replace')

            texts = extract_text_from_slide(xml_content)
            if texts:
                result['slides'].append({
                    'slide_number': slide_num,
                    'content': texts
                })

        # Try to extract presentation properties
        try:
            if 'docProps/core.xml' in pptx.namelist():
                props_xml = pptx.read('docProps/core.xml')
                props_encoding = detect_xml_encoding(props_xml)
                props_content = props_xml.decode(props_encoding, errors='replace')
                props_root = ET.fromstring(props_content)

                props_ns = {'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                           'dc': 'http://purl.org/dc/elements/1.1/'}

                title = props_root.find('.//dc:title', props_ns)
                if title is not None and title.text:
                    result['properties']['title'] = title.text

                creator = props_root.find('.//dc:creator', props_ns)
                if creator is not None and creator.text:
                    result['properties']['creator'] = creator.text
        except:
            pass  # Properties are optional

    return result


def format_output(content: dict, format_type: str = 'text') -> str:
    """
    Format the extracted content for output.

    Args:
        content: Dict with slides content
        format_type: 'text' or 'markdown'

    Returns:
        Formatted string
    """
    lines = []

    if format_type == 'markdown':
        for slide_num, texts in sorted(content.items()):
            lines.append(f"## Slide {slide_num}")
            lines.append("")
            for text in texts:
                lines.append(text)
            lines.append("")
    else:
        for slide_num, texts in sorted(content.items()):
            lines.append(f"=== Slide {slide_num} ===")
            for text in texts:
                lines.append(text)
            lines.append("")

    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(
        description='Extract text from .pptx files with proper Chinese/Unicode support'
    )
    parser.add_argument('input', help='Input .pptx file')
    parser.add_argument('-o', '--output', help='Output file (default: stdout)')
    parser.add_argument('--json', action='store_true',
                       help='Output as JSON with slide structure')
    parser.add_argument('--markdown', action='store_true',
                       help='Output in Markdown format with slide headers')
    parser.add_argument('--metadata', action='store_true',
                       help='Include document metadata')

    args = parser.parse_args()

    pptx_path = Path(args.input)
    if not pptx_path.exists():
        print(f"Error: File not found: {pptx_path}", file=sys.stderr)
        sys.exit(1)

    try:
        if args.metadata or args.json:
            result = extract_with_metadata(pptx_path)
            if args.json:
                output = json.dumps(result, ensure_ascii=False, indent=2)
            else:
                # Format with metadata
                lines = []
                if result['properties']:
                    lines.append(f"# Presentation: {result['properties'].get('title', 'Untitled')}")
                    lines.append(f"# Author: {result['properties'].get('creator', 'Unknown')}")
                    lines.append("")
                for slide in result['slides']:
                    lines.append(f"=== Slide {slide['slide_number']} ===")
                    lines.extend(slide['content'])
                    lines.append("")
                output = '\n'.join(lines)
        else:
            content = extract_text_from_pptx(pptx_path)
            output = format_output(content, 'markdown' if args.markdown else 'text')

        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(output)
            print(f"Extracted text saved to: {args.output}")
        else:
            # Ensure stdout uses UTF-8
            if hasattr(sys.stdout, 'reconfigure'):
                sys.stdout.reconfigure(encoding='utf-8')
            print(output)

    except Exception as e:
        print(f"Error extracting text: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
