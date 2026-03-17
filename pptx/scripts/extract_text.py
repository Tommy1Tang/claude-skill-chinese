#!/usr/bin/env python3
"""
Extract text content from .docx files with proper Chinese/Unicode support.

This script handles encoding detection for documents created on different systems
(Windows GBK/GB18030, UTF-8, etc.) and outputs clean UTF-8 text.

Usage:
    python extract_text.py document.docx
    python extract_text.py document.docx -o output.txt
    python extract_text.py document.docx --json  # Output as JSON with paragraph structure
"""

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# XML namespaces for Word documents
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
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


def extract_text_from_docx(docx_path: Path, preserve_structure: bool = False) -> list:
    """
    Extract text from a .docx file.

    Args:
        docx_path: Path to the .docx file
        preserve_structure: If True, return list of paragraphs; if False, return flat text

    Returns:
        List of text strings (paragraphs) or structured dict if preserve_structure=True
    """
    paragraphs = []

    with zipfile.ZipFile(docx_path, 'r') as docx:
        # Check if document.xml exists
        doc_xml_path = 'word/document.xml'
        if doc_xml_path not in docx.namelist():
            raise ValueError(f"Invalid .docx file: {docx_path} (missing {doc_xml_path})")

        # Read and decode XML content
        xml_bytes = docx.read(doc_xml_path)
        encoding = detect_xml_encoding(xml_bytes)
        xml_content = xml_bytes.decode(encoding, errors='replace')

        # Register namespaces
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

        # Parse XML
        root = ET.fromstring(xml_content)

        # Extract text from paragraphs
        for para in root.findall('.//w:p', NAMESPACES):
            texts = []
            for t_elem in para.findall('.//w:t', NAMESPACES):
                if t_elem.text:
                    texts.append(t_elem.text)

            para_text = ''.join(texts)
            if para_text.strip():  # Only include non-empty paragraphs
                paragraphs.append(para_text)

    return paragraphs


def extract_with_metadata(docx_path: Path) -> dict:
    """
    Extract text with metadata from a .docx file.

    Returns a dict with paragraphs and document properties.
    """
    result = {
        'file': str(docx_path),
        'paragraphs': [],
        'properties': {}
    }

    with zipfile.ZipFile(docx_path, 'r') as docx:
        # Extract main text
        xml_bytes = docx.read('word/document.xml')
        encoding = detect_xml_encoding(xml_bytes)
        xml_content = xml_bytes.decode(encoding, errors='replace')

        root = ET.fromstring(xml_content)

        for para in root.findall('.//w:p', NAMESPACES):
            texts = []
            for t_elem in para.findall('.//w:t', NAMESPACES):
                if t_elem.text:
                    texts.append(t_elem.text)
            para_text = ''.join(texts)
            if para_text.strip():
                result['paragraphs'].append(para_text)

        # Try to extract document properties
        try:
            if 'docProps/core.xml' in docx.namelist():
                props_xml = docx.read('docProps/core.xml')
                props_encoding = detect_xml_encoding(props_xml)
                props_content = props_xml.decode(props_encoding, errors='replace')
                props_root = ET.fromstring(props_content)

                # Extract common properties
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


def main():
    parser = argparse.ArgumentParser(
        description='Extract text from .docx files with proper Chinese/Unicode support'
    )
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('-o', '--output', help='Output file (default: stdout)')
    parser.add_argument('--json', action='store_true',
                       help='Output as JSON with paragraph structure')
    parser.add_argument('--metadata', action='store_true',
                       help='Include document metadata')

    args = parser.parse_args()

    docx_path = Path(args.input)
    if not docx_path.exists():
        print(f"Error: File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    try:
        if args.metadata or args.json:
            result = extract_with_metadata(docx_path)
            if args.json:
                output = json.dumps(result, ensure_ascii=False, indent=2)
            else:
                # Just output paragraphs with metadata header
                lines = []
                if result['properties']:
                    lines.append(f"# Document: {result['properties'].get('title', 'Untitled')}")
                    lines.append(f"# Author: {result['properties'].get('creator', 'Unknown')}")
                    lines.append("")
                lines.extend(result['paragraphs'])
                output = '\n'.join(lines)
        else:
            paragraphs = extract_text_from_docx(docx_path)
            output = '\n'.join(paragraphs)

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
