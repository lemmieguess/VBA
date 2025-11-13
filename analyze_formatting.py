#!/usr/bin/env python3
"""
Analyze Word document formatting and compare two documents.
"""
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import sys

# Word XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

def twips_to_inches(twips):
    """Convert twips to inches (1 inch = 1440 twips)"""
    if twips is None:
        return None
    return float(twips) / 1440

def twips_to_points(twips):
    """Convert twips to points (1 point = 20 twips)"""
    if twips is None:
        return None
    return float(twips) / 20

def half_points_to_points(half_pts):
    """Convert half-points to points"""
    if half_pts is None:
        return None
    return float(half_pts) / 2

def analyze_paragraph_formatting(para):
    """Extract paragraph formatting properties"""
    pPr = para.find('w:pPr', NAMESPACES)
    if pPr is None:
        return {}

    formatting = {}

    # Spacing (line spacing, before, after)
    spacing = pPr.find('w:spacing', NAMESPACES)
    if spacing is not None:
        if 'w:line' in spacing.attrib:
            formatting['line_spacing_twips'] = spacing.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line']
            formatting['line_spacing_points'] = twips_to_points(formatting['line_spacing_twips'])
        if 'w:lineRule' in spacing.attrib:
            formatting['line_spacing_rule'] = spacing.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule']
        if 'w:before' in spacing.attrib:
            formatting['space_before_twips'] = spacing.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before']
            formatting['space_before_points'] = twips_to_points(formatting['space_before_twips'])
        if 'w:after' in spacing.attrib:
            formatting['space_after_twips'] = spacing.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after']
            formatting['space_after_points'] = twips_to_points(formatting['space_after_twips'])

    # Indentation
    ind = pPr.find('w:ind', NAMESPACES)
    if ind is not None:
        if 'w:left' in ind.attrib:
            formatting['left_indent_twips'] = ind.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left']
            formatting['left_indent_inches'] = twips_to_inches(formatting['left_indent_twips'])
        if 'w:right' in ind.attrib:
            formatting['right_indent_twips'] = ind.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right']
            formatting['right_indent_inches'] = twips_to_inches(formatting['right_indent_twips'])
        if 'w:hanging' in ind.attrib:
            formatting['hanging_indent_twips'] = ind.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hanging']
            formatting['hanging_indent_inches'] = twips_to_inches(formatting['hanging_indent_twips'])
        if 'w:firstLine' in ind.attrib:
            formatting['first_line_indent_twips'] = ind.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine']
            formatting['first_line_indent_inches'] = twips_to_inches(formatting['first_line_indent_twips'])

    # Numbering (bullets/numbering)
    numPr = pPr.find('w:numPr', NAMESPACES)
    if numPr is not None:
        formatting['has_numbering'] = True
        ilvl = numPr.find('w:ilvl', NAMESPACES)
        numId = numPr.find('w:numId', NAMESPACES)
        if ilvl is not None:
            formatting['numbering_level'] = ilvl.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
        if numId is not None:
            formatting['numbering_id'] = numId.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']

    # Alignment
    jc = pPr.find('w:jc', NAMESPACES)
    if jc is not None:
        formatting['alignment'] = jc.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']

    return formatting

def analyze_run_formatting(run):
    """Extract run (text) formatting properties"""
    rPr = run.find('w:rPr', NAMESPACES)
    if rPr is None:
        return {}

    formatting = {}

    # Font
    rFonts = rPr.find('w:rFonts', NAMESPACES)
    if rFonts is not None:
        if 'w:ascii' in rFonts.attrib:
            formatting['font_ascii'] = rFonts.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii']
        if 'w:hAnsi' in rFonts.attrib:
            formatting['font_hansi'] = rFonts.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi']

    # Font size
    sz = rPr.find('w:sz', NAMESPACES)
    if sz is not None:
        formatting['font_size_half_pts'] = sz.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
        formatting['font_size_points'] = half_points_to_points(formatting['font_size_half_pts'])

    # Font color
    color = rPr.find('w:color', NAMESPACES)
    if color is not None:
        formatting['font_color'] = color.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']

    # Bold
    b = rPr.find('w:b', NAMESPACES)
    if b is not None:
        formatting['bold'] = True

    # Italic
    i = rPr.find('w:i', NAMESPACES)
    if i is not None:
        formatting['italic'] = True

    return formatting

def get_paragraph_text(para):
    """Get text content of a paragraph"""
    texts = []
    for run in para.findall('.//w:t', NAMESPACES):
        if run.text:
            texts.append(run.text)
    return ''.join(texts)

def analyze_numbering_xml(zip_file):
    """Analyze numbering.xml for bullet/numbering definitions"""
    try:
        numbering_xml = zip_file.read('word/numbering.xml')
        root = ET.fromstring(numbering_xml)

        numbering_info = {}

        # Abstract numbering definitions
        for abstractNum in root.findall('w:abstractNum', NAMESPACES):
            abstract_id = abstractNum.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId']
            numbering_info[f'abstract_{abstract_id}'] = {}

            for lvl in abstractNum.findall('w:lvl', NAMESPACES):
                level_id = lvl.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl']
                level_info = {}

                # Number format
                numFmt = lvl.find('w:numFmt', NAMESPACES)
                if numFmt is not None:
                    level_info['format'] = numFmt.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']

                # Level text
                lvlText = lvl.find('w:lvlText', NAMESPACES)
                if lvlText is not None:
                    level_info['text'] = lvlText.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']

                # Level justification
                lvlJc = lvl.find('w:lvlJc', NAMESPACES)
                if lvlJc is not None:
                    level_info['alignment'] = lvlJc.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']

                # Paragraph properties for this level
                pPr = lvl.find('w:pPr', NAMESPACES)
                if pPr is not None:
                    ind = pPr.find('w:ind', NAMESPACES)
                    if ind is not None:
                        if 'w:left' in ind.attrib:
                            level_info['left_indent_twips'] = ind.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left']
                            level_info['left_indent_inches'] = twips_to_inches(level_info['left_indent_twips'])
                        if 'w:hanging' in ind.attrib:
                            level_info['hanging_indent_twips'] = ind.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hanging']
                            level_info['hanging_indent_inches'] = twips_to_inches(level_info['hanging_indent_twips'])

                # Run properties (font for bullet)
                rPr = lvl.find('w:rPr', NAMESPACES)
                if rPr is not None:
                    rFonts = rPr.find('w:rFonts', NAMESPACES)
                    if rFonts is not None:
                        if 'w:ascii' in rFonts.attrib:
                            level_info['bullet_font'] = rFonts.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii']

                numbering_info[f'abstract_{abstract_id}'][f'level_{level_id}'] = level_info

        # Numbering instances
        for num in root.findall('w:num', NAMESPACES):
            num_id = num.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId']
            abstractNumId = num.find('w:abstractNumId', NAMESPACES)
            if abstractNumId is not None:
                abstract_ref = abstractNumId.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val']
                numbering_info[f'num_{num_id}'] = {'abstract_ref': abstract_ref}

        return numbering_info
    except KeyError:
        return {}

def analyze_document(file_path):
    """Analyze a Word document and extract all formatting"""
    with zipfile.ZipFile(file_path, 'r') as zip_file:
        # Read document.xml
        doc_xml = zip_file.read('word/document.xml')
        root = ET.fromstring(doc_xml)

        # Analyze numbering
        numbering_info = analyze_numbering_xml(zip_file)

        # Analyze paragraphs
        paragraphs = []
        for para in root.findall('.//w:p', NAMESPACES):
            para_text = get_paragraph_text(para)
            para_formatting = analyze_paragraph_formatting(para)

            # Analyze runs within paragraph
            runs = []
            for run in para.findall('w:r', NAMESPACES):
                run_text = ''
                t = run.find('w:t', NAMESPACES)
                if t is not None and t.text:
                    run_text = t.text
                run_formatting = analyze_run_formatting(run)
                if run_text or run_formatting:  # Only include if there's text or formatting
                    runs.append({
                        'text': run_text,
                        'formatting': run_formatting
                    })

            if para_text or para_formatting:  # Only include non-empty paragraphs
                paragraphs.append({
                    'text': para_text,
                    'paragraph_formatting': para_formatting,
                    'runs': runs
                })

        return {
            'paragraphs': paragraphs,
            'numbering': numbering_info
        }

def print_formatting_details(doc_info, doc_name):
    """Print detailed formatting information"""
    print(f"\n{'='*80}")
    print(f"FORMATTING ANALYSIS: {doc_name}")
    print(f"{'='*80}\n")

    # Print numbering definitions
    if doc_info['numbering']:
        print("BULLET/NUMBERING DEFINITIONS:")
        print("-" * 80)
        for key, value in sorted(doc_info['numbering'].items()):
            if key.startswith('abstract_'):
                print(f"\n{key}:")
                for lvl_key, lvl_value in sorted(value.items()):
                    print(f"  {lvl_key}:")
                    for prop, val in sorted(lvl_value.items()):
                        print(f"    {prop}: {val}")
            elif key.startswith('num_'):
                print(f"{key}: {value}")
        print()

    # Print paragraph-by-paragraph analysis
    print("PARAGRAPH-BY-PARAGRAPH ANALYSIS:")
    print("-" * 80)
    for i, para in enumerate(doc_info['paragraphs'], 1):
        text_preview = para['text'][:60] + "..." if len(para['text']) > 60 else para['text']
        print(f"\nParagraph {i}: '{text_preview}'")

        if para['paragraph_formatting']:
            print("  Paragraph Formatting:")
            for key, value in sorted(para['paragraph_formatting'].items()):
                print(f"    {key}: {value}")

        if para['runs']:
            for j, run in enumerate(para['runs'], 1):
                if run['formatting']:
                    run_text = run['text'][:40] + "..." if len(run['text']) > 40 else run['text']
                    print(f"  Run {j} ('{run_text}'):")
                    for key, value in sorted(run['formatting'].items()):
                        print(f"    {key}: {value}")

def compare_documents(doc1_info, doc2_info, doc1_name, doc2_name):
    """Compare two documents and highlight differences"""
    print(f"\n{'='*80}")
    print(f"COMPARISON: {doc1_name} vs {doc2_name}")
    print(f"{'='*80}\n")

    print(f"Total paragraphs: {len(doc1_info['paragraphs'])} vs {len(doc2_info['paragraphs'])}")

    # Compare numbering
    if doc1_info['numbering'] != doc2_info['numbering']:
        print("\n*** NUMBERING DEFINITIONS DIFFER ***")
        print("\nDoc1 numbering keys:", set(doc1_info['numbering'].keys()))
        print("Doc2 numbering keys:", set(doc2_info['numbering'].keys()))

    # Compare each paragraph
    print("\nPARAGRAPH DIFFERENCES:")
    print("-" * 80)
    max_paras = max(len(doc1_info['paragraphs']), len(doc2_info['paragraphs']))

    for i in range(max_paras):
        para1 = doc1_info['paragraphs'][i] if i < len(doc1_info['paragraphs']) else None
        para2 = doc2_info['paragraphs'][i] if i < len(doc2_info['paragraphs']) else None

        if para1 is None:
            print(f"\nParagraph {i+1}: EXISTS ONLY IN {doc2_name}")
            continue
        if para2 is None:
            print(f"\nParagraph {i+1}: EXISTS ONLY IN {doc1_name}")
            continue

        # Compare text
        if para1['text'] != para2['text']:
            print(f"\nParagraph {i+1}: TEXT DIFFERS")
            print(f"  {doc1_name}: '{para1['text'][:60]}'")
            print(f"  {doc2_name}: '{para2['text'][:60]}'")

        # Compare paragraph formatting
        fmt1 = para1['paragraph_formatting']
        fmt2 = para2['paragraph_formatting']

        all_keys = set(fmt1.keys()) | set(fmt2.keys())
        differences = []

        for key in sorted(all_keys):
            val1 = fmt1.get(key)
            val2 = fmt2.get(key)
            if val1 != val2:
                differences.append((key, val1, val2))

        if differences:
            text_preview = para1['text'][:50] + "..." if len(para1['text']) > 50 else para1['text']
            print(f"\nParagraph {i+1}: '{text_preview}'")
            print("  FORMATTING DIFFERS:")
            for key, val1, val2 in differences:
                print(f"    {key}:")
                print(f"      {doc1_name} (PERFECT): {val1}")
                print(f"      {doc2_name} (MACRO):   {val2}")

if __name__ == '__main__':
    # File paths
    perfect_doc = '/home/user/VBA/High School Hall of Fame Interactive Wall_2025-11-13.docx'
    macro_doc = '/home/user/VBA/High School Hall of Fame Interactive Wall.docx'

    print("Analyzing documents...")

    # Analyze both documents
    perfect_info = analyze_document(perfect_doc)
    macro_info = analyze_document(macro_doc)

    # Print detailed analysis of perfect document
    print_formatting_details(perfect_info, "PERFECT DOCUMENT")

    # Print detailed analysis of macro document
    print_formatting_details(macro_info, "MACRO DOCUMENT")

    # Compare the two
    compare_documents(perfect_info, macro_info, "PERFECT", "MACRO")

    print(f"\n{'='*80}")
    print("ANALYSIS COMPLETE")
    print(f"{'='*80}\n")
