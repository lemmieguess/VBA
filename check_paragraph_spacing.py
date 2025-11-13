#!/usr/bin/env python3
"""
Check paragraph spacing and formatting for all paragraphs.
"""
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def twips_to_points(twips):
    if twips is None:
        return None
    return round(float(twips) / 20, 2)

def analyze_all_paragraphs(file_path):
    """Get all paragraph formatting"""
    with zipfile.ZipFile(file_path, 'r') as zip_file:
        doc_xml = zip_file.read('word/document.xml')
        root = ET.fromstring(doc_xml)

        paragraphs = []
        for para in root.findall('.//w:p', NS):
            pPr = para.find('w:pPr', NS)

            text = ''
            for t in para.findall('.//w:t', NS):
                if t.text:
                    text += t.text

            para_info = {'text': text[:50] + '...' if len(text) > 50 else text}

            if pPr is not None:
                # Spacing
                spacing = pPr.find('w:spacing', NS)
                if spacing is not None:
                    line = spacing.get(f'{{{NS["w"]}}}line')
                    line_rule = spacing.get(f'{{{NS["w"]}}}lineRule')
                    before = spacing.get(f'{{{NS["w"]}}}before')
                    after = spacing.get(f'{{{NS["w"]}}}after')

                    if line:
                        para_info['line_spacing'] = f"{line} twips ({twips_to_points(line)} pts)"
                        para_info['line_rule'] = line_rule
                    if before:
                        para_info['space_before'] = f"{before} twips ({twips_to_points(before)} pts)"
                    if after:
                        para_info['space_after'] = f"{after} twips ({twips_to_points(after)} pts)"

                # Indentation
                ind = pPr.find('w:ind', NS)
                if ind is not None:
                    left = ind.get(f'{{{NS["w"]}}}left')
                    if left:
                        para_info['left_indent'] = f"{left} twips ({round(float(left)/1440, 2)} in)"

            paragraphs.append(para_info)

        return paragraphs

def print_formatting_summary(paras, doc_name):
    """Print a summary of formatting"""
    print(f"\n{'='*100}")
    print(f"{doc_name} - FORMATTING SUMMARY")
    print(f"{'='*100}")

    # Count unique spacing combinations
    spacing_combos = Counter()
    for p in paras:
        combo = f"Line: {p.get('line_spacing', 'default')}, Before: {p.get('space_before', '0')}, After: {p.get('space_after', '0')}"
        spacing_combos[combo] += 1

    print("\nSPACING PATTERNS (line spacing, space before, space after):")
    for combo, count in spacing_combos.most_common(10):
        print(f"  {combo} - {count} paragraphs")

    # Show first few paragraphs with their formatting
    print("\nFIRST 15 PARAGRAPHS WITH FORMATTING:")
    for i, p in enumerate(paras[:15], 1):
        if p['text']:
            print(f"\n{i}. '{p['text']}'")
            for key in ['line_spacing', 'line_rule', 'space_before', 'space_after', 'left_indent']:
                if key in p:
                    print(f"   {key}: {p[key]}")

def compare_spacing(perfect_paras, macro_paras):
    """Compare spacing between documents"""
    print(f"\n{'='*100}")
    print("SPACING DIFFERENCES")
    print(f"{'='*100}")

    differences = []
    max_len = max(len(perfect_paras), len(macro_paras))

    for i in range(min(30, max_len)):  # Check first 30
        if i >= len(perfect_paras) or i >= len(macro_paras):
            continue

        p1 = perfect_paras[i]
        p2 = macro_paras[i]

        diff = {}
        for key in ['line_spacing', 'space_before', 'space_after', 'left_indent']:
            val1 = p1.get(key, 'none')
            val2 = p2.get(key, 'none')
            if val1 != val2:
                diff[key] = (val1, val2)

        if diff:
            differences.append((i, p1['text'], diff))

    if differences:
        print(f"\nFound {len(differences)} paragraphs with different spacing:")
        for idx, text, diff in differences:
            print(f"\nParagraph {idx + 1}: '{text}'")
            for key, (val1, val2) in diff.items():
                print(f"  {key}:")
                print(f"    PERFECT: {val1}")
                print(f"    MACRO:   {val2}")
    else:
        print("\nNo spacing differences found in first 30 paragraphs")

if __name__ == '__main__':
    perfect_doc = '/home/user/VBA/High School Hall of Fame Interactive Wall_2025-11-13.docx'
    macro_doc = '/home/user/VBA/High School Hall of Fame Interactive Wall.docx'

    print("Analyzing paragraph spacing...")

    perfect_paras = analyze_all_paragraphs(perfect_doc)
    macro_paras = analyze_all_paragraphs(macro_doc)

    print_formatting_summary(perfect_paras, "PERFECT DOCUMENT")
    print_formatting_summary(macro_paras, "MACRO DOCUMENT")
    compare_spacing(perfect_paras, macro_paras)
