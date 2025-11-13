#!/usr/bin/env python3
"""
Detailed formatting comparison focused on bullet and paragraph formatting.
"""
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

# Word XML namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
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

def analyze_numbering_definitions(zip_file):
    """Get detailed numbering/bullet formatting from numbering.xml"""
    try:
        numbering_xml = zip_file.read('word/numbering.xml')
        root = ET.fromstring(numbering_xml)

        abstracts = {}
        for abstractNum in root.findall('w:abstractNum', NS):
            abstract_id = abstractNum.get(f'{{{NS["w"]}}}abstractNumId')
            abstracts[abstract_id] = {}

            for lvl in abstractNum.findall('w:lvl', NS):
                level_id = lvl.get(f'{{{NS["w"]}}}ilvl')
                level_data = {'level': level_id}

                # Indentation
                pPr = lvl.find('w:pPr', NS)
                if pPr is not None:
                    ind = pPr.find('w:ind', NS)
                    if ind is not None:
                        left = ind.get(f'{{{NS["w"]}}}left')
                        hanging = ind.get(f'{{{NS["w"]}}}hanging')
                        if left:
                            level_data['left_twips'] = left
                            level_data['left_inches'] = round(twips_to_inches(left), 4)
                        if hanging:
                            level_data['hanging_twips'] = hanging
                            level_data['hanging_inches'] = round(twips_to_inches(hanging), 4)

                # Bullet character
                lvlText = lvl.find('w:lvlText', NS)
                if lvlText is not None:
                    level_data['bullet_char'] = lvlText.get(f'{{{NS["w"]}}}val')

                # Number format
                numFmt = lvl.find('w:numFmt', NS)
                if numFmt is not None:
                    level_data['format'] = numFmt.get(f'{{{NS["w"]}}}val')

                abstracts[abstract_id][level_id] = level_data

        # Get numbering instances
        nums = {}
        for num in root.findall('w:num', NS):
            num_id = num.get(f'{{{NS["w"]}}}numId')
            abstractNumId = num.find('w:abstractNumId', NS)
            if abstractNumId is not None:
                abstract_ref = abstractNumId.get(f'{{{NS["w"]}}}val')
                nums[num_id] = abstract_ref

        return abstracts, nums
    except:
        return {}, {}

def get_bullet_paragraphs(zip_file):
    """Get all bullet paragraphs with their formatting"""
    doc_xml = zip_file.read('word/document.xml')
    root = ET.fromstring(doc_xml)

    bullet_paras = []

    for para_idx, para in enumerate(root.findall('.//w:p', NS)):
        pPr = para.find('w:pPr', NS)
        if pPr is None:
            continue

        # Check if it has numbering
        numPr = pPr.find('w:numPr', NS)
        if numPr is None:
            continue

        # Get text
        text = ''
        for t in para.findall('.//w:t', NS):
            if t.text:
                text += t.text

        # Get numbering info
        ilvl = numPr.find('w:ilvl', NS)
        numId = numPr.find('w:numId', NS)

        level = ilvl.get(f'{{{NS["w"]}}}val') if ilvl is not None else None
        num_id = numId.get(f'{{{NS["w"]}}}val') if numId is not None else None

        # Get spacing
        spacing = pPr.find('w:spacing', NS)
        spacing_info = {}
        if spacing is not None:
            for attr in ['line', 'lineRule', 'before', 'after']:
                val = spacing.get(f'{{{NS["w"]}}}{attr}')
                if val:
                    spacing_info[attr] = val
                    if attr in ['line', 'before', 'after']:
                        spacing_info[f'{attr}_pts'] = round(twips_to_points(val), 2)

        # Get direct indentation (overrides)
        ind = pPr.find('w:ind', NS)
        indent_info = {}
        if ind is not None:
            for attr in ['left', 'hanging', 'firstLine']:
                val = ind.get(f'{{{NS["w"]}}}{attr}')
                if val:
                    indent_info[f'{attr}_twips'] = val
                    indent_info[f'{attr}_inches'] = round(twips_to_inches(val), 4)

        bullet_paras.append({
            'index': para_idx,
            'text': text[:60] + '...' if len(text) > 60 else text,
            'numId': num_id,
            'level': level,
            'spacing': spacing_info,
            'indent': indent_info
        })

    return bullet_paras

def print_bullet_comparison(perfect_path, macro_path):
    """Print a detailed comparison of bullet formatting"""

    print("="*100)
    print("BULLET FORMATTING COMPARISON")
    print("="*100)

    with zipfile.ZipFile(perfect_path, 'r') as perfect_zip:
        perfect_abstracts, perfect_nums = analyze_numbering_definitions(perfect_zip)
        perfect_bullets = get_bullet_paragraphs(perfect_zip)

    with zipfile.ZipFile(macro_path, 'r') as macro_zip:
        macro_abstracts, macro_nums = analyze_numbering_definitions(macro_zip)
        macro_bullets = get_bullet_paragraphs(macro_zip)

    print("\n" + "="*100)
    print("PERFECT DOCUMENT - NUMBERING DEFINITIONS")
    print("="*100)
    for abstract_id, levels in sorted(perfect_abstracts.items()):
        print(f"\nAbstract Numbering {abstract_id}:")
        for level_id, level_data in sorted(levels.items()):
            print(f"  Level {level_id}:")
            for key, val in sorted(level_data.items()):
                print(f"    {key}: {val}")

    print("\nNumbering Instances:")
    for num_id, abstract_ref in sorted(perfect_nums.items()):
        print(f"  numId {num_id} -> abstractNum {abstract_ref}")

    print("\n" + "="*100)
    print("MACRO DOCUMENT - NUMBERING DEFINITIONS")
    print("="*100)
    for abstract_id, levels in sorted(macro_abstracts.items()):
        print(f"\nAbstract Numbering {abstract_id}:")
        for level_id, level_data in sorted(levels.items()):
            print(f"  Level {level_id}:")
            for key, val in sorted(level_data.items()):
                print(f"    {key}: {val}")

    print("\nNumbering Instances:")
    for num_id, abstract_ref in sorted(macro_nums.items()):
        print(f"  numId {num_id} -> abstractNum {abstract_ref}")

    print("\n" + "="*100)
    print("PERFECT DOCUMENT - BULLET PARAGRAPHS")
    print("="*100)
    for bp in perfect_bullets[:10]:  # First 10
        print(f"\nPara {bp['index']}: '{bp['text']}'")
        print(f"  numId: {bp['numId']}, level: {bp['level']}")
        if bp['spacing']:
            print(f"  Spacing: {bp['spacing']}")
        if bp['indent']:
            print(f"  Indent: {bp['indent']}")

    print("\n" + "="*100)
    print("MACRO DOCUMENT - BULLET PARAGRAPHS")
    print("="*100)
    for bp in macro_bullets[:10]:  # First 10
        print(f"\nPara {bp['index']}: '{bp['text']}'")
        print(f"  numId: {bp['numId']}, level: {bp['level']}")
        if bp['spacing']:
            print(f"  Spacing: {bp['spacing']}")
        if bp['indent']:
            print(f"  Indent: {bp['indent']}")

    print("\n" + "="*100)
    print("KEY DIFFERENCES")
    print("="*100)

    # Compare abstract definitions
    print("\n1. NUMBERING DEFINITIONS:")
    all_abstracts = set(perfect_abstracts.keys()) | set(macro_abstracts.keys())
    for abstract_id in sorted(all_abstracts):
        if abstract_id in perfect_abstracts and abstract_id in macro_abstracts:
            perf = perfect_abstracts[abstract_id]
            mac = macro_abstracts[abstract_id]
            if perf != mac:
                print(f"\n  Abstract {abstract_id} DIFFERS:")
                for level_id in set(perf.keys()) | set(mac.keys()):
                    if level_id in perf and level_id in mac:
                        if perf[level_id] != mac[level_id]:
                            print(f"    Level {level_id}:")
                            print(f"      PERFECT: {perf[level_id]}")
                            print(f"      MACRO:   {mac[level_id]}")
        elif abstract_id in perfect_abstracts:
            print(f"\n  Abstract {abstract_id} only in PERFECT")
        else:
            print(f"\n  Abstract {abstract_id} only in MACRO")

    # Compare bullet paragraphs
    print("\n2. BULLET PARAGRAPH FORMATTING:")
    print(f"   Perfect has {len(perfect_bullets)} bullet paragraphs")
    print(f"   Macro has {len(macro_bullets)} bullet paragraphs")

    if perfect_bullets and macro_bullets:
        print("\n   First bullet paragraph comparison:")
        print(f"   PERFECT: {perfect_bullets[0]}")
        print(f"   MACRO:   {macro_bullets[0]}")

if __name__ == '__main__':
    perfect_doc = '/home/user/VBA/High School Hall of Fame Interactive Wall_2025-11-13.docx'
    macro_doc = '/home/user/VBA/High School Hall of Fame Interactive Wall.docx'

    print_bullet_comparison(perfect_doc, macro_doc)
