import sys
import os
import json

# Add backend to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../backend')))

from pattern_formatter_backend import DocumentProcessor

def test_structure():
    processor = DocumentProcessor()
    
    with open('pattern-formatter/tests/test_resume_break.txt', 'r', encoding='utf-8') as f:
        text = f.read()
        
    print("Analyzing text...")
    # process_text returns (result, images)
    result, images = processor.process_text(text)
    
    structure = result['structured']
    
    print(f"Found {len(structure)} sections.")
    
    for i, section in enumerate(structure):
        heading = section.get('heading', 'NO_HEADING')
        sec_type = section.get('type', 'unknown')
        fm_type = section.get('front_matter_type', 'N/A')
        needs_break = section.get('needs_page_break', False)
        use_break_before = section.get('use_page_break_before', False)
        
        print(f"Section {i+1}: '{heading}' (Type: {sec_type}, FM: {fm_type})")
        print(f"  - needs_page_break: {needs_break}")
        print(f"  - use_page_break_before: {use_break_before}")
        
        # Check if our logic in _add_section would trigger
        heading_upper = heading.strip().upper()
        force_break_headings = ['RESUME', 'RÉSUMÉ', 'ACKNOWLEDGEMENTS', 'ACKNOWLEDGMENTS', 'ACKNOWLEDGEMENT']
        should_force = any(h in heading_upper for h in force_break_headings)
        print(f"  - Should force break? {should_force}")

if __name__ == "__main__":
    test_structure()