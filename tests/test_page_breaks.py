"""Test page breaks for front matter sections"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from pattern_formatter_backend import DocumentProcessor, WordGenerator

def test_page_breaks():
    p = DocumentProcessor()
    
    # Test text with front matter sections
    lines = [
        '# CHAPTER ONE',
        'Introduction text paragraph one.',
        'More text in chapter one.',
        '',
        '# LIST OF TABLES',
        'Table 1: Results.....5',
        'Table 2: Analysis.....10',
        '',
        '# LIST OF FIGURES',
        'Figure 1: Diagram.....15',
        'Figure 2: Chart.....20',
        '',
        '# RESUME',
        'This is the abstract in French. It summarizes the document.',
        '',
        '# LIST OF ABBREVIATIONS',
        'API: Application Programming Interface',
        'HTTP: HyperText Transfer Protocol',
    ]
    
    result, images = p.process_text('\n'.join(lines))
    
    print("Sections found:")
    for s in result['structured']:
        heading = s.get('heading', 'N/A')
        sec_type = s.get('type', 'N/A')
        needs_break = s.get('needs_page_break', False)
        print(f"  {heading}: type={sec_type}, needs_page_break={needs_break}")
    
    # Generate Word document
    output_path = r"C:\Users\user\Desktop\PATTERN\Samples\test_page_breaks_output.docx"
    generator = WordGenerator()
    generator.generate(result['structured'], output_path, images=images)
    print(f"\nOutput saved to: {output_path}")
    print("Check the document to verify page breaks are working.")

if __name__ == '__main__':
    test_page_breaks()
