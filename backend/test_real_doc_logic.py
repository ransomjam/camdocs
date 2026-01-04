import sys
import os

print("Starting test_real_doc_logic.py...")
print(f"__name__ is {__name__}")

# Add the current directory to sys.path to ensure we can import the backend
sys.path.append(os.getcwd())

try:
    print("Importing pattern_formatter_backend...")
    import pattern_formatter_backend
    print(f"pattern_formatter_backend imported. Name: {pattern_formatter_backend.__name__}")
    from pattern_formatter_backend import HierarchyCorrector, HeadingNumberer
except ImportError as e:
    print(f"Import failed: {e}")
    sys.exit(1)

def test_on_sample_file():
    sample_file_path = os.path.join('..', 'tests', 'sample_academic_paper.txt')
    
    try:
        with open(sample_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print(f"Error: Could not find sample file at {sample_file_path}")
        # Create a dummy sample if the file doesn't exist
        content = """
1. Introduction
This is the intro.
1.1 Background
Some background.
1.1.1. Specifics
Details here.
2. Methodology
Methods used.
2.1 Data Collection
How we got data.
        """
        print("Using dummy content instead.")

    print("Original Content (First 500 chars):")
    print(content[:500])
    print("-" * 40)

    # Apply Hierarchy Correction
    corrector = HierarchyCorrector()
    corrected_content = corrector.correct(content)

    # Apply Heading Numbering
    numberer = HeadingNumberer()
    numberer.current_chapter = 1 # Force chapter 1 context
    
    print("Processed Content (Headings only):")
    lines = corrected_content.split('\n')
    
    # Use the built-in batch processing method
    results = numberer.process_document_headings(lines)
    
    final_lines = []
    for res in results:
        final_lines.append(res['numbered'])
        if res.get('is_heading') and res.get('number'):
            print(res['numbered'])
        elif res.get('is_heading') and res.get('level') == 1:
             # Print unnumbered headings like ABSTRACT, INTRODUCTION
             print(res['numbered'])

    final_content = '\n'.join(final_lines)
            
    print("-" * 40)
    print("Full Processed Content (First 500 chars):")
    print(final_content[:500])

if __name__ == "__main__":
    test_on_sample_file()
