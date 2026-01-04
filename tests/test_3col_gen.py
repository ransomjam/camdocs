"""Test 3-column cover page layout generation"""
import sys
sys.path.insert(0, 'c:/Users/user/Desktop/PATTERN/pattern-formatter/backend')

from pattern_formatter_backend import DocumentProcessor, WordGenerator

# Process document
processor = DocumentProcessor()
result, images = processor.process_docx('c:/Users/user/Desktop/PATTERN/Samples/sample_dissertation.docx')

print("Cover page data:")
for key, value in processor.cover_page_data.items():
    print(f"  {key}: {value}")

# Generate output
generator = WordGenerator()
output_path = 'c:/Users/user/Desktop/PATTERN/pattern-formatter/tests/outputs/test_3col.docx'
generator.generate(result['structured'], output_path, images=images, cover_page_data=processor.cover_page_data)

print(f"\nGenerated: {output_path}")
print("Done!")
