"""
Test List of Tables (LOT) generation with the sample document.
"""
import sys
import os
import time

# Add backend to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from pattern_formatter_backend import DocumentProcessor, WordGenerator, TableFormatter, update_toc_with_word

def test_table_detection():
    """Test TableFormatter detection capabilities."""
    print("=" * 60)
    print("TEST 1: TableFormatter Detection")
    print("=" * 60)
    
    tf = TableFormatter()
    
    test_cases = [
        'Table 1: Summary of Variables',
        'Table 2.1: Demographic Characteristics',
        'Tbl. 3: Cost Sharing Mechanisms',
        'Tab 4.5: Regression Results',
        'TABLE 5: ANOVA Summary',
    ]
    
    detected = 0
    for test in test_cases:
        result = tf.detect_table_caption(test)
        if result:
            detected += 1
            print(f"  [OK] {test}")
            print(f"       Number: {result['number']}, Title: {result['title']}")
        else:
            print(f"  [FAIL] {test}")
    
    print(f"\nDetection rate: {detected}/{len(test_cases)}")
    return detected == len(test_cases)

def test_sample_document():
    """Test processing the sample document with tables."""
    print("\n" + "=" * 60)
    print("TEST 2: Process Sample Document with Tables")
    print("=" * 60)
    
    # Path to sample document
    sample_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Samples', 'sample project with tables.docx')
    
    if not os.path.exists(sample_path):
        print(f"  [SKIP] Sample file not found: {sample_path}")
        # Try alternate path
        sample_path = r"C:\Users\user\Desktop\PATTERN\Samples\sample project with tables.docx"
        if not os.path.exists(sample_path):
            print(f"  [SKIP] Sample file not found at alternate path either")
            return False
    
    print(f"  Processing: {sample_path}")
    
    # Process document
    processor = DocumentProcessor()
    try:
        result = processor.process_file(sample_path)
        print(f"  [OK] Document processed successfully")
        print(f"       Sections: {len(result.get('structured_data', []))}")
        
        # Count table captions detected
        tf = TableFormatter()
        table_count = 0
        for section in result.get('structured_data', []):
            for item in section.get('content', []):
                if item.get('type') == 'paragraph':
                    text = item.get('text', '')
                    if tf.is_table_caption(text):
                        table_count += 1
                        print(f"       Found table: {text[:50]}...")
                elif item.get('type') == 'table':
                    caption = item.get('caption', '')
                    if caption and tf.is_table_caption(caption):
                        table_count += 1
                        print(f"       Found table with caption: {caption[:50]}...")
        
        print(f"       Total tables detected: {table_count}")
        
        # Generate output document
        output_path = os.path.join(os.path.dirname(__file__), '..', 'outputs', f'test_lot_{int(time.time())}.docx')
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        generator = WordGenerator()
        generator.generate(
            result['structured_data'],
            output_path,
            images=result.get('images', []),
            cover_page_data=result.get('cover_page_data'),
            certification_data=result.get('certification_data')
        )
        
        print(f"  [OK] Document generated: {output_path}")
        print(f"       Has tables: {generator.has_tables}")
        print(f"       Table entries tracked: {len(generator.table_entries)}")
        
        # Check if LOT was generated
        if generator.has_tables:
            print("  [OK] List of Tables should be included")
        else:
            print("  [WARN] No tables detected - LOT may not be included")
        
        return True
        
    except Exception as e:
        print(f"  [FAIL] Error processing document: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("=" * 60)
    print("LIST OF TABLES (LOT) TEST SUITE")
    print("=" * 60)
    
    results = []
    
    # Test 1: TableFormatter detection
    results.append(("TableFormatter Detection", test_table_detection()))
    
    # Test 2: Sample document processing
    results.append(("Sample Document Processing", test_sample_document()))
    
    # Summary
    print("\n" + "=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    
    passed = 0
    for name, result in results:
        status = "PASS" if result else "FAIL"
        print(f"  {name}: {status}")
        if result:
            passed += 1
    
    print(f"\nTotal: {passed}/{len(results)} tests passed")
    return passed == len(results)

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
