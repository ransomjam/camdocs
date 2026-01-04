"""Test certification page extraction and generation"""

import sys
sys.path.insert(0, 'c:/Users/user/Desktop/PATTERN/pattern-formatter/backend')

from pattern_formatter_backend import CertificationPageHandler, DocumentProcessor, WordGenerator

def test_certification_handler():
    """Test the CertificationPageHandler class"""
    print("=" * 60)
    print("Testing CertificationPageHandler")
    print("=" * 60)
    
    handler = CertificationPageHandler()
    
    # Check all patterns are defined
    expected_patterns = ['topic', 'author', 'degree', 'supervisor', 'hod', 'director', 'header', 'institution']
    print(f"\nPatterns defined: {list(handler.PATTERNS.keys())}")
    
    for pattern in expected_patterns:
        assert pattern in handler.PATTERNS, f"Missing pattern: {pattern}"
    print("✓ All expected patterns are defined")
    
    # Check initial state
    print(f"\nInitial extracted_data: {handler.extracted_data}")
    print("✓ CertificationPageHandler initialization successful")
    
    return True

def test_document_processor():
    """Test that DocumentProcessor has certification handler"""
    print("\n" + "=" * 60)
    print("Testing DocumentProcessor Integration")
    print("=" * 60)
    
    processor = DocumentProcessor()
    
    # Check certification handler is attached
    assert hasattr(processor, 'certification_handler'), "Missing certification_handler"
    assert hasattr(processor, 'certification_data'), "Missing certification_data attribute"
    assert hasattr(processor, 'certification_start_index'), "Missing certification_start_index"
    
    print("✓ DocumentProcessor has all certification attributes")
    print(f"  - certification_handler: {type(processor.certification_handler).__name__}")
    print(f"  - certification_data: {processor.certification_data}")
    print(f"  - certification_start_index: {processor.certification_start_index}")
    
    return True

def test_word_generator():
    """Test that WordGenerator has certification method"""
    print("\n" + "=" * 60)
    print("Testing WordGenerator Integration")
    print("=" * 60)
    
    generator = WordGenerator()
    
    # Check method exists
    assert hasattr(generator, '_create_certification_page'), "Missing _create_certification_page method"
    print("✓ WordGenerator has _create_certification_page method")
    
    # Check generate signature accepts certification_data
    import inspect
    sig = inspect.signature(generator.generate)
    params = list(sig.parameters.keys())
    assert 'certification_data' in params, "generate() missing certification_data parameter"
    print("✓ generate() method accepts certification_data parameter")
    
    return True

if __name__ == '__main__':
    try:
        test_certification_handler()
        test_document_processor()
        test_word_generator()
        
        print("\n" + "=" * 60)
        print("ALL TESTS PASSED!")
        print("=" * 60)
    except Exception as e:
        print(f"\n❌ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
