"""
Test Cover Page Extraction and Replacement
December 30, 2025

Tests the CoverPageExtractor class that:
1. Detects cover page content in academic documents
2. Extracts key fields (name, department, supervisors, etc.)
3. Fills a template with extracted data
4. Replaces the original cover page with the filled template
"""

import os
import sys

# Add backend to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from pattern_formatter_backend import CoverPageExtractor, DocumentProcessor, WordGenerator


class ColorPrint:
    """Colored console output"""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    BOLD = '\033[1m'
    END = '\033[0m'
    
    @staticmethod
    def success(msg): print(f"{ColorPrint.GREEN}[OK] {msg}{ColorPrint.END}")
    
    @staticmethod
    def error(msg): print(f"{ColorPrint.RED}[FAIL] {msg}{ColorPrint.END}")
    
    @staticmethod
    def info(msg): print(f"{ColorPrint.BLUE}[INFO] {msg}{ColorPrint.END}")
    
    @staticmethod
    def header(msg): print(f"\n{ColorPrint.BOLD}{msg}{ColorPrint.END}\n{'='*60}")


def test_cover_page_extractor():
    """Test the CoverPageExtractor class with sample data"""
    ColorPrint.header("TEST: CoverPageExtractor Class")
    
    # Sample cover page paragraphs (simulating what we'd get from a DOCX)
    test_paragraphs = [
        {'text': 'THE UNIVERSITY OF BAMENDA', 'bold': True, 'italic': False, 'index': 0},
        {'text': '', 'bold': False, 'italic': False, 'index': 1},
        {'text': '', 'bold': False, 'italic': False, 'index': 2},
        {'text': '', 'bold': False, 'italic': False, 'index': 3},
        {'text': 'DEPARTMENT OF MANAGEMENT AND ENTREPRENEURSHIP', 'bold': True, 'italic': False, 'index': 4},
        {'text': 'HIGHER INSTITUTE OF COMMERCE AND MANAGEMENT', 'bold': True, 'italic': False, 'index': 5},
        {'text': '', 'bold': False, 'italic': False, 'index': 6},
        {'text': 'THE EFFECT OF COMMUNICATION SKILL ON THE GROWTH OF INTER URBAN TRAVELLING AGENCIES IN BAMENDA', 'bold': True, 'italic': False, 'index': 7},
        {'text': '', 'bold': False, 'italic': False, 'index': 8},
        {'text': 'A Dissertation submitted to the Department of Management and Entrepreneurship, Higher Institute of Commerce and Management in Partial Fulfillment of the Requirements for the Award of a Masters in Business Administration Degree (MBA) in Management and Entrepreneurship', 'bold': False, 'italic': True, 'index': 9},
        {'text': '', 'bold': False, 'italic': False, 'index': 10},
        {'text': 'BY', 'bold': False, 'italic': False, 'index': 11},
        {'text': 'LATAR FLAVIAN DZEKASHU', 'bold': True, 'italic': False, 'index': 12},
        {'text': 'Registration Number: UBA23CP043', 'bold': False, 'italic': False, 'index': 13},
        {'text': '', 'bold': False, 'italic': False, 'index': 14},
        {'text': 'SUPERVISOR', 'bold': False, 'italic': False, 'index': 15},
        {'text': 'Dr. FORBE HODU NGANGNCHI', 'bold': True, 'italic': False, 'index': 16},
        {'text': '', 'bold': False, 'italic': False, 'index': 17},
        {'text': 'JULY 2025', 'bold': True, 'italic': False, 'index': 18},
    ]
    
    passed = 0
    failed = 0
    
    # Test 1: Cover page detection
    extractor = CoverPageExtractor()
    has_cover, data, end_idx = extractor.detect_and_extract(test_paragraphs)
    
    if has_cover:
        ColorPrint.success(f"Cover page detected (ends at index {end_idx})")
        passed += 1
    else:
        ColorPrint.error("Cover page NOT detected")
        failed += 1
    
    # Test 2: Department extraction
    if data['department'] and 'MANAGEMENT' in data['department'].upper():
        ColorPrint.success(f"Department extracted: {data['department']}")
        passed += 1
    else:
        ColorPrint.error(f"Department extraction failed: {data['department']}")
        failed += 1
    
    # Test 3: Faculty extraction
    if data['faculty_or_school'] and 'COMMERCE' in data['faculty_or_school'].upper():
        ColorPrint.success(f"Faculty extracted: {data['faculty_or_school']}")
        passed += 1
    else:
        ColorPrint.error(f"Faculty extraction failed: {data['faculty_or_school']}")
        failed += 1
    
    # Test 4: Name extraction
    if data['name'] and 'LATAR' in data['name'].upper():
        ColorPrint.success(f"Name extracted: {data['name']}")
        passed += 1
    else:
        ColorPrint.error(f"Name extraction failed: {data['name']}")
        failed += 1
    
    # Test 5: Registration number extraction
    if data['registration_number'] and 'UBA23CP043' in data['registration_number']:
        ColorPrint.success(f"Registration number extracted: {data['registration_number']}")
        passed += 1
    else:
        ColorPrint.error(f"Registration number extraction failed: {data['registration_number']}")
        failed += 1
    
    # Test 6: Supervisor extraction
    if data['supervisor'] and 'FORBE' in data['supervisor'].upper():
        ColorPrint.success(f"Supervisor extracted: {data['supervisor']}")
        passed += 1
    else:
        ColorPrint.error(f"Supervisor extraction failed: {data['supervisor']}")
        failed += 1
    
    # Test 7: Date extraction
    if data['month_year'] and 'JULY' in data['month_year'].upper() or '2025' in str(data['month_year']):
        ColorPrint.success(f"Date extracted: {data['month_year']}")
        passed += 1
    else:
        ColorPrint.error(f"Date extraction failed: {data['month_year']}")
        failed += 1
    
    # Test 8: Degree extraction
    if data['degree'] and 'MBA' in str(data['degree']).upper():
        ColorPrint.success(f"Degree extracted: {data['degree']}")
        passed += 1
    else:
        ColorPrint.error(f"Degree extraction failed: {data['degree']}")
        failed += 1
    
    # Test 9: Topic extraction
    if data['topic'] and 'COMMUNICATION' in str(data['topic']).upper():
        ColorPrint.success(f"Topic extracted: {data['topic'][:50]}...")
        passed += 1
    else:
        ColorPrint.error(f"Topic extraction failed: {data['topic']}")
        failed += 1
    
    print(f"\n{'='*60}")
    print(f"Passed: {passed}, Failed: {failed}")
    
    return passed, failed


def test_template_filling():
    """Test that the template can be filled correctly"""
    ColorPrint.header("TEST: Template Filling")
    
    passed = 0
    failed = 0
    
    extractor = CoverPageExtractor()
    
    # Manually set extracted data
    extractor.extracted_data = {
        'department': 'MANAGEMENT AND ENTREPRENEURSHIP',
        'faculty_or_school': 'Higher Institute of Commerce and Management',
        'topic': 'THE EFFECT OF COMMUNICATION SKILL',
        'degree': 'MBA',
        'name': 'LATAR FLAVIAN DZEKASHU',
        'registration_number': 'UBA23CP043',
        'supervisor': 'Dr. FORBE HODU NGANGNCHI',
        'co_supervisor': None,
        'month_year': 'JULY 2025'
    }
    
    # Test template filling
    try:
        filled_doc = extractor.fill_template()
        
        if filled_doc:
            ColorPrint.success("Template loaded successfully")
            passed += 1
            
            # Check if placeholders were replaced
            full_text = '\n'.join([p.text for p in filled_doc.paragraphs])
            
            if 'LATAR FLAVIAN DZEKASHU' in full_text:
                ColorPrint.success("Name placeholder replaced")
                passed += 1
            else:
                ColorPrint.error("Name placeholder NOT replaced")
                failed += 1
            
            if 'UBA23CP043' in full_text:
                ColorPrint.success("Registration number placeholder replaced")
                passed += 1
            else:
                ColorPrint.error("Registration number placeholder NOT replaced")
                failed += 1
            
            if 'JULY 2025' in full_text:
                ColorPrint.success("Date placeholder replaced")
                passed += 1
            else:
                ColorPrint.error("Date placeholder NOT replaced")
                failed += 1
        else:
            ColorPrint.error("Template failed to load")
            failed += 1
    except Exception as e:
        ColorPrint.error(f"Template filling failed: {e}")
        failed += 1
    
    print(f"\n{'='*60}")
    print(f"Passed: {passed}, Failed: {failed}")
    
    return passed, failed


def test_full_document_processing():
    """Test full document processing with sample_dissertation.docx"""
    ColorPrint.header("TEST: Full Document Processing")
    
    passed = 0
    failed = 0
    
    # Path to sample dissertation
    sample_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Samples', 'sample_dissertation.docx')
    
    if not os.path.exists(sample_path):
        ColorPrint.error(f"Sample file not found: {sample_path}")
        return 0, 1
    
    ColorPrint.info(f"Processing: {sample_path}")
    
    try:
        # Process document
        processor = DocumentProcessor()
        result, images = processor.process_docx(sample_path)
        
        # Check if cover page was detected
        if processor.filled_template_doc:
            ColorPrint.success("Cover page detected and template filled")
            passed += 1
            
            # Get extracted data
            data = processor.cover_page_extractor.extracted_data
            ColorPrint.info(f"Extracted data:")
            for key, value in data.items():
                if value:
                    print(f"    {key}: {value[:50] if len(str(value)) > 50 else value}")
        else:
            ColorPrint.error("Cover page NOT detected")
            failed += 1
        
        # Generate output document
        output_path = os.path.join(os.path.dirname(__file__), 'outputs', 'test_cover_page_output.docx')
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        generator = WordGenerator()
        generator.generate(
            result['structured'], 
            output_path, 
            images=images,
            cover_page_template_doc=processor.filled_template_doc
        )
        
        if os.path.exists(output_path):
            ColorPrint.success(f"Output document created: {output_path}")
            passed += 1
            
            # Check output file size
            size = os.path.getsize(output_path)
            ColorPrint.info(f"Output file size: {size} bytes")
        else:
            ColorPrint.error("Output document NOT created")
            failed += 1
            
    except Exception as e:
        ColorPrint.error(f"Document processing failed: {e}")
        import traceback
        traceback.print_exc()
        failed += 1
    
    print(f"\n{'='*60}")
    print(f"Passed: {passed}, Failed: {failed}")
    
    return passed, failed


def run_all_tests():
    """Run all cover page extraction tests"""
    ColorPrint.header("COVER PAGE EXTRACTION TEST SUITE")
    
    total_passed = 0
    total_failed = 0
    
    # Test 1: Extractor class
    p, f = test_cover_page_extractor()
    total_passed += p
    total_failed += f
    
    # Test 2: Template filling
    p, f = test_template_filling()
    total_passed += p
    total_failed += f
    
    # Test 3: Full document processing
    p, f = test_full_document_processing()
    total_passed += p
    total_failed += f
    
    # Summary
    ColorPrint.header("FINAL SUMMARY")
    print(f"Total Passed: {ColorPrint.GREEN}{total_passed}{ColorPrint.END}")
    print(f"Total Failed: {ColorPrint.RED}{total_failed}{ColorPrint.END}")
    
    success_rate = (total_passed / (total_passed + total_failed)) * 100 if (total_passed + total_failed) > 0 else 0
    print(f"Success Rate: {ColorPrint.BOLD}{success_rate:.1f}%{ColorPrint.END}")
    
    return total_failed == 0


if __name__ == '__main__':
    success = run_all_tests()
    sys.exit(0 if success else 1)
