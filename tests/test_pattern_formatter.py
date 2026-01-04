# test_pattern_formatter.py
# Comprehensive test suite for pattern-based document formatter

import sys
import os
import time

# Add backend to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

try:
    from pattern_formatter_backend import PatternEngine, DocumentProcessor, WordGenerator
except ImportError:
    print("Error: Could not import backend modules.")
    print("Make sure pattern_formatter_backend.py is in the backend folder.")
    sys.exit(1)


class ColorPrint:
    """Colored output for terminal"""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    BOLD = '\033[1m'
    END = '\033[0m'
    
    @staticmethod
    def success(msg):
        print(f"{ColorPrint.GREEN}✓{ColorPrint.END} {msg}")
    
    @staticmethod
    def error(msg):
        print(f"{ColorPrint.RED}✗{ColorPrint.END} {msg}")
    
    @staticmethod
    def info(msg):
        print(f"{ColorPrint.BLUE}ℹ{ColorPrint.END} {msg}")
    
    @staticmethod
    def warning(msg):
        print(f"{ColorPrint.YELLOW}⚠{ColorPrint.END} {msg}")
    
    @staticmethod
    def header(msg):
        print(f"\n{ColorPrint.BOLD}{ColorPrint.BLUE}{msg}{ColorPrint.END}")
        print("=" * 60)


class TestPatternEngine:
    """Test pattern recognition engine"""
    
    def __init__(self):
        self.engine = PatternEngine()
        self.passed = 0
        self.failed = 0
    
    def test_heading_level_1(self):
        """Test H1 heading pattern matching"""
        ColorPrint.header("TEST 1: Heading Level 1 Detection")
        
        test_cases = [
            # (line, should_be_h1)
            ("INTRODUCTION", True),
            ("ABSTRACT", True),
            ("CHAPTER 1: FOUNDATIONS", True),
            ("CHAPTER 5 RESULTS", True),
            ("PART I OVERVIEW", True),
            ("CONCLUSION", True),
            ("REFERENCES", True),
            ("BIBLIOGRAPHY", True),
            ("EXECUTIVE SUMMARY", True),
            ("LITERATURE REVIEW", True),
            ("METHODOLOGY", True),
            ("APPENDIX", True),
            ("This is a normal paragraph.", False),
            ("The quick brown fox jumps over the lazy dog.", False),
        ]
        
        for line, should_be_h1 in test_cases:
            result = self.engine.analyze_line(line, 0)
            is_h1 = result['type'] == 'heading' and result.get('level') == 1
            
            if is_h1 == should_be_h1:
                ColorPrint.success(f"'{line[:40]}...' → {'H1' if is_h1 else 'Not H1'}")
                self.passed += 1
            else:
                ColorPrint.error(f"'{line[:40]}...' → Got {'H1' if is_h1 else result['type']}, expected {'H1' if should_be_h1 else 'Not H1'}")
                self.failed += 1
    
    def test_heading_level_2(self):
        """Test H2 heading pattern matching"""
        ColorPrint.header("TEST 2: Heading Level 2 Detection")
        
        test_cases = [
            # (line, should_be_h2)
            ("Background and Motivation", True),
            ("Methods and Results", True),
            ("Literature Review", True),
            ("1.1 Overview of the Study", True),
            ("2.3 Data Collection Methods", True),
            ("Analysis of Results", True),
            ("This is a paragraph with some text that is too long to be a heading.", False),
        ]
        
        for line, should_be_h2 in test_cases:
            result = self.engine.analyze_line(line, 0)
            is_h2 = result['type'] == 'heading' and result.get('level') == 2
            
            if is_h2 == should_be_h2:
                ColorPrint.success(f"'{line[:40]}...' → {'H2' if is_h2 else 'Not H2'}")
                self.passed += 1
            else:
                ColorPrint.error(f"'{line[:40]}...' → Got {'H2' if is_h2 else result['type']} L{result.get('level', 0)}, expected {'H2' if should_be_h2 else 'Not H2'}")
                self.failed += 1
    
    def test_heading_level_3(self):
        """Test H3 heading pattern matching"""
        ColorPrint.header("TEST 3: Heading Level 3 Detection")
        
        test_cases = [
            # (line, should_be_h3)
            ("1.1.1 Detailed Analysis", True),
            ("2.3.4 Sub-subsection Title", True),
            ("a) First Subsection", True),
            ("(b) Second Subsection", True),
        ]
        
        for line, should_be_h3 in test_cases:
            result = self.engine.analyze_line(line, 0)
            is_h3 = result['type'] == 'heading' and result.get('level') == 3
            
            if is_h3 == should_be_h3:
                ColorPrint.success(f"'{line[:40]}' → {'H3' if is_h3 else 'Not H3'}")
                self.passed += 1
            else:
                ColorPrint.error(f"'{line[:40]}' → Got {result['type']} L{result.get('level', 0)}, expected {'H3' if should_be_h3 else 'Not H3'}")
                self.failed += 1
    
    def test_reference_detection(self):
        """Test reference pattern matching"""
        ColorPrint.header("TEST 4: Reference Detection")
        
        test_cases = [
            # APA format
            "Smith, J. (2024). Title of paper. Journal Name.",
            "Johnson, A. B. (2023). Research findings. Publisher.",
            # Multiple authors
            "Smith, J. & Brown, M. (2024). Collaborative research.",
            # Et al format
            "Johnson et al. 2023. Research findings in academia.",
            # IEEE format
            "[1] Author, I. (2024). Book title. Publisher.",
            "[15] Smith, J. Technical report on systems.",
            # Web reference
            "Brown, A. Retrieved from https://example.com/article",
            "Jones, M. (2024). Online resource. https://journal.org",
        ]
        
        for line in test_cases:
            result = self.engine.analyze_line(line, 0)
            
            if result['type'] == 'reference':
                ColorPrint.success(f"Reference detected: {line[:50]}...")
                self.passed += 1
            else:
                ColorPrint.error(f"Failed to detect reference: {line[:50]}...")
                self.failed += 1
    
    def test_bullet_list_detection(self):
        """Test bullet list pattern matching"""
        ColorPrint.header("TEST 5: Bullet List Detection")
        
        test_cases = [
            ("• First bullet point with standard bullet", "bullet_list"),
            ("● Filled circle bullet point", "bullet_list"),
            ("○ Empty circle bullet point", "bullet_list"),
            ("- Dash bullet point", "bullet_list"),
            ("– En-dash bullet point", "bullet_list"),
            ("— Em-dash bullet point", "bullet_list"),
            ("* Asterisk bullet point", "bullet_list"),
            ("▪ Square bullet point", "bullet_list"),
        ]
        
        for line, expected_type in test_cases:
            result = self.engine.analyze_line(line, 0)
            
            if result['type'] == expected_type:
                ColorPrint.success(f"'{line[:40]}' → {expected_type}")
                self.passed += 1
            else:
                ColorPrint.error(f"'{line[:40]}' → Got {result['type']}, expected {expected_type}")
                self.failed += 1
    
    def test_numbered_list_detection(self):
        """Test numbered list pattern matching"""
        ColorPrint.header("TEST 6: Numbered List Detection")
        
        test_cases = [
            ("1. First numbered item", "numbered_list"),
            ("2) Second numbered item", "numbered_list"),
            ("a. Lettered item lowercase", "numbered_list"),
            ("b) Lettered item with parenthesis", "numbered_list"),
            ("i. Roman numeral item", "numbered_list"),
            ("ii) Roman numeral with paren", "numbered_list"),
            ("A. Capital letter item", "numbered_list"),
            ("(1) Parenthesized number", "numbered_list"),
        ]
        
        for line, expected_type in test_cases:
            result = self.engine.analyze_line(line, 0)
            
            if result['type'] == expected_type:
                ColorPrint.success(f"'{line}' → {expected_type}")
                self.passed += 1
            else:
                ColorPrint.error(f"'{line}' → Got {result['type']}, expected {expected_type}")
                self.failed += 1
    
    def test_definition_detection(self):
        """Test definition pattern matching"""
        ColorPrint.header("TEST 7: Definition Detection")
        
        test_cases = [
            ("Definition: A clear explanation of a term or concept.",),
            ("Objective: The main goal of the research study.",),
            ("Method: The approach used in the study.",),
            ("Conclusion: Final remarks and findings of the work.",),
            ("Purpose: To investigate the relationship between variables.",),
            ("Goal: Achieve a 95% accuracy rate in classification.",),
            ("Result: The experiment showed significant improvement.",),
            ("Note: This is an important consideration.",),
            ("Key Point: Understanding the fundamentals is crucial.",),
        ]
        
        for (line,) in test_cases:
            result = self.engine.analyze_line(line, 0)
            
            if result['type'] == 'definition':
                ColorPrint.success(f"Definition: {result.get('term', 'N/A')}")
                self.passed += 1
            else:
                ColorPrint.error(f"Failed to detect definition in: {line[:50]}")
                self.failed += 1
    
    def test_table_detection(self):
        """Test table pattern matching"""
        ColorPrint.header("TEST 8: Table Detection")
        
        test_cases = [
            ("[TABLE START]", "table_start"),
            ("[TABLE END]", "table_end"),
            ("[table start]", "table_start"),
            ("[table end]", "table_end"),
            ("Table 1: Sample data comparison", "table_caption"),
            ("TABLE 2: Results Summary", "table_caption"),
            ("| Header 1 | Header 2 | Header 3 |", "table_row"),
            ("| Data 1 | Data 2 | Data 3 |", "table_row"),
            ("| --- | --- | --- |", "table_separator"),
        ]
        
        for line, expected_type in test_cases:
            result = self.engine.analyze_line(line, 0)
            
            if result['type'] == expected_type:
                ColorPrint.success(f"'{line}' → {expected_type}")
                self.passed += 1
            else:
                ColorPrint.error(f"'{line}' → Got {result['type']}, expected {expected_type}")
                self.failed += 1
    
    def test_figure_detection(self):
        """Test figure caption detection"""
        ColorPrint.header("TEST 9: Figure Detection")
        
        test_cases = [
            ("Figure 1: System architecture diagram", "figure"),
            ("Figure 10: Results visualization", "figure"),
            ("Fig. 1: Comparison chart", "figure"),
            ("Fig.2: Data flow diagram", "figure"),
            ("Image 1: Screenshot of the interface", "figure"),
            ("Chart 5: Performance metrics", "figure"),
        ]
        
        for line, expected_type in test_cases:
            result = self.engine.analyze_line(line, 0)
            
            if result['type'] == expected_type:
                ColorPrint.success(f"'{line}' → {expected_type}")
                self.passed += 1
            else:
                ColorPrint.error(f"'{line}' → Got {result['type']}, expected {expected_type}")
                self.failed += 1
    
    def run_all_tests(self):
        """Run all pattern tests"""
        ColorPrint.header("🧪 PATTERN ENGINE TEST SUITE")
        
        self.test_heading_level_1()
        self.test_heading_level_2()
        self.test_heading_level_3()
        self.test_reference_detection()
        self.test_bullet_list_detection()
        self.test_numbered_list_detection()
        self.test_definition_detection()
        self.test_table_detection()
        self.test_figure_detection()
        
        # Print summary
        ColorPrint.header("PATTERN ENGINE TEST SUMMARY")
        total = self.passed + self.failed
        percentage = (self.passed / total * 100) if total > 0 else 0
        
        print(f"Total Tests: {total}")
        print(f"Passed: {ColorPrint.GREEN}{self.passed}{ColorPrint.END}")
        print(f"Failed: {ColorPrint.RED}{self.failed}{ColorPrint.END}")
        print(f"Success Rate: {ColorPrint.BOLD}{percentage:.1f}%{ColorPrint.END}")
        
        return self.failed == 0


class TestDocumentProcessor:
    """Test document processing"""
    
    def __init__(self):
        self.processor = DocumentProcessor()
        self.passed = 0
        self.failed = 0
    
    def test_sample_document(self):
        """Test processing a complete sample document"""
        ColorPrint.header("TEST 10: Document Processing")
        
        sample_doc = """
INTRODUCTION

Background and Motivation

Cloud computing has revolutionized enterprise IT infrastructure. The following points highlight key advantages:

• Cost efficiency through shared resources
• Scalability on demand
• Global infrastructure access
- Reduced maintenance overhead

Definition: Cloud computing refers to on-demand delivery of computing resources over the internet.

1.1 Economies of Scale

Major providers achieve cost advantages through:

1. Bulk purchasing power
2. Operational efficiency
3. Automated management

[TABLE START]
| Provider | Cost | Rating |
| AWS | $0.10 | 4.5 |
| Azure | $0.09 | 4.3 |
| GCP | $0.08 | 4.4 |
[TABLE END]

Figure 1: Cloud provider comparison chart

CONCLUSION

Cloud computing provides significant benefits for enterprises.

REFERENCES

Smith, J. (2024). Cloud Economics. Tech Journal.
Johnson, A. et al. 2023. Scaling strategies for cloud.
Brown, M. Retrieved from https://cloudresearch.org/paper
"""
        
        ColorPrint.info("Processing sample document...")
        # process_text returns (result_dict, images_list)
        result_tuple = self.processor.process_text(sample_doc)
        
        if isinstance(result_tuple, tuple):
            result, _ = result_tuple
        else:
            result = result_tuple
        
        stats = result['stats']
        structured = result['structured']
        
        # Validate results
        tests = [
            (stats['headings'] >= 3, f"Headings detected: {stats['headings']} (expected >= 3)"),
            (stats['paragraphs'] >= 2, f"Paragraphs detected: {stats['paragraphs']} (expected >= 2)"),
            (stats['references'] >= 3, f"References detected: {stats['references']} (expected >= 3)"),
            (stats['lists'] >= 4, f"List items detected: {stats['lists']} (expected >= 4)"),
            (stats['definitions'] >= 1, f"Definitions detected: {stats['definitions']} (expected >= 1)"),
            (stats['tables'] >= 1, f"Tables detected: {stats['tables']} (expected >= 1)"),
            (len(structured) >= 3, f"Sections structured: {len(structured)} (expected >= 3)"),
        ]
        
        for test, msg in tests:
            if test:
                ColorPrint.success(msg)
                self.passed += 1
            else:
                ColorPrint.error(msg)
                self.failed += 1
        
        # Show structure
        ColorPrint.info("\nDocument Structure:")
        for section in structured:
            print(f"  H{section['level']}: {section['heading']} ({len(section['content'])} items)")
        
        return self.failed == 0
    
    def test_edge_cases(self):
        """Test edge cases and unusual patterns"""
        ColorPrint.header("TEST 11: Edge Cases")
        
        edge_cases = [
            # Empty lines
            ("", "empty"),
            ("   ", "empty"),
            
            # Very long lines (should be paragraph)
            ("This is a very long paragraph that exceeds the typical length of a heading and should be classified as a regular paragraph rather than a heading even though it might start with a capital letter. " * 3, "paragraph"),
            
            # Numbers only
            ("123456", "paragraph"),
            
            # Special characters only (treated as paragraph)
            ("@#$%^&*()", "paragraph"),
            
            # Single word (too short for most patterns)
            ("Hi", "paragraph"),
        ]
        
        for line, expected_type in edge_cases:
            result = self.processor.engine.analyze_line(line.strip() if line.strip() else line, 0)
            
            if result['type'] == expected_type:
                display_line = line[:30] + '...' if len(line) > 30 else line
                display_line = display_line.replace('\n', ' ')
                ColorPrint.success(f"Edge case: '{display_line}' → {expected_type}")
                self.passed += 1
            else:
                display_line = line[:30] + '...' if len(line) > 30 else line
                display_line = display_line.replace('\n', ' ')
                ColorPrint.error(f"Edge case: '{display_line}' → Got {result['type']}, expected {expected_type}")
                self.failed += 1
        
        return self.failed == 0
    
    def test_no_duplication(self):
        """Test that content is not duplicated"""
        ColorPrint.header("TEST 12: No Content Duplication")
        
        sample_text = """
INTRODUCTION

This is the first paragraph of content.

Background Information

This is the second paragraph of content.

• First bullet item
• Second bullet item

CONCLUSION

Final paragraph of the document.
"""
        
        result_tuple = self.processor.process_text(sample_text)
        
        if isinstance(result_tuple, tuple):
            result, _ = result_tuple
        else:
            result = result_tuple
        
        # Count all content items
        all_content = []
        for section in result['structured']:
            for item in section.get('content', []):
                if item.get('text'):
                    all_content.append(item['text'])
                elif item.get('items'):
                    # Handle both string items and dict items (enhanced bullets)
                    for list_item in item['items']:
                        if isinstance(list_item, dict):
                            # Extract content from enhanced bullet dict or other dict types
                            content = list_item.get('content', list_item.get('text', ''))
                            if content:
                                all_content.append(content)
                        else:
                            all_content.append(str(list_item))
        
        # Check for duplicates
        unique_content = set(all_content)
        has_duplicates = len(all_content) != len(unique_content)
        
        if not has_duplicates:
            ColorPrint.success(f"No duplicates found in {len(all_content)} content items")
            self.passed += 1
        else:
            ColorPrint.error(f"Duplicates found! {len(all_content)} items but only {len(unique_content)} unique")
            self.failed += 1
        
        return not has_duplicates


class TestPerformance:
    """Test system performance"""
    
    def __init__(self):
        self.processor = DocumentProcessor()
    
    def test_speed(self):
        """Test processing speed"""
        ColorPrint.header("TEST 13: Performance Benchmarks")
        
        # Generate test documents of various sizes
        test_sizes = [
            (100, "Small (100 lines)"),
            (500, "Medium (500 lines)"),
            (1000, "Large (1,000 lines)"),
            (5000, "Very Large (5,000 lines)"),
        ]
        
        all_passed = True
        
        for line_count, label in test_sizes:
            # Generate document
            lines = []
            for i in range(line_count):
                if i % 50 == 0:
                    lines.append(f"SECTION {i // 50}")
                elif i % 25 == 0:
                    lines.append(f"Subsection Title {i // 25}")
                elif i % 10 == 0:
                    lines.append(f"• Bullet point number {i}")
                elif i % 5 == 0:
                    lines.append(f"Definition: A sample definition for item {i}.")
                else:
                    lines.append(f"This is paragraph number {i} with some sample text content for testing purposes.")
            
            doc_text = "\n".join(lines)
            
            # Measure processing time
            start_time = time.time()
            result = self.processor.process_text(doc_text)
            elapsed = time.time() - start_time
            
            lines_per_second = line_count / elapsed if elapsed > 0 else float('inf')
            
            # Check performance threshold (should process at least 500 lines/second)
            if lines_per_second >= 500:
                ColorPrint.success(f"{label}: {elapsed:.3f}s ({lines_per_second:.0f} lines/sec)")
            else:
                ColorPrint.warning(f"{label}: {elapsed:.3f}s ({lines_per_second:.0f} lines/sec) - Below threshold")
                all_passed = False
        
        return all_passed
    
    def test_memory_efficiency(self):
        """Test memory efficiency with large documents"""
        ColorPrint.header("TEST 14: Memory Efficiency")
        
        import gc
        
        # Force garbage collection before test
        gc.collect()
        
        # Generate a large document
        lines = [f"Line {i}: Sample content for memory testing." for i in range(10000)]
        doc_text = "\n".join(lines)
        
        # Process and check if it completes without memory issues
        try:
            start_time = time.time()
            result_tuple = self.processor.process_text(doc_text)
            if isinstance(result_tuple, tuple):
                result, _ = result_tuple
            else:
                result = result_tuple
            
            elapsed = time.time() - start_time
            
            ColorPrint.success(f"Processed 10,000 lines in {elapsed:.2f}s without memory issues")
            ColorPrint.info(f"Stats: {result['stats']['paragraphs']} paragraphs, {result['stats']['headings']} headings")
            return True
        except MemoryError:
            ColorPrint.error("Memory error occurred during processing")
            return False


class TestWordGeneration:
    """Test Word document generation"""
    
    def __init__(self):
        self.processor = DocumentProcessor()
        self.generator = WordGenerator()
    
    def test_word_generation(self):
        """Test Word document generation"""
        ColorPrint.header("TEST 15: Word Document Generation")
        
        sample_doc = """
SAMPLE DOCUMENT TITLE

INTRODUCTION

This is an introductory paragraph explaining the purpose of this document.

Background

The background section provides context for the reader.

• First key point
• Second key point
• Third key point

Definition: A term is defined here for clarity.

Table 1: Sample Data

| Name | Value | Description |
| Item A | 100 | First item |
| Item B | 200 | Second item |

CONCLUSION

This concludes the sample document.

REFERENCES

Smith, J. (2024). Sample Reference. Journal Name.
"""
        
        # Process document
        result_tuple = self.processor.process_text(sample_doc)
        if isinstance(result_tuple, tuple):
            result, _ = result_tuple
        else:
            result = result_tuple
        
        # Generate Word document
        output_path = os.path.join(os.path.dirname(__file__), 'test_output.docx')
        
        try:
            self.generator.generate(result['structured'], output_path)
            
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                ColorPrint.success(f"Word document generated: {output_path} ({file_size} bytes)")
                
                # Clean up
                os.remove(output_path)
                ColorPrint.info("Test file cleaned up")
                return True
            else:
                ColorPrint.error("Word document was not created")
                return False
        except Exception as e:
            ColorPrint.error(f"Word generation failed: {str(e)}")
            return False


def run_comprehensive_tests():
    """Run all tests"""
    print("\n" + "=" * 60)
    print(f"{ColorPrint.BOLD}{ColorPrint.BLUE}PATTERN-BASED DOCUMENT FORMATTER")
    print(f"COMPREHENSIVE TEST SUITE{ColorPrint.END}")
    print("=" * 60)
    
    all_passed = True
    total_passed = 0
    total_failed = 0
    
    # Test 1-9: Pattern Engine
    pattern_tests = TestPatternEngine()
    all_passed &= pattern_tests.run_all_tests()
    total_passed += pattern_tests.passed
    total_failed += pattern_tests.failed
    
    # Test 10-12: Document Processor
    processor_tests = TestDocumentProcessor()
    all_passed &= processor_tests.test_sample_document()
    all_passed &= processor_tests.test_edge_cases()
    all_passed &= processor_tests.test_no_duplication()
    total_passed += processor_tests.passed
    total_failed += processor_tests.failed
    
    # Test 13-14: Performance
    performance_tests = TestPerformance()
    all_passed &= performance_tests.test_speed()
    all_passed &= performance_tests.test_memory_efficiency()
    
    # Test 15: Word Generation
    word_tests = TestWordGeneration()
    all_passed &= word_tests.test_word_generation()
    
    # Final summary
    print("\n" + "=" * 60)
    print(f"{ColorPrint.BOLD}FINAL TEST SUMMARY{ColorPrint.END}")
    print("=" * 60)
    print(f"Pattern Tests Passed: {ColorPrint.GREEN}{total_passed}{ColorPrint.END}")
    print(f"Pattern Tests Failed: {ColorPrint.RED}{total_failed}{ColorPrint.END}")
    
    if total_passed + total_failed > 0:
        rate = (total_passed / (total_passed + total_failed)) * 100
        print(f"Success Rate: {rate:.1f}%")
    
    print("=" * 60)
    
    if all_passed:
        print(f"{ColorPrint.GREEN}{ColorPrint.BOLD}✓ ALL TESTS PASSED{ColorPrint.END}")
        print(f"{ColorPrint.GREEN}System is ready for production use!{ColorPrint.END}")
    else:
        print(f"{ColorPrint.YELLOW}{ColorPrint.BOLD}⚠ SOME TESTS HAD ISSUES{ColorPrint.END}")
        print(f"{ColorPrint.YELLOW}Review the results above for details{ColorPrint.END}")
    
    print("=" * 60 + "\n")
    
    return all_passed


def run_interactive_test():
    """Interactive test mode"""
    ColorPrint.header("INTERACTIVE TEST MODE")
    print("Enter text to analyze (or 'quit' to exit):\n")
    
    engine = PatternEngine()
    
    while True:
        try:
            line = input(f"{ColorPrint.BLUE}> {ColorPrint.END}")
            
            if line.lower() in ['quit', 'exit', 'q']:
                break
            
            if not line.strip():
                continue
            
            result = engine.analyze_line(line, 0)
            
            print(f"\n  Type: {ColorPrint.BOLD}{result['type']}{ColorPrint.END}")
            if result.get('level'):
                print(f"  Level: {result['level']}")
            print(f"  Confidence: {result['confidence']*100:.0f}%")
            print(f"  Content: {result['content'][:80]}")
            if result.get('term'):
                print(f"  Term: {result['term']}")
            if result.get('cells'):
                print(f"  Cells: {result['cells']}")
            print()
            
        except KeyboardInterrupt:
            break
        except EOFError:
            break
    
    ColorPrint.info("Exiting interactive mode...")


def create_sample_documents():
    """Create sample test documents"""
    ColorPrint.header("CREATING SAMPLE DOCUMENTS")
    
    tests_dir = os.path.dirname(__file__)
    
    # Sample 1: Academic Paper
    academic_sample = """CLOUD COMPUTING AND ECONOMIES OF SCALE
A Comprehensive Analysis

ABSTRACT

This paper examines the economic principles underlying cloud computing infrastructure and its impact on enterprise IT delivery models.

INTRODUCTION

Background and Context

Cloud computing has transformed enterprise IT delivery models. Key innovations include:

• Infrastructure as a Service (IaaS)
• Platform as a Service (PaaS)
• Software as a Service (SaaS)
• Function as a Service (FaaS)

Definition: Cloud computing refers to on-demand delivery of computing resources over the internet with pay-as-you-go pricing.

1.1 Economies of Scale

Three primary mechanisms drive cost advantages in cloud computing:

1. Resource consolidation across multiple tenants
2. Bulk purchasing power for hardware and facilities
3. Operational automation reducing labor costs

1.2 Market Analysis

The cloud computing market has shown significant growth:

[TABLE START]
| Year | Market Size | Growth Rate |
| 2022 | $480B | 20% |
| 2023 | $590B | 23% |
| 2024 | $720B | 22% |
[TABLE END]

Figure 1: Cloud market growth trajectory

METHODOLOGY

Research Approach

This study employed a mixed-methods approach combining:

a) Quantitative analysis of pricing data from major providers
b) Qualitative interviews with enterprise cloud architects
c) Case study analysis of cloud migration projects

Data Collection

Primary data was collected through:

1. Survey of 500 enterprise IT managers
2. Analysis of public pricing APIs
3. Review of financial reports

RESULTS

Key Findings

The research revealed several important patterns:

Objective: Determine the cost savings achieved through cloud adoption.

Result: Organizations reported an average of 35% reduction in IT infrastructure costs after cloud migration.

Note: Results varied significantly based on organization size and industry.

Performance Comparison

Table 2: Cost Comparison by Provider

| Provider | Compute Cost | Storage Cost | Overall Rating |
| AWS | $0.10/hr | $0.023/GB | 4.5/5 |
| Azure | $0.09/hr | $0.018/GB | 4.3/5 |
| GCP | $0.08/hr | $0.020/GB | 4.4/5 |

DISCUSSION

Analysis and Interpretation

The findings support the hypothesis that cloud computing delivers measurable economic benefits through economies of scale.

Key Point: Larger organizations tend to realize greater cost savings due to their ability to leverage reserved capacity pricing.

Limitations

Several limitations should be considered:

- Sample size was limited to North American enterprises
- Pricing data may not reflect negotiated enterprise discounts
- Rapidly changing market conditions affect generalizability

CONCLUSION

Summary of Findings

Cloud computing provides significant economic benefits for enterprises of all sizes. The primary drivers of cost savings include:

• Reduced capital expenditure requirements
• Lower operational overhead
• Improved resource utilization
• Access to enterprise-grade infrastructure

Future Work

Additional research is needed to examine:

1. Long-term cost trends in cloud computing
2. Impact of edge computing on economies of scale
3. Environmental sustainability of cloud infrastructure

REFERENCES

Smith, J. (2024). Cloud Economics: A Comprehensive Guide. Tech Publishing.
Johnson, A. et al. 2023. Scaling strategies for enterprise cloud adoption. Journal of Cloud Computing.
Brown, M. & Wilson, K. (2024). Cost optimization in multi-cloud environments. IEEE Transactions.
Garcia, L. Retrieved from https://cloudresearch.org/economics-of-scale
[1] Chen, W. (2023). Infrastructure as a Service pricing models. ACM Computing Surveys.
Williams, R. (2024). The future of enterprise IT. Harvard Business Review.
"""
    
    with open(os.path.join(tests_dir, 'sample_academic_paper.txt'), 'w', encoding='utf-8') as f:
        f.write(academic_sample)
    
    ColorPrint.success("Created: sample_academic_paper.txt")
    
    # Sample 2: Business Report
    business_sample = """QUARTERLY BUSINESS REPORT
Q4 2024 Performance Review

EXECUTIVE SUMMARY

Q4 2024 exceeded expectations with 15% year-over-year revenue growth and improved operational efficiency across all business units.

FINANCIAL PERFORMANCE

Revenue Breakdown

The company achieved the following results in Q4:

1. Product sales: $5M (up 20% YoY)
2. Professional services: $3M (up 10% YoY)
3. Subscription revenue: $2M (up 25% YoY)

Objective: Maintain revenue growth trajectory while improving margins.

Result: Operating margin improved from 18% to 22%.

Key Metrics

Table 1: Financial Summary

| Metric | Q4 2024 | Q4 2023 | Change |
| Revenue | $10M | $8.7M | +15% |
| Gross Profit | $6.5M | $5.4M | +20% |
| Operating Income | $2.2M | $1.6M | +38% |
| Net Income | $1.8M | $1.2M | +50% |

MARKET ANALYSIS

Industry Trends

Key trends observed in the market:

• Digital transformation acceleration across industries
• Increased cloud adoption among mid-market companies
• Growing demand for AI and automation solutions
- Shift towards subscription-based business models

Competitive Landscape

Our market position has strengthened:

a) Gained 3% market share in core segment
b) Expanded into two new geographic markets
c) Launched three new product offerings

OPERATIONAL HIGHLIGHTS

Team Performance

The organization achieved several milestones:

Definition: Employee Net Promoter Score (eNPS) measures employee satisfaction and loyalty.

• Employee satisfaction score: 78 (up from 72)
• Customer retention rate: 94%
• New customer acquisition: 150 enterprise accounts

Process Improvements

1. Reduced order-to-delivery time by 25%
2. Implemented automated quality control
3. Launched self-service customer portal

RECOMMENDATIONS

Strategic Initiatives

Priority actions for Q1 2025:

1) Expand product portfolio with AI-powered features
2) Enhance customer success program
3) Invest in R&D for next-generation platform
4) Strengthen partnerships in key markets

Goal: Achieve 20% revenue growth in 2025.

Resource Allocation

Recommended budget distribution:

| Category | Allocation | Priority |
| R&D | 25% | High |
| Sales & Marketing | 35% | High |
| Operations | 25% | Medium |
| G&A | 15% | Medium |

CONCLUSION

Strong foundation established for continued growth. Focus areas for next quarter include:

- Accelerating product innovation
- Expanding market presence
- Enhancing customer experience
- Building operational excellence

Note: Detailed appendices are available upon request.

APPENDIX

Figure 1: Revenue trend chart
Figure 2: Customer satisfaction scores
Figure 3: Market share analysis
"""
    
    with open(os.path.join(tests_dir, 'sample_business_report.txt'), 'w', encoding='utf-8') as f:
        f.write(business_sample)
    
    ColorPrint.success("Created: sample_business_report.txt")
    
    ColorPrint.info("\nSample documents created in tests folder!")
    ColorPrint.info("Use these to test the formatter application.")


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == 'interactive':
            run_interactive_test()
        elif sys.argv[1] == 'samples':
            create_sample_documents()
        elif sys.argv[1] == 'help':
            print("Usage:")
            print("  python test_pattern_formatter.py           # Run all tests")
            print("  python test_pattern_formatter.py interactive # Interactive mode")
            print("  python test_pattern_formatter.py samples     # Create sample documents")
            print("  python test_pattern_formatter.py help        # Show this help")
        else:
            print(f"Unknown command: {sys.argv[1]}")
            print("Use 'help' for usage information.")
    else:
        # Create sample documents first
        create_sample_documents()
        
        # Run comprehensive tests
        success = run_comprehensive_tests()
        sys.exit(0 if success else 1)
