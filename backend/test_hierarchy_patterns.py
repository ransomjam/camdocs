import unittest
import re
from pattern_formatter_backend import HierarchyCorrector

class TestHierarchyCorrector(unittest.TestCase):
    def setUp(self):
        self.corrector = HierarchyCorrector()

    def test_week_day_pattern(self):
        text = "3.6 WEEK TWO\n3.7 DAY ONE"
        expected = "3.6 WEEK TWO\n3.6.1 DAY ONE"
        self.assertEqual(self.corrector.correct(text), expected)

    def test_unit_lesson_pattern(self):
        text = "4.2 UNIT THREE\n4.3 LESSON ONE"
        expected = "4.2 UNIT THREE\n4.2.1 LESSON ONE"
        self.assertEqual(self.corrector.correct(text), expected)

    def test_chapter_section_pattern(self):
        text = "1.0 CHAPTER ONE\n1.1 SECTION ONE"
        expected = "1.0 CHAPTER ONE\n1.0.1 SECTION ONE"
        self.assertEqual(self.corrector.correct(text), expected)

    def test_module_topic_pattern(self):
        text = "2.1 MODULE FIVE\n2.2 TOPIC THREE"
        expected = "2.1 MODULE FIVE\n2.1.1 TOPIC THREE"
        self.assertEqual(self.corrector.correct(text), expected)

    def test_lettered_hierarchy(self):
        text = "2.1 PART A\n2.2 1. Introduction"
        expected = "2.1 PART A\n2.1.1 1. Introduction"
        self.assertEqual(self.corrector.correct(text), expected)

    def test_general_specific_pattern(self):
        text = "3.6 THEORY\n3.7 IMPLEMENTATION STRATEGIES"
        expected = "3.6 THEORY\n3.6.1 IMPLEMENTATION STRATEGIES"
        self.assertEqual(self.corrector.correct(text), expected)

    def test_category_subcategory_pattern(self):
        text = "6.4 TYPES\n6.5 TYPE A"
        expected = "6.4 TYPES\n6.4.1 TYPE A"
        self.assertEqual(self.corrector.correct(text), expected)

    def test_correct_lines_logic(self):
        lines = [
            "3.6 WEEK TWO",
            "3.7 DAY ONE",
            "4.0 OTHER"
        ]
        expected = [
            "3.6 WEEK TWO",
            "3.6.1 DAY ONE",
            "4.0 OTHER"
        ]
        result = self.corrector.correct_lines(lines)
        self.assertEqual(result, expected)

    def test_correct_lines_semantic_pairs(self):
        # Test pairs defined in HIERARCHICAL_PAIRS
        lines = [
            "1.1 ANALYSIS",
            "1.2 FINDING 1",
            "2.1 METHOD",
            "2.2 STEP 1"
        ]
        expected = [
            "1.1 ANALYSIS",
            "1.1.1 FINDING 1",
            "2.1 METHOD",
            "2.1.1 STEP 1"
        ]
        result = self.corrector.correct_lines(lines)
        self.assertEqual(result, expected)

if __name__ == '__main__':
    print("STARTING HIERARCHY PATTERN TESTS")
    unittest.main(exit=False)
    print("FINISHED HIERARCHY PATTERN TESTS")
