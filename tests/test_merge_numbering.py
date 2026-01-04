import unittest
import sys
import os

# Add backend to path
sys.path.append(os.path.join(os.getcwd(), 'backend'))

from pattern_formatter_backend import QuestionnaireProcessor

class TestQuestionNumbering(unittest.TestCase):
    def setUp(self):
        self.processor = QuestionnaireProcessor()

    def test_standard_numbered_question(self):
        line = "1. What is your name?"
        result = self.processor.detect_question(line)
        self.assertIsNotNone(result)
        self.assertEqual(result['number'], "1")
        self.assertEqual(result['text'], "What is your name?")

    def test_parenthesis_numbered_question(self):
        line = "2) How old are you?"
        result = self.processor.detect_question(line)
        self.assertIsNotNone(result)
        self.assertEqual(result['number'], "2")
        self.assertEqual(result['text'], "How old are you?")

    def test_unnumbered_question(self):
        line = "What is your favorite color?"
        result = self.processor.detect_question(line)
        self.assertIsNotNone(result)
        self.assertEqual(result['number'], "")
        self.assertEqual(result['text'], "What is your favorite color?")

    def test_colon_ending_question(self):
        line = "Description of incident:"
        result = self.processor.detect_question(line)
        self.assertIsNotNone(result)
        self.assertEqual(result['number'], "")
        self.assertEqual(result['text'], "Description of incident:")

    def test_numbered_statement_colon(self):
        # This currently might fail if regex requires '?'
        line = "3. Description:"
        result = self.processor.detect_question(line)
        self.assertIsNotNone(result, "Should detect numbered statement ending in colon")
        self.assertEqual(result['number'], "3")
        self.assertEqual(result['text'], "Description:")

    def test_merge_numbering_logic(self):
        # Simulate the logic in format_questionnaire_in_word
        q_num = "1"
        q_text = "What is your name?"
        full_text = f"{q_num} {q_text}".strip() if q_num else q_text
        self.assertEqual(full_text, "1 What is your name?")

        q_num = ""
        q_text = "What is your name?"
        full_text = f"{q_num} {q_text}".strip() if q_num else q_text
        self.assertEqual(full_text, "What is your name?")

if __name__ == '__main__':
    unittest.main()
