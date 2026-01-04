import sys
import os
import json
from docx import Document

# Add backend to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'backend'))

from pattern_formatter_backend import QuestionnaireProcessor, format_questionnaire_in_word

def test_generation():
    text = """Household Income Assessment Questionnaire
Introduction
Thank you for participating in this household income assessment survey. Your responses will remain confidential and will be used only for research/assessment purposes. This questionnaire should be completed by the primary income earner or someone knowledgeable about the household's financial situation.

Section 1: Household Demographics
How many people live in your household?

1

2

3

4

5 or more

How many household members contribute to income?

1

2

3

4 or more

What is the highest education level in your household?

Less than high school

High school graduate

Some college/vocational training

Bachelor's degree

Graduate/professional degree

Section 2: Income Sources
Which of the following income sources does your household receive? (Select all that apply)

Employment wages/salaries

Self-employment/business income

Retirement/pension

Social Security/disability benefits

Investment income (dividends, interest, rentals)

Government assistance (SNAP, TANF, WIC)

Child support/alimony

Other: ________________

What is your approximate total monthly household income (before taxes)?

Less than $1,500

$1,501 - $3,000

$3,001 - $5,000

$5,001 - $7,500

$7,501 - $10,000

$10,001 - $15,000

More than $15,000

Section 3: Financial Security Perceptions (Likert Scale Tables)
Instructions: Please indicate how strongly you agree or disagree with each statement about your household's financial situation.

Table 1: Income Adequacy
Statement	Strongly Disagree	Disagree	Neutral	Agree	Strongly Agree
Our household income is sufficient to cover our basic needs (food, housing, utilities)	[ ]	[ ]	[ ]	[ ]	[ ]
We have enough income for discretionary spending (entertainment, dining out, hobbies)	[ ]	[ ]	[ ]	[ ]	[ ]
Our income allows us to save money regularly	[ ]	[ ]	[ ]	[ ]	[ ]
We could handle an unexpected $500 expense without financial hardship	[ ]	[ ]	[ ]	[ ]	[ ]
Table 2: Income Stability
Statement	Strongly Disagree	Disagree	Neutral	Agree	Strongly Agree
Our household income is predictable from month to month	[ ]	[ ]	[ ]	[ ]	[ ]
We are confident our main income sources will continue	[ ]	[ ]	[ ]	[ ]	[ ]
Our income has kept pace with inflation/cost of living	[ ]	[ ]	[ ]	[ ]	[ ]
We have multiple reliable income sources	[ ]	[ ]	[ ]	[ ]	[ ]
Table 3: Financial Stress
Statement	Never	Rarely	Sometimes	Often	Always
We worry about having enough money to pay bills	[ ]	[ ]	[ ]	[ ]	[ ]
We delay medical/dental care due to cost	[ ]	[ ]	[ ]	[ ]	[ ]
We have to borrow money or use credit for basic expenses	[ ]	[ ]	[ ]	[ ]	[ ]
Financial concerns cause tension in our household	[ ]	[ ]	[ ]	[ ]	[ ]
Section 4: Expenses and Budgeting
What percentage of your monthly income goes toward housing (rent/mortgage, taxes, insurance)?

Less than 25%

26-35%

36-50%

More than 50%

How would you describe your household's budgeting practices?

We follow a detailed written budget

We have a mental budget we generally follow

We track expenses but don't set specific limits

We don't track expenses or use a budget

Table 4: Budget Management
Statement	Strongly Disagree	Disagree	Neutral	Agree	Strongly Agree
We regularly review our household expenses	[ ]	[ ]	[ ]	[ ]	[ ]
We have a plan to reduce debt	[ ]	[ ]	[ ]	[ ]	[ ]
We allocate funds for savings before other spending	[ ]	[ ]	[ ]	[ ]	[ ]
We discuss financial goals as a household	[ ]	[ ]	[ ]	[ ]	[ ]
Section 5: Future Outlook and Support
Compared to last year, your household's financial situation is:

Much worse

Somewhat worse

About the same

Somewhat better

Much better

What is your household's primary financial goal for the next year? (Select one)

Pay down debt

Build emergency savings

Save for a major purchase (home, car, education)

Increase retirement savings

Maintain current situation

Other: ________________

Has your household accessed any financial assistance or counseling in the past 2 years?

Yes, government assistance programs

Yes, nonprofit/community organizations

Yes, private financial advisor

No, we have not sought assistance

Prefer not to answer

Section 6: Additional Comments
Is there anything else you would like to share about your household's financial situation, challenges, or successes?

Section 7: Optional Contact Information
(Only complete if you're open to follow-up or resource sharing)

Name: ________________________

Email/Phone: ________________________

Best time to contact: ________________________

Thank you for completing this questionnaire!
Your responses will help us better understand household financial situations and develop appropriate resources.
"""
    
    processor = QuestionnaireProcessor()
    
    # 1. Detect
    detection_result = processor.detect_questionnaire(text)
    print(f"Is Questionnaire: {detection_result['is_questionnaire']}")
    
    if detection_result['is_questionnaire']:
        # 2. Parse
        structure = processor.parse_questionnaire_structure(text)
        print("\n--- Structure ---")
        # print(json.dumps(structure, indent=2))
        
        # Check specific issues
        print("\nChecking Likert Tables:")
        for section in structure['sections']:
            for q in section['questions']:
                if q['type'] == 'likert_table':
                    print(f"Found Likert Table: {q['text']}")
                    print(f"  Scale Items: {q.get('scale', {}).get('items')}")
                    print(f"  Sub-questions: {len(q.get('sub_questions', []))}")
                    for sq in q.get('sub_questions', []):
                        if '[ ]' in sq:
                            print(f"  WARNING: '[ ]' found in sub-question text: {sq}")
        
        # 3. Generate
        doc = Document()
        doc = format_questionnaire_in_word(doc, structure)
        output_path = os.path.join(os.path.dirname(__file__), 'test_output.docx')
        doc.save(output_path)
        print(f"\nGenerated document saved to {output_path}")

if __name__ == "__main__":
    test_generation()
