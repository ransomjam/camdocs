import os
import sys
from coverpage_generator import generate_cover_page

# Test cases covering different templates and scenarios
test_cases = [
    {
        "name": "Assignment - Science",
        "data": {
            "documentType": "Assignment",
            "institution": "The University of Bamenda",
            "faculty": "Faculty of Science",
            "department": "Computer Science",
            "courseCode": "COMP301",
            "courseTitle": "Data Structures",
            "title": "AVL Tree Implementation",
            "studentName": "John Doe",
            "studentId": "UBa21E001",
            "instructor": "Dr. Smith",
            "date": "2025-01-15",
            "level": "300 Level",
            "assignmentNumber": "1"
        }
    },
    {
        "name": "Project Report - Engineering",
        "data": {
            "documentType": "Project Report",
            "institution": "The University of Bamenda",
            "faculty": "National Higher Polytechnic Institute",
            "department": "Computer Engineering",
            "title": "Smart Home Automation System",
            "studentName": "Jane Smith",
            "studentId": "UBa21E002",
            # "supervisor": "Prof. Johnson", # REMOVED to test fallback from academicSupervisor
            "academicSupervisor": "Dr. Academic Sup",
            "fieldSupervisor": "Engr. Field Sup",
            "date": "2025-06-20",
            "level": "400 Level"
        }
    },
    {
        "name": "Dissertation - Arts",
        "data": {
            "documentType": "Dissertation",
            "institution": "The University of Bamenda",
            "faculty": "Faculty of Arts",
            "department": "English",
            "title": "Modern Linguistics in Cameroon",
            "studentName": "Alice Brown",
            "studentId": "UBa20A005",
            "supervisor": "Dr. Williams",
            "date": "2025-11-10"
        }
    }
]

print("Starting Cover Page Generation Tests...")
print("=" * 60)

success_count = 0
for case in test_cases:
    print(f"\nTesting: {case['name']}")
    print("-" * 30)
    
    data = case['data']
    output_path, error = generate_cover_page(data)
    
    if error:
        print(f"❌ FAILED: {error}")
    else:
        if os.path.exists(output_path):
            print(f"✅ SUCCESS")
            print(f"   File: {os.path.basename(output_path)}")
            success_count += 1
        else:
            print(f"❌ FAILED: File not found at {output_path}")

print("\n" + "=" * 60)
print(f"Test Summary: {success_count}/{len(test_cases)} passed")

if success_count == len(test_cases):
    sys.exit(0)
else:
    sys.exit(1)
