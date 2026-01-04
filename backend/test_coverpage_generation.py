import os
import sys
from coverpage_generator import generate_cover_page

# Mock data based on user request
data = {
    "documentType": "Assignment",
    "institution": "The University of Bamenda",
    "faculty": "Faculty of Science",
    "department": "Computer Science",
    "courseCode": "COMP301",
    "courseTitle": "Data Structures and Algorithms",
    "title": "Implementation of AVL Trees",
    "studentName": "John Doe",
    "studentId": "UBa21E001",
    "instructor": "Dr. Smith",
    "date": "2025-01-15",
    "level": "300 Level",
    "assignmentNumber": "2"
}

print("Generating cover page with data:")
for k, v in data.items():
    print(f"  {k}: {v}")

output_path, error = generate_cover_page(data)

if error:
    print(f"Error: {error}")
    sys.exit(1)
else:
    print(f"Success! Generated file at: {output_path}")
    
    # Verify file exists
    if os.path.exists(output_path):
        print("File verified on disk.")
    else:
        print("File not found on disk!")
