# Create Exams
###### 帮李泽文同学写的小小程序 :sparkles:
Auto analyze and generate exams from a bunch of docx study materials

Materials only have 4 types of questions:
1. Single Choice
2. Multiple Choice
3. Filling blanks
4. (optional) Correct or Incorrect

For each question:
- Each question is a line, including the parenthess
- A question always starts with a number
- The next line must be the answer

Procedure:
1. Read docx line by line, categorize each question and put them into an XML
2. Load question database (XML) and use random sample to pick some questions
3. Write questions into docx and save both the exam and it's answers with timestamp
