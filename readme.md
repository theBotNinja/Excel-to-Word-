# 📄 Excel to Word Form Generator – Documentation

## 📝 Overview

This script converts Excel sheets into well-structured Word documents suitable for creating **survey forms or questionnaires**. It allows users to select:

- A column of questions from an Excel file.
- The type of multiple-choice options (Likert scales, frequency, satisfaction, etc.).
- The starting row of questions.
- And finally, it generates `.docx` Word files—one for each sheet in the Excel file.

---

## 🚀 Features

- Interactive command-line interface using colors.
- Supports multiple scales for answers (3-point, 5-point, Satisfaction, etc.).
- Generates a separate Word file per sheet.
- Option to work in **"One-by-One"** mode (select column for each sheet).
- Optionally auto-selects column/sheet for faster conversion.

---

## 🔧 Requirements

Install dependencies using pip:

```bash
pip install pandas python-docx
```

### Built-in modules used:

- sys
- os
- tkinter
- time

### File Structure
```plaintext
main.py               # The main script file
word_files/           # Output folder (auto-created) for generated .docx files
your_excel_file.xlsx  # The Excel file with questions
```
---
## How to Run the Script

### Option 1: With File Argument
```bash
python main.py path/to/your_excel_file.xlsx
```

### Option 2: Without File Argument
This will open a file picker to select the Excel file.
```bash
python main.py
```
---
## Usage Flow
### After launch:
- You are prompted to choose the Excel file (if not passed via command-line).
- The script reads all sheet names from the file.
- You’re given the choice:
- Press Enter to apply the same column to all sheets.
- Press Y to enable "One-by-One" column selection for each sheet.
- Type about to see script author info.
### For each sheet:
- Select a column (questions).
- Choose a row index to start from.
- Select a predefined set of options.
- Generates Word documents saved inside the word_files folder.
---
##  Customization
### You can easily change:

- The heading for each document (currently uses the sheet name).
- Option templates inside the SelectOptions() function.
- Output folder path in the create_questionnaire() function.

### Example Options Provided
- 3-Point Scale: Agree, Neutral, Disagree
- 5-Point Scale: Strongly Agree, Agree, Neutral, Disagree, Strongly Disagree
- Frequency Scale: Always, Often, Sometimes, Rarely, Never
- Importance Scale: Very Important, Important, Neutral, Unimportant, Very Unimportant
- Satisfaction Scale: Very Satisfied, Satisfied, Neutral, Dissatisfied, Very Dissatisfied

