# ``` Drug Naming Quiz```
Overview
The Drug Naming Quiz is an educational application that tests users' knowledge of drug's Brand/Generic name. The user is presented with a drug name and must identify its corresponding Brand/Generic name. The quiz is powered by Python, utilizing Tkinter for the user interface and OpenPyXL for reading drug data from an Excel file.

## Features
1. User-friendly GUI with Tkinter.
2. Excel-based drug data for dynamic quiz content.
3. Real-time score tracking throughout the quiz.
4. History of previous questions displayed during the quiz.
5. Options to retry or exit the quiz after completion.

## Requirements
1. Python 3.x
2. Tkinter (comes pre-installed with Python)
3. OpenPyXL (for working with Excel files)

## Installation
1. Install OpenPyXL:
- bash
- Copy code
- pip install openpyxl
- Prepare the Excel file: The quiz relies on an Excel file (e.g., drugnames.xlsx) with the following structure:

- Column A: Drug Name
- Column B: Brand Name
- Column C: Generic Name
- Ensure that the drug data starts from the second row.

## Usage
1. Run the Script: To start the quiz, run the Python script:
- bash
- Copy code
- python drug_quiz.py

2. Start the Quiz:
- Click Start Quiz to select a sheet from the Excel file.
- The quiz will display a series of questions asking for the Brand/Generic name of a drug.

3. Answering Questions:
- Type your answer in the input field and click Submit Answer.
The application will inform you if your answer is correct or incorrect.

4. End of Quiz:
After completing the quiz, your score is displayed.
You can choose to retry with another sheet or exit the application.
