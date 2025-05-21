AC Report Validation Script
===========================

Overview
--------

This R script performs a validation check of the annual animal experimentation statistics based on the official plausibility check rules of the Federal Food Safety and Veterinary Office (FSVO).

It parses and validates .xlsx exports from Animex's AC Reports, ensuring all logical dependencies between statistical chapters are respected. The output is a color-coded Excel file showing failed rules per animal category and per AC report.

IMPORTANT:
If any of the validation rules fail, it indicates that the AC report has not been filled out correctly. In such cases, the report will be returned to you for corrections - either by your Cantonal Veterinary Office or by the FSVO.

To avoid this, we strongly recommend running this validation script BEFORE submitting your AC Reports. You can do this by exporting the report in Draft mode. Simply open the AC report in Animex, press the PDF button in the top right corner, and retrieve the automatically generated .xlsx file. This file can be validated with the script to catch and fix errors.

Note:
This plausibility check script does not cover errors of a subject-matter nature. For example:

Animals categorized into the wrong severity degrees
Duplicates submitted across similar animal cards
Animals reported as experimental that were actually used in breeding

These types of errors still require careful review and expert judgment by the reporting lab or by the Cantonal Veterinary Office.



Features
--------

- Validates AC report chapter logic using the following plausibility / validation rules:
- Rule 1: 9.2 needs to be greater than 9.4.1
- Rule 2: 10.1 must equal the sum of 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4 OR 10.1 must equal to the sum of 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4 + 9.4.2  OR 10.1 must equal to the sum of 9.4.2
- Rule 3 11.1 + 11.2 + 11.3 + 11.4 must equal to the sum of 9.2 + 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4
- Rule 4: 12.1 + 12.2.1 + 12.2.2 + 12.2.3 + 12.2.4 + 12.3 must equal to the sum of 10.1 + 10.2 + 10.3 + 10.4
- Rule 5: If 9.4.1 is bigger than 0 then 9.2 must be greater than the sum of 10.1 to 10.4


Usage Instructions
------------------

1. Export from Animex

   - Open your lab's AC Report in Animex
   - In the top-right corner, press the PDF button
   - This will automatically download two files:
     - A .pdf file of the report
     - A .xlsx Excel file containing the raw statistics

   Important: This script only uses the .xlsx files. You can save multiple AC reports *.xlsx files in the same folder. 

2. Prepare Input Folder

   - Create a folder on your machine (e.g. C:/Rtests/statistics/animal_excels/)
   - Copy all .xlsx files from step 1 into this folder
   - In the R script, set the path to this folder on line 27:

     input_dir <- "C:/Rtests/statistics/animal_excels"  # <-- update this to your folder

3. Run the Script

   - Open RStudio or R terminal
   - Source or run the script
   - After processing, the script generates:

     Validation_Report.xlsx

   This file will be saved in the path defined in output_file (also editable).
   output_file <- "C:/Rtests/statistics/Validation_Report.xlsx" # <-- alsoupdate this to your folder


Output
------

- Excel file: Validation_Report.xlsx
- Sheet: Validation_Results
- Columns: File name, Animal Category, Failed Rule, all chapter values
- Conditional formatting for visual feedback per rule

Author
------

Cristian Berce  
Federal Food Safety and Veterinary Office (FSVO)  
Switzerland

License
-------

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
