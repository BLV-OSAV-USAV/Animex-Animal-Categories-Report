# AC Report Validation Script

## 

## Overview



**Important note on scope and limitations**



This repository provides an R script intended to support the pre-submission plausibility checking of Animex-generated AC Excel reports. It implements a defined subset of arithmetic plausibility checks used in the context of the annual animal experimentation statistics to help identify common internal inconsistencies between selected chapters.



This tool does not constitute an official FSVO validation system and does not confer any formal approval or legal standing. It does not replace the official review performed by Cantonal Veterinary Offices or by the FSVO, and a “valid” result does not guarantee acceptance of a submitted report.



The script is tightly coupled to the current Animex AC Excel export structure (including sheet name, chapter numbering, and INST markers). Changes to the export format, missing values, or deviations from the expected layout may lead to incomplete or misleading results. Users remain fully responsible for the correctness and completeness of their data and for ensuring compliance with all applicable legal and procedural requirements.



No individual support or troubleshooting is provided. Issues may be reported via GitHub; however, responses are not guaranteed.



**Usage recommendation**



To reduce the risk of avoidable corrections, we strongly recommend running this script before submitting AC reports. This can be done by exporting the AC report from Animex in Draft mode: open the report in Animex, use the PDF button in the top right corner to generate the .xlsx file, and run the validation script on that file to identify potential inconsistencies early.



Note on subject-matter checks



This plausibility check does not cover errors of a subject-matter nature, such as:

* animals assigned to incorrect severity degrees
* duplicate reporting across similar animal cards
* animals reported as experimental when they were used exclusively for breeding



Such issues require careful expert review by the reporting institution and, where applicable, by the Cantonal Veterinary Office.



## Features

* Validates AC report chapter logic using the following plausibility / validation rules:
* Rule 1: 9.2 needs to be greater than 9.4.1
* Rule 2: 10.1 must equal the sum of 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4 OR 10.1 must equal to the sum of 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4 + 9.4.2  OR 10.1 must equal to the sum of 9.4.2
* Rule 3 11.1 + 11.2 + 11.3 + 11.4 must equal to the sum of 9.2 + 9.1.1 + 9.1.2 + 9.1.3 + 9.1.4
* Rule 4: 12.1 + 12.2.1 + 12.2.2 + 12.2.3 + 12.2.4 + 12.3 must equal to the sum of 10.1 + 10.2 + 10.3 + 10.4
* Rule 5: If 9.4.1 is bigger than 0 then 9.2 must be greater than the sum of 10.1 to 10.4



## Usage Instructions

1. Export from Animex

   * Open your lab's AC Report in Animex
   * In the top-right corner, press the PDF button
   * This will automatically download two files:

     * A .pdf file of the report
     * A .xlsx Excel file containing the raw statistics

   Important: This script only uses the .xlsx files. You can save multiple AC reports \*.xlsx files in the same folder.

2. Prepare Input Folder

   * Create a folder on your machine (e.g. C:/Rtests/statistics/animal\_excels/)
   * Copy all .xlsx files from step 1 into this folder
   * In the R script, set the path to this folder on line 27:

     input\_dir <- "C:/Rtests/statistics/animal\_excels"  # <-- update this to your folder

3. Run the Script

   * Open RStudio or R terminal
   * Source or run the script
   * After processing, the script generates:

     Validation\_Report.xlsx

   This file will be saved in the path defined in output\_file (also editable).
   output\_file <- "C:/Rtests/statistics/Validation\_Report.xlsx" # <-- alsoupdate this to your folder

   

   ## Output

* Excel file: Validation\_Report.xlsx
* Sheet: Validation\_Results
* Columns: File name, Animal Category, Failed Rule, all chapter values
* Conditional formatting for visual feedback per rule



  ## License

  Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

