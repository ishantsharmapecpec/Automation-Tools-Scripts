# Automation-Tools-Scripts
Script 1

Automating 6x6 Stiffness Matrix Extraction from PDFs
Overview:
Task: Extract and consolidate 6x6 stiffness matrices from over 100 PDFs into an Excel sheet.
Method: Developed a VBA script to automate the extraction process, significantly reducing time and manual effort.

Script Functionality:
Automated Extraction: The script loops through each PDF in a specified directory, extracts the stiffness matrix, and populates the corresponding rows in an Excel sheet.
Data Accuracy: The script precisely locate and extract relevant data from PDF content.
Dynamic Processing: Adjusts automatically to different formats or variations within the PDFs.

Time Efficiency:
Manual Process: Extracting data from 100 PDFs manually could take several hours, depending on the complexity and format variations.
Script Execution: Completed the entire extraction in just 10 minutes, significantly reducing time-to-completion.

Benefits :
Consistency: Eliminates human errors such as incorrect data entry or missing rows.
Speed: Processes data at a much faster rate than manual extraction.
Scalability: Can be easily adapted to process additional PDFs or expanded to handle other data types with minimal modification.
Reusability: The script can be reused for future projects, ensuring consistency across different datasets.

Script 2

a) Overview:
Objective: Automate the integration of load combinations from Excel into Repute files.
Current Process: The script loops through each Repute file in a specified directory, populate the load cases for each file from the spread sheet and make them ready for carrying out the further analysis. Updated 882 files each, with 6 load cases per file and each requiring 10 elements.

(b) Script Functionality
1. Automated Processing:
Batch Processing: Handles multiple files and load combinations in bulk.
Data Integration: Reads load combinations and parameters directly from an Excel sheet.

2. XML Manipulation:
Dynamic Updates: Adjusts XML content based on Excel data.
Element Duplication: Efficiently duplicates and inserts elements as needed.

3. Error Handling
Validation: Checks and warns if expected patterns or values are missing. Reduces manual entry errors, ensuring consistent data integration.
Robust Parsing: Handles various XML structures and updates accordingly

(C)Time Efficiency:
Manual Process: Coping these load cases with this much detail can easily take weeks depending on the complexity and format variations.
Script Execution: Completed the entire extraction in just less than 1 minutes, significantly reducing time-to-completion with 100% accuracy. 

Script 3

(a) Overview
Objective: Streamline the process of extracting and compiling pile and pile cap responses for engineering analysis.
Current Challenge: Manual extraction of data is time-consuming, prone to errors, and not scalable for large datasets.
Current Process: The script loops through each Repute file in a specified directory and extract the results and save them in a given folder.

(b) The Script Overview
Purpose: Automatically extract pile and pile cap response data, compile the results, and rename files accurately for further use
Python: Used for scripting and automation.
PyAutoGUI: Controls the mouse and keyboard to interact with the software, simulating manual tasks.

(c)Time Efficiency
Manual Process: Extracting data manually would take 0.5-1 hours per repute file. Which is prone to human errors, such as incorrect file naming or missing data points.
Automated Script: Completes the same task in 6 minutes per file, allowing engineers to focus on analysis rather than data extraction

(d) Scalability
Manual Process: Becomes increasingly unmanageable with larger datasets.
Automated Script: Easily handles large volumes of data.

Script 4

(a) Overview
Objective: Automate the process of compiling pile group analysis results to improve efficiency and accuracy.
Current Challenge: Manually handling over 10,000 extracted pile and pile cap response is not feasible and is highly error-prone.

(b) The Script Overview
Purpose: To automatically compile results from multiple output files into a single, consolidated sheet for each pile group.
OpenPyXL: Python library for reading and writing Excel files, allowing for efficient data manipulation and copying.

(c) How the Script Works
Input Files: The script reads from over 10,000 extracted files from Repute.
Target Output: Compiles data into 882 Excel files, with each pile group’s results summarized on a single sheet

(d) Automation Flow:
Search & Identify: The script identifies and loads the relevant Excel files based on a naming convention from a master file.
Data Transfer: Data for up to 6 load cases per pile group is copied from individual output files into a corresponding global summary sheet.
Error Handling: Includes checks for potential errors (e.g., missing sheets or files) and handles them to ensure a smooth run.



