# 
## Description
This code is designed to compare all the attributes from an XML file with their corresponding lines in an Excel file. Additionally, certain attributes are expected to have constant values, which are specified in the dictionary_t_r.py file. The tool generates a CSV file that documents the comparison results and creates a PowerPoint presentation to present the findings.
Furthermore, the docx generator all files.py file generates a Word document with detailed results of the analysis. This tool aims to automate the process of creating informative and visually appealing presentations by utilizing the data and analysis outputs.

## Files 
This project comprises three key files:

- `dictionary_t_r.py`: This file contains constant values that are assigned to specific attributes.
- `final_check_t_r.py`: This file contains the main code responsible for comparing the attributes.
- `docs_generator_all_files.py`: This tool generates a Word file with tables presenting the statistical analysis results.

## Requirements
To run this tool, make sure you have the following packages installed:

http.client
json
ssl
requests
pandas
ast
numpy
tqdm
time
datetime
inquirer
csv
matplotlib
pptx
dataframe_image
seaborn
statistics
termcolor
docx

## Usage
1. Prepare the Excel and XML files that you want to analyze. Ensure that the necessary data is available in both files.
2. Run `final_check_t_r.py` to initiate the attribute analysis process. This code will compare the attributes and generate a PowerPoint presentation summarizing the findings. The presentation will provide insights into any discrepancies or differences identified during the analysis.
3. Execute `docs_generator_all_files.py` to generate tables with statistical results. These tables will be based on the analysis conducted in final_check_t_r.py and will offer a comprehensive overview of the attribute comparison findings.

## Contributors
Agnieszka Chilimoniuk
