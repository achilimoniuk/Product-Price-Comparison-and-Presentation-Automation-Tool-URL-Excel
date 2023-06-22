# Product Price Comparison and Presentation Automation Tool

## Description
The purpose of this code is to compare the prices of products extracted from an Excel file with those obtained from a URL. The code aims to identify discrepancies in prices for each product. If any discrepancies are found, the tool further investigates the issue based on additional attributes information.
The tool generates tables in a Microsoft Word document (docx) summarizing the comparison results and creates a PowerPoint presentation to present the findings. It aims to automate the process of creating informative and visually appealing presentations using the data and analysis outputs.


## Files
- `step1-compare_url_excel.py`: This tool extracts product information from a given URL and saves it as a CSV file. The extracted information includes details such as product name, price and any other relevant attributes available on the webpage.
- `step2_issue_classification.py`: This tool compares prices obtained from a URL and an Excel file, and identifies the types of mismatches that occur between the two sources. 
- `step3-generate table.py`: This tool generates tables in a Microsoft Word document based on the obtained results.
- `step4-creating_presentation.py`: This tool generates a PowerPoint presentation with statistics and plots based on the obtained results. 

## Requirements
To run this tool, make sure you have the following packages installed:
- http.client
- json
- ssl
- requests
- pandas
- ast
- numpy
- tqdm 
- time 
- datetime
- inquirer
- csv 
- matplotlib
- pptx 
- dataframe_image 
- seaborn 
- statistics
- termcolor 
- docx

## Usage
1. Prepare the URL and Excel file that is needed to be analyzed.
2. Run `step1-compare_url_excel.py`: Extract product information from the specified URL.Save the extracted information, including product name, price, and relevant attributes, as a CSV file.
3. Run `step2_issue_classification.py`:Compare the prices obtained from the URL with those from the Excel file.Identify any discrepancies or mismatches between the two sources.
4. Run `step3-generate_table.py`: Use the obtained results to generate tables in a Microsoft Word document (docx).
5. Run `step4-creating_presentation.py`: Utilize the obtained results to create a PowerPoint presentation.

## Contributors
Agnieszka Chilimoniuk
