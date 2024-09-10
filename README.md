# Automated Data Processing and Conditional Formatting for Alumnae Relations
## Project Overview
This project focuses on automating the processing, cleaning, and formatting of large datasets containing alumnae contact information. The goal was to streamline data management tasks by developing Python scripts that automate repetitive tasks, apply conditional formatting, and ensure consistent data structures across multiple Excel files.
## Key Features
- **Data Transformation and Cleaning:** 
  - Automated deletion of unnecessary columns and standardization of contact information.
  - Dynamic handling of varying file structures with error-handling for missing data.

- **Batch Processing:** 
  - The script processes multiple Excel files in a folder, applying bulk data transformations and reducing manual intervention.

- **Sorting and Formatting:**
  - Automated sorting of data based on the "Last" column in ascending order for both sheets.
  - Freezing of header rows in all sheets to improve usability.

- **Conditional Formatting:**
  - Conditional formatting rules applied to highlight critical data, such as "Inactive", "Invalid Address", and "Deceased?" columns, where cells containing "Yes" are marked in red.
## Technologies Used
- **Python**
	- Pandas for data manipulation.
	- Openpyxl for Excel processing, formatting, and conditional formatting.
## How It Works
1. Data Processing:
   - The script loads each Excel file and processes the "Contact List" and "Class Directory" sheets.
   - Unnecessary columns are deleted, and data is sorted based on the "Last" column.
2. Conditional Formatting:
	- Conditional formatting is applied to highlight important data points such as inactive statuses and invalid addresses.
3. Batch Automation:
	- The script processes all Excel files in a folder, ensuring consistent formatting and cleaning across the entire dataset.

## Project Impact
This project improves the efficiency of alumnae data management by automating repetitive tasks, reducing manual effort, and ensuring data consistency across large datasets.
