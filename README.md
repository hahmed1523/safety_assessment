# Automate Safety Assessment Report

## Description
The Safety Assessment Report is manually run on a quarterly basis. The report requires various excel manipulation steps. To automate the steps I used the Pandas, Openpyxl, Xlsxwriter, and Pyodbc libraries.

## Automation Tasks
- Connect to the Microsoft Access database.
- Ask user to enter the date range for the report.
- Use the date range provided to create a SQL statement to query MS Access.
- Query MS Access and put the data into a Pandas Dataframe.
- Change all NaN values to 'Blank'.
- Create a new Dataframe with only the columns that have questions in it and put them in the correct order.
- Change any comments in the data into "Other".
- Create a list of Yes, No, N/A, Blank, and Other values per question.
- Create a Dataframe from the lists.
- Add columns to get the percentage for Yes, No, etc.
- Create a pivot tables for the Safety Assessment Questions.
- If the question does not have a certain value then use 0 for the value. This is to keep all the tables have the same number of columns and rows.
- Create a total pivot table at the end to sum all values in SA pivot tables.
- Export all tables and data into Excel with proper formatting.
- Adjust column width dynamically for the Raw Data tab.

##See example folder to test the code.
