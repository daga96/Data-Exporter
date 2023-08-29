# Excel to SQL Converter

This program is a simple application built using JavaScript and Electron that enables you to convert Excel files in a specific format into SQL files.

## Features

- **Excel Format:** The program requires the input Excel file to adhere to a predetermined format.

- **SQL File Generation:** After selecting an Excel file, the program reads data from each sheet and creates SQL statements for table creation and data insertion.

## Excel Format

The input Excel file should follow a structured format:

- Each sheet within the Excel file represents a table.
- The first row of each sheet contains column headers.

Ensure that your Excel files are formatted correctly to ensure accurate conversion.

## SQL File Generation

Once the program receives your selected Excel file:

1. It reads the content of each sheet, where each sheet represents a table.
2. For each sheet, it generates SQL statements:
   - **Table Creation:** Creates the necessary SQL statement to generate the table structure based on the column headers.
   - **Data Insertion:** Generates SQL INSERT statements for each row of data in the sheet.

The resulting SQL file will contain both the table creation and data insertion statements.
