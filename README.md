# CS384 Attendance Record Generation Task

This project is designed to process attendance data for students and generate a formatted Excel report. The attendance data is collected through QR code scanning, and the system automatically processes the records based on the input CSV files. The project includes the following functionalities:

- Splitting roll number and student name from the input data.
- Processing student attendance records based on time and lecture dates.
- Generating a color-coded Excel report indicating attendance status (Absent, Partial, Full).

## Table of Contents

- [Features](#features)
- [Libraries Used](#libraries-used)
- [File Structure](#file-structure)

## Features

- Splits the `Roll` column into two separate columns: `Roll Number` and `Name`.
- Processes attendance based on lecture times, using predefined time slots.
- Color-coded Excel output:
  - **Absent (0)**: Highlighted in Red.
  - **Partial attendance (1)**: Highlighted in Yellow.
  - **Full attendance (2)**: Highlighted in Green.
  - **Else**: No highlight.

## Libraries Used

### 1. pandas

- **Description**: `pandas` is a data manipulation and analysis library for Python that provides DataFrames for handling structured data and supports reading from and writing to formats like CSV and Excel.

- **Installation**:
  ```bash
  pip install pandas
  ```

### 2. openpyxl

- **Description**: `openpyxl` is a Python library for reading and writing Excel files, allowing you to create, modify, and format spreadsheets with features like styling, charts, and conditional formatting.

- **Installation**:
  ```bash
  pip install openpyxl
  ```

### 3. datetime (built-in library)

- **Description**: The `datetime` module provides classes for manipulating dates and times, including arithmetic and formatting functions. It is part of the Python Standard Library and requires no separate installation.

- **Installation**: No installation required. This module is included with Python.

## File Structure

The project contains the following key files:

```bash
├── input_attendance.csv            # Original attendance data with roll number and name combined.
├── input_attendance_processed.csv  # Processed attendance data with roll number and name split.
├── stud_list.txt                   # List of students in the format 'RollNumber Name'.
├── dates.txt                       # Comma-separated lecture dates (e.g., 06/08/2024, 13/08/2024).
├── output_excel.xlsx               # Final Excel report with color-coded attendance status.
├── app.py                          # Main Python script for processing attendance.
├── README.md                       # This README file.

```
