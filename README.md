# CS384 Attendance Record Generation Task

This project is a Python script that processes attendance data from QR code-based attendance tracking and generates a formatted Excel output. The project highlights student attendance statuses in different colors based on their attendance for each lecture session.

## Task Overview

- **Lecture Schedule**: Every Tuesday, 6:00 - 8:00 PM.
- **Attendance Tracking**: QR code scanning.
- **Lecture Format**: Two consecutive lectures, allowing students to mark attendance twice during the period.

### Input Files

- `stud_list.txt`: Contains the list of students. Students not attending due to add/drop are ignored by the script.
- `dates.txt`: Contains the date information of lecture sessions.
- `input_attendance.csv`: The raw attendance record with timestamps.

### Output File

- **output_excel.xlsx**: An Excel file with the attendance status for each student, color-coded for easier interpretation.
  - **Absent (0)**: Highlighted in Red.
  - **Partial attendance (1)**: Highlighted in Yellow.
  - **Full attendance (2)**: Highlighted in Green.
  - **Else**: No highlight.

## Features

- **Input Processing**: Parses the student list, lecture dates, and raw attendance data.
- **Attendance Calculation**: Determines attendance status based on the lecture schedule and attendance timestamps.
- **Excel Output**: Generates a neatly formatted Excel file with attendance statuses and highlights.

## Prerequisites

Ensure the following Python libraries are installed:

```bash
pip install pandas openpyxl
```

## Code Structure

The code is organized into a modular and easily understandable format. Below is an overview of the structure of the Python script:

```plaintext
📂 attendance-record-generation
 ┣ 📜 attendance_record.py        # Main script to handle the processing and generation of attendance records
 ┣ 📜 utils.py                    # Helper functions for file reading, attendance calculation, and Excel formatting
 ┣ 📜 constants.py                # Stores constants like lecture timings and column mappings for better code readability
 ┣ 📜 stud_list.txt               # Input file: Student list
 ┣ 📜 dates.txt                   # Input file: Lecture dates
 ┣ 📜 input_attendance.csv        # Input file: Attendance records
 ┗ 📜 output_excel.xlsx           # Output file: Generated Excel file with attendance
```
