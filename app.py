import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill

def preprocess(input_file, output_file):
    
    """
    Logic to split Roll column in original csv file into roll no and Name column for easier processing.
    Generated input_data_processed.csv from original csv file.
    
    """
    
    df = pd.read_csv(input_file)
    df[['Roll Number', 'Name']] = df['Roll'].str.split(' ', n=1, expand=True)
    df = df.drop(columns=['Roll'])                       
    df.to_csv(output_file, index=False)
    
    

def getStudentData(inpFile):
    
    """
    Returns:
        _type_: Dictionary containing roll no ---> name mapping.
        
    """
    data = {}
    with open(inpFile, 'r') as f:
        for line in f:
            roll, name = line.strip().split(' ',1) 
            data[roll] = name 
    return data



def preprocessAttendanceData(file):
    """
    Convert Timestamp data in input attendance file so that it can be processed and manipulated by Pandas.
    
    Args:
        file : Input attendance file.

    Returns:
        pandas Dataframe object containing info about attendance.
    
    Datetime format: '%d-%m-%Y %H:%M'
    
    """ 
    data = pd.read_csv(file)
    data['Timestamp'] = pd.to_datetime(data['Timestamp'], format='%d-%m-%Y %H:%M')
    return data



def attendanceSummary(attendance_data, dates, students): 
    
    """
    Processes attendance data for each student and returns a summary of attendance status.

    For each student and class date:
    - Full attendance (2): if the student has two valid attendance records for the class.
    - Partial attendance (1): if the student has one valid attendance record for the class.
    - More than full attendance: if there are more than 2 attendance records for the class, 
      the actual count is recorded.
    - Absent (0): if no attendance record exists for the class.
    
    Args:
        attendance_data : Dataframe of attendance data
        dates : list containing data when the classes were taken.
        students : Dictionary of students.
    
    Returns:
        summary : Dictionary containing student mapped to attendance data on a given date.
        
    """
    
    summary = {}
    for roll in students.keys():
        summary[roll] = {}
        for date in dates:
            summary[roll][date] = 0

    for roll in students.keys(): 
        for date in dates:
            
            date_formatted = datetime.strptime(date, '%d/%m/%Y').strftime('%Y-%m-%d')
            lecture_time_start = datetime.strptime(f"{date_formatted} 18:00", '%Y-%m-%d %H:%M')
            lecture_time_end = datetime.strptime(f"{date_formatted} 20:00", '%Y-%m-%d %H:%M')
            
            student_attendance = attendance_data[
                (attendance_data['Roll Number'] == roll) & 
                (attendance_data['Timestamp'] >= lecture_time_start) & 
                (attendance_data['Timestamp'] <= lecture_time_end)
            ]

            if len(student_attendance) == 2:
                summary[roll][date] = 2  
            elif len(student_attendance) == 1:
                summary[roll][date] = 1  
            elif len(student_attendance) > 2:
                summary[roll][date] = len(student_attendance)
            else:
                summary[roll][date] = 0  

    return summary



def totalAttendanceCount(input_file):
    """

    Args:
        input_file : Input attendance processed.

    Returns:
        total_attendance : roll no -- > Total Attendance.
         
    """
    
    df = pd.read_csv(input_file)    
    total_attendance = df.groupby('Roll Number').size()
    return total_attendance

def writeExcel(summary, students, dates, total_attendance, total_classes_taken, output_file):
    
    """
    
    Generates an Excel file with attendance summary for students. 
    The output includes total attendance marked, sum of attendance, total classes allowed, and proxy attendance.

    Args:
        summary (dict): A dictionary mapping student roll numbers to their attendance status on given dates.
        students (dict): A dictionary mapping student roll numbers to their names.
        dates (list): A list of dates when classes were held.
        total_attendance (dict): A dictionary mapping student roll numbers to their total attendance marked.
        total_classes_taken (int): The total number of classes taken.
        output_file (str): The path where the Excel file will be saved.
    
    """

    df = pd.DataFrame(summary).T
    df.index = [f"{roll} {students[roll]}" for roll in df.index]  
    df.columns = dates

    df['Total Attendance Marked'] = df.index.map(lambda x: total_attendance.get(x.split()[0], 0))
    df['Sum of Attendance'] = df[dates].sum(axis=1)
    df['Total Attendance Allowed'] = total_classes_taken
    df['Proxy'] = (df['Total Attendance Marked'] - df['Sum of Attendance']).abs()

    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    df.to_excel(writer, sheet_name='Attendance', index=True)

    worksheet = writer.sheets['Attendance']

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    for row in worksheet.iter_rows(min_row=2, min_col=2, max_row=len(df)+1, max_col=len(dates)+1):
        for cell in row:
            if cell.value > 2:
                cell.fill = red_fill
            elif cell.value == 1:
                cell.fill = yellow_fill
            elif cell.value == 2:
                cell.fill = green_fill

    writer.close()

def main():
    preprocess('input_attendance.csv','input_attendance_processed.csv')
    students = getStudentData('stud_list.txt')
    
    dates = []
    with open('dates.txt', 'r') as f:
        dates = f.read().strip().split(', ')
        
    total_classes_taken = 2 * len(dates)
    attendance_data = preprocessAttendanceData('input_attendance_processed.csv')
    summary = attendanceSummary(attendance_data, dates, students)
    total_attendance = totalAttendanceCount('input_attendance_processed.csv')
    writeExcel(summary, students, dates, total_attendance, total_classes_taken, 'output_excel.xlsx')
