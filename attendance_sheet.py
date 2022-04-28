import pandas as pd
import xlsxwriter

day_time = "M 7:30-9:25"
core_course = "MATH 1920"
course_id = "ENGRG 1009-101 (9946)"
facilitator_1 = ""
facilitator_2 = ""

start_date = "2022-09-06"
end_date = "2022-12-05"

holidays = ["9/5/2022", "10/10/2022", "10/11/2022", "11/23/2022", "11/24/2022", "11/25/2022"]

def get_number(course_id):
    _, num, _ = course_id.split()
    _, number = num.split("-")
    return number

def get_day(day_time):
    day, time = day_time.split()
    if day == "M":
        return "Monday"
    elif day == "T":
        return "Tuesday"
    elif day == "W":
        return "Wednesday"
    elif day == "R":
        return "Thursday"
    else:
        return "Friday"

def format_day_time(day_time):
    day = get_day(day_time)
    _, time = day_time.split()
    return day + " " + time + "pm"

def format_workbook_name(day_time, course_id, core_course, facilitator_1, facilitator_2):
    name = get_day(day_time) + "_" + core_course + "_" + get_number(course_id)
    if facilitator_1 != "" and facilitator_2 != "":
        name = name + "_" + facilitator_1 + "&" + facilitator_2
    elif facilitator_1 != "":
        name = name + "_" + facilitator_1
    elif facilitator_2 != "":
        name = name + "_" + facilitator_2
    return name

def remove_zeroes(string):
    if string[0] == "0":
        return string[1]
    else:
        return string

def format_day(day):
    if day == "Monday":
        return "W-MON"
    elif day == "Tuesday":
        return "W-TUES"
    else:
        return "W-MON"

def get_dates(start_date, end_date, holidays):
    timestamps = pd.date_range(start_date, end_date, freq=format_day(get_day(day_time))).tolist()

    dates = []
    for i in range(len(timestamps)):
        date = str(timestamps[i]).split()[0]
        year, month, day = date.split("-")
        dates.append(remove_zeroes(month) + "/" + remove_zeroes(day) + "/" + year)
    dates = list(set(dates)-set(holidays))

    dates.append("Column1")
    dates.append("Grades")
    dates.append("Column2")

    return dates

def get_letter(number):
    return chr(ord('@')+number)

def write_dates(dates):
    current_col = 2
    for date in dates:
        current_col = current_col + 1
        worksheet.write(get_letter(current_col+1) + "8", date, header)

def write_rows():
    current_row = 10
    while current_row < 23:
        current_col = 0
        worksheet.write("B" + current_row, current_col)
        while current_col < 18:
            current_col = current_col + 1
            if(current_row%2 == 0):
                worksheet.write(get_letter(current_col+1) + str(current_row), "TEST", light_blue_background)
            else:
                worksheet.write(get_letter(current_col+1) + str(current_row), "TEST", light_gray_background)
        current_row = current_row + 1

workbook = xlsxwriter.Workbook(format_workbook_name(day_time, course_id, core_course, facilitator_1, facilitator_2) + ".xlsx")
worksheet = workbook.add_worksheet("Attendance and Grades")

bold = workbook.add_format({'bold': True})
red = workbook.add_format({'color': 'red'})
gray = workbook.add_format({'color': '#A5A5A5'})
gray_background = workbook.add_format({'bg_color': '#A5A5A5'})
blue_background = workbook.add_format({'bg_color': '#BDD6EE'})
light_gray_background = workbook.add_format({'bg_color': '#ECECEC'})
light_blue_background = workbook.add_format({'bg_color': '#DEEAF6'})

title = workbook.add_format({'bg_color': '#A5A5A5', 'bold': True})
header = workbook.add_format({'bg_color': '#BDD6EE', 'bold': True, 'align': 'center', 'text_wrap': True})

worksheet.write('B1', 'Fall 2022 Academic Excellence Workshop Attendance & Grade Sheet', title)
worksheet.set_row(0, None, gray_background)
worksheet.set_row(7, 36, blue_background)

worksheet.write('B3', 'AEW Class:', bold)
worksheet.write('B4', 'Meeting day:', bold)
worksheet.write('B5', 'Course ID:', bold)
worksheet.write('B6', 'Facilitator 1:', bold)
worksheet.write('B7', 'Facilitator 2:', bold)

worksheet.write('C3', core_course, bold)
worksheet.write('C4', format_day_time(day_time) + ";" + "Location: ", bold)
worksheet.write('C5', course_id, bold)
worksheet.write('C6', facilitator_1, bold)
worksheet.write('C7', facilitator_2, bold)

worksheet.write('B8', "Enrolled Students (Alphabetize by Last Name)", header)
worksheet.write('C8', "Learning Contract", header)

dates = get_dates(start_date, end_date, holidays)
write_dates(dates)

worksheet.set_column(1, 1, 18)
worksheet.set_row(7, 50)
worksheet.set_column(3, 1+len(dates), 12)

workbook.close()



