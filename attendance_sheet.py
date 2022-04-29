import pandas as pd
import xlsxwriter

day_time = "M 7:30-9:25"
core_course = "MATH 2940"
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
        worksheet_at.set_column(current_col-1, current_col-1, 8)
        current_col = current_col + 1
        worksheet_at.write(get_letter(current_col+1) + "8", date, header)

def write_dates_space(dates):
    current_col = 2
    for date in dates:
        worksheet_class.set_column(current_col-1, current_col-1, 8)
        current_col = current_col + 1
        worksheet_class.write(get_letter(current_col+1) + "8", date, header)

def write_rows():
    current_row = 9
    while current_row < 32:
        current_col = 0
        if(current_row != 9):
            worksheet_at.write("A" + str(current_row), current_row-9, border)
        while current_col < 17:
            current_col = current_col + 1
            if(current_row%2 == 0):
                worksheet_at.write(get_letter(current_col+1) + str(current_row), "", light_blue_background)
            else:
                worksheet_at.write(get_letter(current_col+1) + str(current_row), "", light_gray_background)
        current_row = current_row + 1

workbook = xlsxwriter.Workbook(format_workbook_name(day_time, course_id, core_course, facilitator_1, facilitator_2) + ".xlsx")
worksheet_at = workbook.add_worksheet("Attendance and Grades")
worksheet_class = workbook.add_worksheet("Class List")

bold = workbook.add_format({'bold': True})
bold_border = workbook.add_format({'bold': True})
side_border = workbook.add_format({'bold': True})
left_border = workbook.add_format({})
red = workbook.add_format({'color': 'red', 'bold': True})
gray = workbook.add_format({'color': '#A5A5A5', 'bold': True})
gray_background = workbook.add_format({'bg_color': '#A5A5A5'})
blue_background = workbook.add_format({'bg_color': '#BDD6EE'})
light_gray_background = workbook.add_format({'bg_color': '#ECECEC'})
light_blue_background = workbook.add_format({'bg_color': '#DEEAF6'})
light_gray_background.set_border()
light_blue_background.set_border()
bold_border.set_border()
side_border.set_right()
side_border.set_bottom()
side_border.set_top()
left_border.set_left()
left_border.set_bottom()
left_border.set_top()

light_gray_background_bold = workbook.add_format({'bg_color': '#ECECEC', 'bold': True})
light_gray_background_red = workbook.add_format({'bg_color': '#ECECEC', 'bold': True, 'color': 'red'})
light_gray_background_gray = workbook.add_format({'bg_color': '#ECECEC', 'bold': True, 'color': '#A5A5A5'})

title = workbook.add_format({'bg_color': '#A5A5A5', 'bold': True})
header = workbook.add_format({'bg_color': '#BDD6EE', 'bold': True, 'align': 'center', 'text_wrap': True, 'size': 9})
header.set_border()
red.set_border()
gray.set_border()
light_gray_background_red.set_border()
light_gray_background_gray.set_border()

border = workbook.add_format({})
border.set_border()

worksheet_at.write('B1', 'Fall 2022 Academic Excellence Workshop Attendance & Grade Sheet', title)
worksheet_at.set_row(0, None, gray_background)
worksheet_at.set_row(7, 36, blue_background)

worksheet_at.write('B3', 'AEW Class:', bold)
worksheet_at.write('B4', 'Meeting day:', bold)
worksheet_at.write('B5', 'Course ID:', bold)
worksheet_at.write('B6', 'Facilitator 1:', bold)
worksheet_at.write('B7', 'Facilitator 2:', bold)

worksheet_at.write('C3', core_course, bold)
worksheet_at.write('C4', format_day_time(day_time) + ";" + "Location: ", bold)
worksheet_at.write('C5', course_id, bold)
worksheet_at.write('C6', facilitator_1, bold)
worksheet_at.write('C7', facilitator_2, bold)

worksheet_at.write('O3:Q3', 'Attendance Symbols*   ', bold_border)
worksheet_at.write('O4', 'A', red)
worksheet_at.write('O5', 'T', gray)
worksheet_at.write('O6', 'X', border)

worksheet_at.write('P4:Q4', 'Absent    ', left_border)
worksheet_at.write('P5:Q5', 'Tardy    ', left_border)
worksheet_at.write('P6:Q6', 'Present    ', left_border)

worksheet_at.write('A8', "", header)
worksheet_at.write('B8', "Enrolled Students (Alphabetize by Last Name)", header)
worksheet_at.write('C8', "Learning Contract", header)

worksheet_at.write('B32:E32', 'I verify that these grades are complete and accurate. ', bold)
worksheet_at.write('O32:Q32', '*See Attendance Policies Sheet for details', bold)
worksheet_at.write('I33', 'Signature')
worksheet_at.write('L33', 'Date')
worksheet_at.write('F32:I32', '__________________________________________________', bold)
worksheet_at.write('K32:L32', '_________________________', bold)

worksheet_at.write('B34:E34', 'I verify that these grades are complete and accurate. ', bold)
worksheet_at.write('I35', 'Signature')
worksheet_at.write('L35', 'Date')
worksheet_at.write('F34:I34', '')
worksheet_at.write('K34:L34', '')
worksheet_at.write('F34:I34', '__________________________________________________', bold)
worksheet_at.write('K34:L34', '_________________________', bold)

dates = get_dates(start_date, end_date, holidays)
write_dates(dates)

worksheet_at.set_column(1, 1, 22)
worksheet_at.set_row(7, 50)
worksheet_at.set_row(31, 60)
worksheet_at.set_row(33, 60)
worksheet_at.set_row(1, 3)

write_rows()

worksheet_at.write('A9', "", light_gray_background)
worksheet_at.write('B9', 'Students, Ima (example)', light_gray_background_bold)
worksheet_at.write('D9', 'A', light_gray_background_red)
worksheet_at.write('E9', 'T', light_gray_background_gray)
worksheet_at.write('F9', 'X', light_gray_background)
worksheet_at.write('P3', "", side_border)
worksheet_at.write('Q3', "", side_border)
worksheet_at.write('Q4', "", side_border)
worksheet_at.write('Q5', "", side_border)
worksheet_at.write('Q6', "", side_border)

workbook.close()



