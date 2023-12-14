import docx
import os

# open the word document
doc = docx.Document('C:\\Users\\Польз   ователь\\PycharmProjects\\NIR\\certificate_with_marks.docx')

# semester I choose from the app
semester = 2

# list where I save all the data from each semester
all_subjects = []
cnt_s = 0

for table in doc.tables:
    sem_subjects = {}
    # extract data from cells in the first row
    first_row = table.rows[0].cells[0]
    first_row_data = first_row.text
    sem = str(cnt_s + 1) + " " + "семестр"

    print(first_row_data)



    # checking if semester that we need in the table
    if sem in first_row_data:
        cnt_s += 1
        # find the column indices
        header_row = table.rows[1]
        date_column_index = None
        marks_column_index = None

        for i, cell in enumerate(header_row.cells):
            if cell.text.strip().lower() == 'дата сдачи':
                date_column_index = i
            elif cell.text.strip().lower() == 'оценка (ects)':
                marks_column_index = i


        for row in table.rows[2:]:

            discipline = row.cells[0].text.strip()
            if "зачёт" in discipline:
                control = "З"
                discipline1 = discipline.replace("(зачёт)", '')
            elif "экзамен" in discipline:
                control = "Э"
                discipline1 = discipline.replace("(экзамен)", '')
            elif "аттестация всех разделов" in discipline:
                control = "аттестовано"
                discipline1 = discipline.replace("(аттестация всех разделов)", '')
            else:
                control = ''
                discipline1 = discipline

            # getting the date and marks using identified column indices
            date = row.cells[date_column_index].text.strip() if date_column_index is not None else ''
            marks = row.cells[marks_column_index].text.strip() if marks_column_index is not None else ''

            # create dictionary
            sem_subjects[discipline1] = (control, date, marks)

    if sem_subjects:
        all_subjects.append(sem_subjects)
    if cnt_s == semester:
        break




print(all_subjects)


