import pandas as pd

semester = 6
course = 6 + semester // 2
def read_and_save_to_dict(file_path, column_indices=None, row_start=None, row_end=None):
    data_dict = {}

    # Use ExcelFile to handle multiple sheets
    xls = pd.ExcelFile(file_path)

    # Iterate through each sheet
    for sheet_name in xls.sheet_names[6:course]:
        print(sheet_name)

        # Read the sheet into a DataFrame
        df = pd.read_excel(xls, sheet_name)

        # If column_indices is not provided, use all the columns
        if column_indices is None:
            column_indices = range(len(df.columns))

        # Validate column indices
        if any(idx >= len(df.columns) for idx in column_indices):
            raise ValueError("Invalid column index provided.")

        # Iterate over rows within the specified range and save it to the dictionary
        for index, row in df.iterrows():
            if row_start is not None and index < row_start:
                continue # skip rows before the specified start index

            if row_end is not None and index > row_end:
                break

            for i in range(1, len(column_indices), 2):
                key = row.iloc[column_indices[0]]
                values = list(row.iloc[i] for i in column_indices[i:i+2])  # Use the rest as values

                if key not in data_dict:
                    data_dict[key] = values
                else:
                    if values[0] != 0:
                        data_dict[key][0] += values[0]
                        data_dict[key][1] = values[1]



    print(data_dict)
    print(len(data_dict))
    return data_dict

# Example usage:
file_path = 'C:\\Users\\Пользователь\\Downloads\\Telegram Desktop\\ex.xlsx'
start = 5
end = 30
selected_columns = [2, 11, 12, 21, 22]

result_dict = read_and_save_to_dict(file_path, column_indices=selected_columns, row_start=start, row_end=end)

# Print the resulting dictionary
#print(result_dict)

# self.set_cell_format(A, '№')
#         self.set_cell_format(A1, 'Название дисциплины из текущего РУПа /Название ранее сданной дисциплины')
#         self.set_cell_format(A2, 'Объем текущего РУПа')
#         self.set_cell_format(A3, 'З / Э/ ЗО')
#         self.set_cell_format(A4, 'Семестр')
#         self.set_cell_format(A5, 'Отметка о сдаче')
#         self.set_cell_format(hdr_cells[2], 'Зач.ед')
#         self.set_cell_format(hdr_cells[3], 'Часы')
#         self.set_cell_format(hdr_cells[6], 'оценка')
#         self.set_cell_format(hdr_cells[7], 'дата')
