import xlrd
import os

path = "/home/rozengold/PycharmProjects/Arduino/FIRST YEAR"
row_data = []


def slice_sheet(single_sheet, rows=0, columns=0):
    if type(single_sheet) == xlrd.sheet.Sheet:
        try:
            for row_index in range(int(rows - 3)):
                if row_index == 0 or row_index == 1: continue
                # if columns == 7:
                for column_index in range(columns):
                    if (column_index == 3 and columns == 6) or ((column_index == 0 and columns == 7) or \
                             (column_index == 4 and columns == 7)): continue
                    # print('i ame here at row_index {0} column_index {1}'.format(row_index, column_index))
                    try:
                        if column_index == 0 and columns == 6: row_data.append(int(single_sheet.cell_value(row_index, column_index)))
                        elif column_index == 1 and columns == 7: row_data.append(int(single_sheet.cell_value(row_index, column_index)))
                        else : row_data.append("\"" + str(single_sheet.cell_value(row_index, column_index)).upper().replace(' ', '-') + "\"")
                    except ValueError:
                        break
                print(row_data, end='\n')
                write_data_to_file(row_data)
                row_data.clear()
        except IndexError:
            pass
    return None


def write_data_to_file(row):
    with open('temp.txt', 'a') as temp:
        for item in row[1:]:
            temp.write(str(item) + '\t')
        temp.write(str('\n'))
        temp.close()


def main():
    for excel_sheet_file_name in os.listdir(path):
        if excel_sheet_file_name.endswith('.xls') and True:  # check it is file
            workbook = xlrd.open_workbook(path + '/' + excel_sheet_file_name)
            # print(path + '/' + excel_sheet_file_name)
            try:
                for index in range(workbook.nsheets):
                    # print(workbook.nsheets , index)
                    sheet = workbook.sheet_by_index(index)
                    slice_sheet(sheet, sheet.nrows, sheet.ncols)
                # print(sheet.nrows, " ", sheet.ncols)
            except IndexError:
                pass


main()
