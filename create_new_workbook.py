import openpyxl


def create_new_workbook():
    createNewExcel = openpyxl.Workbook()
    sheet = createNewExcel.worksheets[0]
    sheet.title = 'd_cnt.amnt'
    sheet['A1'] = 'DayOfMonth'
    sheet['B1'] = 'Currency'
    sheet['C1'] = 'Count'
    sheet['D1'] = 'Amount'
    sheet['E1'] = 'Users'

    createNewExcel.save('summary.xlsx')
