import openpyxl


def create_new_workbook():
    createNewExcel = openpyxl.Workbook()
    sheet = createNewExcel.worksheets[0]
    sheet.title = 'by currency'
    sheet['A1'] = 'DayOfMonth'
    sheet['B1'] = 'Currency'
    sheet['C1'] = 'Users'
    sheet['D1'] = 'Count'
    sheet['E1'] = 'Volume'
    sheet.column_dimensions['A'].width = 20.0
    sheet.column_dimensions['B'].width = 20.0
    sheet.column_dimensions['C'].width = 20.0
    sheet.column_dimensions['D'].width = 20.0
    sheet.column_dimensions['E'].width = 20.0

    createNewExcel.save('summary.xlsx')
