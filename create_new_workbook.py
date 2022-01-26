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

    sheet_d = createNewExcel.create_sheet("by currency withdraw")
    # sheet_d.title = 'by currency withdraw'
    sheet_d['A1'] = 'DayOfMonth'
    sheet_d['B1'] = 'Currency'
    sheet_d['C1'] = 'Users'
    sheet_d['D1'] = 'Count'
    sheet_d['E1'] = 'Volume'
    sheet_d.column_dimensions['A'].width = 20.0
    sheet_d.column_dimensions['B'].width = 20.0
    sheet_d.column_dimensions['C'].width = 20.0
    sheet_d.column_dimensions['D'].width = 20.0
    sheet_d.column_dimensions['E'].width = 20.0

    createNewExcel.save('summary.xlsx')
