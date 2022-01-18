import openpyxl as op


def deposit_sum():
    load_excel = op.load_workbook('summary.xlsx')
    data_sheet = load_excel["d_cnt.amnt"]
    if "d" in load_excel.sheetnames:
        sheet = load_excel["d"]
    else:
        sheet = load_excel.create_sheet("d")
        sheet['A1'] = 'DayOfMonth'
        sheet['B1'] = 'DAU'
        sheet['C1'] = 'Count'
        sheet['D1'] = 'Amount'
        sheet['H14'] = 'Monthly AVG'
        sheet['H17'] = 'Total'
        sheet['H18'] = 'W1'
        sheet['H19'] = 'W2'
        sheet['H20'] = 'W3'
        sheet['H21'] = 'W4'
        sheet['H23'] = 'AVG'
        sheet['H24'] = 'W1'
        sheet['H25'] = 'W2'
        sheet['H26'] = 'W3'
        sheet['H27'] = 'W4'
        sheet['I13'] = 'A'
        sheet['I14'] = '=AVERAGE(D2:D32)'
        sheet['I22'] = '=(I21-I19)/I19'
        sheet['I28'] = '=(I27-I25)/I25'
        sheet['J13'] = 'C'
        sheet['J14'] = '=AVERAGE(C2:C32)'
        sheet['J22'] = '=(J21-J19)/J19'
        sheet['J28'] = '=(J27-I25)/J25'
        sheet['K13'] = 'DAU'
        sheet['K14'] = '=AVERAGE(B2:B32)'
        sheet['K22'] = '=(K21-K19)/K19'
        sheet['K28'] = '=(K27-K25)/K25'
    start = 0
    count = 0
    amount = 0

    for i in range(2, data_sheet.max_row + 2):
        if start == 0:
            start = int(data_sheet["A" + str(i)].value)
        if data_sheet["A" + str(i)].value is None:
            sheet["A" + str(start + 1)] = str(start)
            sheet["C" + str(start + 1)] = str(count)
            sheet["D" + str(start + 1)] = str(amount)
            break
        elif int(data_sheet["A" + str(i)].value) == start:
            if data_sheet["C" + str(i)].value is not None:
                count += int(data_sheet["C" + str(i)].value)
            if data_sheet["D" + str(i)].value is not None:
                amount += int(data_sheet["D" + str(i)].value)
        elif int(data_sheet["A" + str(i)].value) > start:
            sheet["A" + str(start + 1)] = str(start)
            sheet["C" + str(start + 1)] = str(count)
            sheet["D" + str(start + 1)] = str(amount)
            count = 0
            amount = 0
            start = int(data_sheet["A" + str(i)].value)

    load_excel.save('summary.xlsx')
