import openpyxl as op
import matplotlib.pyplot as plt


def deposit_plot():
    load_excel = op.load_workbook('summary.xlsx')
    sheet = load_excel["deposit"]
    read_count = sheet["C"]
    read_amount = sheet["D"]
    count_array = []
    amount_array = []
    day = []
    count = 1

    for i in read_count:
        if type(i.value) != str and i.value is not None:
            count_array.append(i.value)
            day.append(count)
            count += 1

    for j in read_amount:
        if type(j.value) != str and j.value is not None:
            amount_array.append(j.value)

    fig, ax1 = plt.subplots()
    ax2 = ax1.twinx()

    ax1.plot(day, count_array, color='tab:orange')
    ax2.bar(day, amount_array, color='blue', alpha=0.7)
    plt.show()


deposit_plot()
