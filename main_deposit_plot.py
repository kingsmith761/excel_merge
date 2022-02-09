import matplotlib.pyplot as plt
import openpyxl as op
from matplotlib import ticker
from matplotlib.ticker import MultipleLocator


def main_deposit_plot():
    load_excel = op.load_workbook('summary.xlsx')
    sheet = load_excel["main"]
    read_count_idr = sheet["B"]
    read_count_THB = sheet["C"]
    read_count_VND = sheet["D"]
    count_idr_array = []
    count_thb_array = []
    count_vnd_array = []
    read_amount_idr = sheet["H"]
    read_amount_THB = sheet["I"]
    read_amount_VND = sheet["J"]
    amount_idr_array = []
    amount_thb_array = []
    amount_vnd_array = []
    day = []
    count = 1

    for i in read_count_idr:
        if type(i.value) != str and i.value is not None:
            count_idr_array.append(i.value)
            day.append(count)
            count += 1

    for j in read_count_THB:
        if type(j.value) != str and j.value is not None:
            count_thb_array.append(j.value)

    for k in read_count_VND:
        if type(k.value) != str and k.value is not None:
            count_vnd_array.append(k.value)

    for i in read_amount_idr:
        if type(i.value) != str and i.value is not None:
            amount_idr_array.append(i.value)

    for j in read_amount_THB:
        if type(j.value) != str and j.value is not None:
            amount_thb_array.append(j.value)

    for k in read_amount_VND:
        if type(k.value) != str and k.value is not None:
            amount_vnd_array.append(k.value)

    font = {'family': 'cambria',
            'size': 16,
            'color': 'white'}

    plt.rcParams['figure.figsize'] = [9, 4.8]
    count_idr, = plt.plot(day, count_idr_array, color='#A5A5A5', linewidth=3, marker='o', markersize=6)
    count_thb, = plt.plot(day, count_thb_array, color='#ED7D31', linewidth=3, marker='o', markersize=6)
    count_vnd, = plt.plot(day, count_vnd_array, color='#00B050', linewidth=3, marker='o', markersize=6)
    ax = plt.gca()
    ax.xaxis.set_major_locator(MultipleLocator(1))
    ax.yaxis.set_major_locator(MultipleLocator(500))
    count_max = max(max(count_idr_array), max(count_thb_array), max(count_vnd_array))
    ax.set_ylim(0, (int(count_max / 1000) + 1) * 1000)
    ax.yaxis.set_label_position('right')
    ax.xaxis.set_tick_params(labelcolor='white')
    ax.yaxis.set_tick_params(labelcolor='white')
    plt.ylabel('thousand', fontdict=font, labelpad=10)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.1f}'.format(y / 1000)))
    plt.xlim(0.5, day[len(day) - 1] + 0.5)
    plt.title('Main market transaction count', fontdict=font)
    legend = ax.legend([count_idr, count_thb, count_vnd],
              ["IDR", "THB", "VND"],
              loc='upper center',
              ncol=3,
              frameon=False,
              bbox_to_anchor=(0.5, -0.06))
    for text in legend.get_texts():
        text.set_color("white")
    plt.savefig("main-deposit-count.png", dpi=300, transparent=True)
    # plt.show()
    plt.clf()

    amount_idr, = plt.plot(day, amount_idr_array, color='#A5A5A5', linewidth=3, marker='o', markersize=6)
    amount_thb, = plt.plot(day, amount_thb_array, color='#ED7D31', linewidth=3, marker='o', markersize=6)
    amount_vnd, = plt.plot(day, amount_vnd_array, color='#00B050', linewidth=3, marker='o', markersize=6)
    ax = plt.gca()
    ax.xaxis.set_major_locator(MultipleLocator(1))
    ax.yaxis.set_major_locator(MultipleLocator(50000))
    amount_max = max(max(amount_idr_array), max(amount_thb_array), max(amount_vnd_array))
    ax.set_ylim(0, (((int(amount_max)) / 100000) + 1) * 100000)
    ax.yaxis.set_label_position('right')
    ax.xaxis.set_tick_params(labelcolor='white')
    ax.yaxis.set_tick_params(labelcolor='white')
    plt.ylabel('thousand', fontdict=font, labelpad=10)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.2f}'.format(y / 1000000)))
    plt.xlim(0.5, day[len(day) - 1] + 0.5)
    plt.title('Main market volume', fontdict=font)
    legend = ax.legend([amount_idr, amount_thb, amount_vnd],
              ["IDR", "THB", "VND"],
              loc='upper center',
              ncol=3,
              frameon=False,
              bbox_to_anchor=(0.5, -0.06))
    for text in legend.get_texts():
        text.set_color("white")
    plt.savefig("main-deposit-amount.png", dpi=300, transparent=True)
    # plt.show()
    plt.clf()
