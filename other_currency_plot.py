import matplotlib.pyplot as plt
import openpyxl as op
from matplotlib import ticker
from matplotlib.ticker import MultipleLocator


def draw_plot(day, count, amount, currency):
    plt.rcParams['figure.figsize'] = [9, 6.4]
    fig, ax1 = plt.subplots()
    ax2 = ax1.twinx()
    font = {'family': 'cambria',
            'size': 14}

    if currency == "MYR":
        scale = 100
        scale_2 = 10000
        y_scale = 50
        y_scale_2 = 5000
    elif currency == "KRW":
        scale = 10
        scale_2 = 10000
        y_scale = 10
        y_scale_2 = 10000
    elif currency == "INR":
        scale = 10
        scale_2 = 1000
        y_scale = 20
        y_scale_2 = 500
    elif currency == "MMK":
        scale = 10
        scale_2 = 1000
        y_scale = 5
        y_scale_2 = 2000

    ax1_plot = ax1.bar(day, amount, color='#4472C4', width=0.6)
    ax1.xaxis.set_major_locator(MultipleLocator(1))
    ax1.yaxis.set_major_locator(MultipleLocator(y_scale_2))
    ax1.set_ylim(0, (int(max(amount) / scale_2) + 1) * scale_2)
    ax1.yaxis.set_label_position('right')
    ax1.set_ylabel('thousand', fontdict=font, labelpad=10)
    ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y / 1000)))
    ax1.yaxis.set_ticks_position('right')

    ax2_plot, = ax2.plot(day, count, color='#ED7D31', linewidth=3, marker='o', markersize=6)
    ax2.xaxis.set_major_locator(MultipleLocator(1))
    ax2.yaxis.set_major_locator(MultipleLocator(y_scale))
    ax2.set_ylim(0, (int(max(count) / scale) + 1) * scale)
    ax2.yaxis.set_label_position('left')
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y / 1)))
    ax2.yaxis.set_ticks_position('left')

    plt.xlim(0.5, day[len(day) - 1] + 0.5)
    plt.title(currency, fontdict=font)
    fig.legend([ax1_plot, ax2_plot], ["Volume", "Count"], loc='lower center', ncol=2, frameon=False)
    plt.savefig("other-currency-" + currency + ".png", dpi=300, transparent=True)
    # plt.show()
    plt.clf()


def other_currency_plot():
    load_excel = op.load_workbook('summary.xlsx')
    sheet = load_excel["other market"]
    read_count_myr = sheet["B"]
    read_amount_myr = sheet["C"]
    read_count_krw = sheet["F"]
    read_amount_krw = sheet["G"]
    read_count_inr = sheet["J"]
    read_amount_inr = sheet["K"]
    read_count_mmk = sheet["N"]
    read_amount_mmk = sheet["O"]
    count_myr_array = []
    amount_myr_array = []
    count_krw_array = []
    amount_krw_array = []
    count_inr_array = []
    amount_inr_array = []
    count_mmk_array = []
    amount_mmk_array = []
    day = []
    count = 1

    for i in read_count_myr:
        if type(i.value) != str and i.value is not None:
            count_myr_array.append(i.value)
            day.append(count)
            count += 1

    for j in read_amount_myr:
        if type(j.value) != str and j.value is not None:
            amount_myr_array.append(j.value)

    for k in read_count_krw:
        if type(k.value) != str and k.value is not None:
            count_krw_array.append(k.value)

    for l in read_amount_krw:
        if type(l.value) != str and l.value is not None:
            amount_krw_array.append(l.value)

    for m in read_count_inr:
        if type(m.value) != str and m.value is not None:
            count_inr_array.append(m.value)

    for n in read_amount_inr:
        if type(n.value) != str and n.value is not None:
            amount_inr_array.append(n.value)

    for o in read_count_mmk:
        if type(o.value) != str and o.value is not None:
            count_mmk_array.append(o.value)

    for p in read_amount_mmk:
        if type(p.value) != str and p.value is not None:
            amount_mmk_array.append(p.value)

    draw_plot(day, count_myr_array, amount_myr_array, "MYR")
    draw_plot(day, count_krw_array, amount_krw_array, "KRW")
    draw_plot(day, count_inr_array, amount_inr_array, "INR")
    draw_plot(day, count_mmk_array, amount_mmk_array, "MMK")
