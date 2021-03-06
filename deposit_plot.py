import matplotlib.pyplot as plt
import openpyxl as op
from matplotlib import ticker
from matplotlib.ticker import MultipleLocator


def deposit_plot():
    load_excel = op.load_workbook('summary.xlsx')
    sheet = load_excel["deposit"]
    read_user = sheet["B"]
    read_count = sheet["C"]
    read_amount = sheet["D"]
    user_array = []
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

    for k in read_user:
        if type(k.value) != str and k.value is not None:
            user_array.append(k.value)

    plt.rcParams['figure.figsize'] = [9, 6.4]
    fig, ax1 = plt.subplots()
    ax2 = ax1.twinx()
    font = {'family': 'cambria',
            'size': 22,
            'color': 'white'}

    ax1_plot = ax1.bar(day, amount_array, color='#4472C4', width=0.6)
    ax1.xaxis.set_major_locator(MultipleLocator(1))
    ax1.yaxis.set_major_locator(MultipleLocator(100000))
    ax1.set_ylim(0, (int(max(amount_array) / 100000) + 1) * 100000)
    ax1.yaxis.set_label_position('right')
    ax1.set_ylabel('million', fontdict=font, labelpad=10)
    ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.1f}'.format(y / 1000000)))
    ax1.xaxis.set_tick_params(labelcolor='white', labeltext=22)
    ax1.yaxis.set_tick_params(labelcolor='white')
    ax1.yaxis.set_ticks_position('right')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.spines['bottom'].set_visible(False)
    ax1.spines['left'].set_visible(False)
    ax1.xaxis.set_tick_params(top=False, bottom=False, left=False, right=False)
    ax1.yaxis.set_tick_params(top=False, bottom=False, left=False, right=False)

    ax2_plot, = ax2.plot(day, count_array, color='#ED7D31', linewidth=3, marker='o', markersize=6)
    ax2.xaxis.set_major_locator(MultipleLocator(1))
    ax2.yaxis.set_major_locator(MultipleLocator(1000))
    ax2.set_ylim(0, (int(max(count_array) / 1000) + 2) * 1000)
    ax2.yaxis.set_label_position('left')
    ax2.set_ylabel('thousand', fontdict=font, labelpad=10)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000)))
    ax2.xaxis.set_tick_params(labelcolor='white')
    ax2.yaxis.set_tick_params(labelcolor='white')
    ax2.yaxis.set_ticks_position('left')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.spines['bottom'].set_visible(False)
    ax2.spines['left'].set_visible(False)
    ax2.xaxis.set_tick_params(top=False, bottom=False, left=False, right=False)
    ax2.yaxis.set_tick_params(top=False, bottom=False, left=False, right=False)

    plt.xlim(0.5, day[len(day) - 1] + 0.5)
    plt.title('Deposit', fontdict=font)
    legend = fig.legend([ax1_plot, ax2_plot], ["Volume", "Count"], loc='lower center', ncol=2, frameon=False)
    for text in legend.get_texts():
        text.set_color("white")
    plt.savefig("deposit-volume-count.png", dpi=300, transparent=True)
    plt.show()
    plt.clf()

    # User plot start
    plt.rcParams['figure.figsize'] = [6.4, 4.8]
    plt.plot(day, user_array, color='#70AD47', linewidth=3, marker='o', markersize=6)
    plt.tick_params(top=False, bottom=False, left=False, right=False)
    ax = plt.gca()
    ax.xaxis.set_major_locator(MultipleLocator(1))
    ax.yaxis.set_major_locator(MultipleLocator(500))
    ax.set_ylim(0, (int(max(user_array) / 1000) + 1) * 1000)
    ax.xaxis.set_tick_params(labelcolor='white')
    ax.yaxis.set_tick_params(labelcolor='white')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)
    plt.ylabel('thousand', fontdict=font, labelpad=10)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y / 1000)))
    plt.xlim(0.5, day[len(day) - 1] + 0.5)
    plt.title('User', fontdict=font)
    plt.savefig("deposit-user.png", dpi=300, transparent=True)
    # plt.show()
    plt.clf()

deposit_plot()