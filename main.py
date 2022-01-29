import os

from create_new_workbook import create_new_workbook
from deposit_merge import deposit_merge
from deposit_plot import deposit_plot
from deposit_sum import deposit_sum
from withdraw_merge import withdraw_merge
from withdraw_plot import withdraw_plot
from withdraw_sum import withdraw_sum
from main_deposit import main_deposit
from main_deposit_plot import main_deposit_plot
from other_currency import other_currency
from other_currency_plot import other_currency_plot

if not os.path.isfile('./summary.xlsx'):
    create_new_workbook()

deposit_merge()
print("please insert year: ", end="")
year = input()
print("please insert month: ", end="")
month = input()
deposit_sum(year, month)
deposit_plot()
withdraw_merge()
withdraw_sum(year, month)
withdraw_plot()
main_deposit()
main_deposit_plot()
other_currency()
other_currency_plot()
