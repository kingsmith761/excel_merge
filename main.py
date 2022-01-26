import os
from create_new_workbook import create_new_workbook
from deposit_count import deposit_count
from deposit_merge import deposit_merge
from deposit_sum import deposit_sum
from withdraw_merge import withdraw_merge
from withdraw_sum import withdraw_sum

if not os.path.isfile('./summary.xlsx'):
    create_new_workbook()

deposit_merge()
deposit_sum()
# deposit_count()
# withdraw_merge()
# withdraw_sum()
