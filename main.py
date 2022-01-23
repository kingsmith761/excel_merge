import os
from create_new_workbook import create_new_workbook
from deposit_count import deposit_count
from deposit_merge import deposit_merge
from deposit_sum import deposit_sum

if not os.path.isfile('./summary.xlsx'):
    create_new_workbook()

deposit_merge()
deposit_sum()
# deposit_count()
