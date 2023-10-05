import numpy as np
import pandas as pd

print('reading bidder matrix ...')
# bidder matrix
df_tender_matrix = pd.read_excel('tender-matrix.xlsx', sheet_name='tender-matrix', header=0, index_col=0)
lot_tender_matrix = df_tender_matrix.to_numpy()

print('reading max requested lots for each bidder ...')
# read max lot each bidder is willing to take
df_max_lots_each_bidder = pd.read_excel('tender-matrix.xlsx', sheet_name='lot_each_bidder', header=0)
max_lots_each_bidder = df_max_lots_each_bidder.to_numpy().flatten()

print('correcting maximum number of lots to make things fair and square ...')
# your 'fair' maximum number of lots per bidder
max_lot_for_bidder = 2

# correct maximum of lots for each bidder
# reduce maximum of lots per bidder to your limit
max_lots_each_bidder[max_lots_each_bidder == 0] = max_lot_for_bidder
max_lots_each_bidder[max_lots_each_bidder > max_lot_for_bidder] = max_lot_for_bidder

print('creating a tender matrix ...')
# create a tender-placed-matrix with 1 for tender placed and 0 for no tender placed
tender_matrix = np.zeros(lot_tender_matrix.shape)
tender_matrix[lot_tender_matrix > 0] = 1
total_number_of_combinations = tender_matrix.sum(axis=1).prod()

print('creating all possible lot distributions ...')
# get indices of bidders for each lot
bidders_for_lots = [np.where(row > 0)[0] for row in tender_matrix]

# Get all possible combinations of indices
combinations = np.array(np.meshgrid(*bidders_for_lots)).T.reshape(-1, len(bidders_for_lots))

print('correcting lot distribution by request and fairness ...')
# Filter out combinations that contain more lots for bidders than planed
for idx, max_lots in enumerate(max_lots_each_bidder):
    combinations = combinations[(np.sum(combinations == idx, axis=1) <= max_lots)]

print('brute forcing our way to save tax payers money ...')
best_case_sum = 0
best_case_distribution = None

for idx in range(len(combinations)):
    if best_case_sum == 0 or lot_tender_matrix[np.arange(len(combinations[0])), combinations[0]].sum() < best_case_sum:
        # save total tender and combination
        best_case_sum = lot_tender_matrix[np.arange(len(combinations[idx])), combinations[0]].sum()
        best_case_distribution = combinations[idx]

df_best_distribution = pd.DataFrame({'lot': df_tender_matrix.index, 'bidder': df_tender_matrix.columns.values[combinations[0]]})
print('Best case tender sum:', best_case_sum)
print('Best case lot distribution:')
print(df_best_distribution)
print('Exporting best distribution to Excel ...')
new_bidder_matrix = np.zeros(tender_matrix.shape)
new_bidder_matrix[np.arange(len(best_case_distribution)), best_case_distribution] = 1
new_df_bidder_matrix = df_tender_matrix.copy()
new_df_bidder_matrix.iloc[:, :] = new_bidder_matrix

new_df_tender_matrix = df_tender_matrix.copy()
new_df_tender_matrix.iloc[:, :] = new_bidder_matrix * lot_tender_matrix
writer = pd.ExcelWriter(f'best_lot_distribution_option.xlsx', engine='openpyxl')
new_df_bidder_matrix.to_excel(writer, sheet_name=f'bidder_matrix')
new_df_tender_matrix.to_excel(writer, sheet_name=f'tender_matrix')
df_best_distribution.to_excel(writer, sheet_name=f'lot_list')

writer.close()

print('done and good buy')
