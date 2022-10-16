import os.path
import pandas as pd
from datetime import datetime

# We get the data of all files that have the extension '.xls'
pyname = [name for name in os.listdir() if name.endswith('.xls')]

# We take the first file from the list and read the data into a variable, replace the unknown data with 0
file = pyname[0]
excel_data_0 = pd.read_excel(file, 'Требование')
excel_data_0 = excel_data_0.fillna(0)

# During manual analysis, it was found that a cell with date data can be located in several places, so we introduce a
# check for a cell with a date. if necessary, change the date format
if excel_data_0.iloc[8, 5] == 0 and excel_data_0.iloc[6, 5] == 0:
    date = excel_data_0.columns[7]
else:
    if excel_data_0.iloc[6, 5] != 0:
        date = excel_data_0.iloc[6, 5]
    else:
        date = excel_data_0.iloc[8, 5]
if type(date) == str:
    date = datetime.strptime(date, format='%Y-%m-%dT')

# Choosing the necessary data for further analysis
data_0 = excel_data_0.iloc[11:, [2, 6]]
data_0 = data_0.rename(columns={data_0.columns[0]: 'Detal', data_0.columns[1]: date})

# Repeat all procedures for the second data sheet
excel_data_2 = pd.read_excel(file, 'Накладная склада')
excel_data_2 = excel_data_2.fillna(0)
if excel_data_2.iloc[6, 7] == 0:
    date_2 = excel_data_2.columns[7]
else:
    date_2 = excel_data_2.iloc[6, 7]
if type(date_2) == str:
    date_2 = datetime.strptime(date_2, format='%Y-%m-%dT')
excel_data_2 = excel_data_2.iloc[127:, [1,8]]
data_2 = excel_data_2.rename(columns={excel_data_2.columns[0]: 'Detal', excel_data_2.columns[1]: date_2})

# We start a cycle that will go through all the files, select the necessary data and add them to the initial
# DataFrame with help 'merge'.
for i in range(len(pyname)):
    if i > 0:
        file = pyname[i]
        excel_data_1 = pd.read_excel(file, 'Требование')
        excel_data_1 = excel_data_1.fillna(0)
        if excel_data_1.iloc[8, 5] == 0 and excel_data_1.iloc[6, 5] == 0:
            date = excel_data_1.columns[7]
        else:
            if excel_data_1.iloc[6, 5] != 0:
                date = excel_data_1.iloc[6, 5]
            else:
                date = excel_data_1.iloc[8, 5]
        if type(date) == str:
            date = datetime.strptime(date, "%d.%m.%Y")
        data_1 = excel_data_1 .iloc[11:, [2, 6]]
        data_1 = data_1.rename(columns={data_1.columns[0]: 'Detal', data_1.columns[1]: date})
        data_0 = pd.merge(data_0, data_1, on = 'Detal', how = 'outer', indicator= False)

# After creation, group data by column 'Detal'
data_0 = data_0.groupby('Detal').sum().reset_index()

# Repeat all procedures for the second data sheet
for i in range(len(pyname)):
    if i > 0:
        file = pyname[i]
        excel_data_3 = pd.read_excel(file, 'Накладная склада')
        excel_data_3 = excel_data_3.fillna(0)
        excel_data_3 = pd.DataFrame(excel_data_3)
        if excel_data_3.iloc[6, 7] == 0:
            date_3 = excel_data_3.columns[7]
        else:
            date_3 = excel_data_3.iloc[6, 7]
        if type(date_3) == str:
            date_3 = datetime.strptime(date_3, format='%Y-%m-%dT')
        excel_data_3 = excel_data_3.iloc[127:250, [1, 8]]
        data_3 = excel_data_3
        data_3 = excel_data_3.rename(columns={excel_data_3.columns[0]: 'Detal', excel_data_3.columns[1]: date_3})
        data_2 = data_2.merge(data_3,  how = 'outer', left_on='Detal', right_on='Detal', indicator = False)
        data_2 = data_2[data_2['Detal'] != 0]
data_2 = data_2.groupby(data_2.columns[0]).sum().reset_index()

# Combining the two DataFrame
data = pd.merge(data_0, data_2, how='outer', on = 'Detal', indicator= False)
data = data.groupby(data.columns[0]).sum().reset_index()

# Save to file
data.to_csv('data.csv')
data.to_excel('data.xlsx')
print('Выполнено!')