import os.path
import pandas as pd
from datetime import datetime
from collections import Counter

# Get a list of all files with extension '.xls'

path = r'C:\Users\v.makarov\PycharmProjects\Material_consumption'
pyname = [name for name in os.listdir(path) if name.endswith('.xls')]


# A function that will fix the date data type
def modify_time(data: datetime) -> float:
    if str(type(data)) == "<class 'datetime.time'>":  # Checking if a variable belongs data to datetime.time
        data = datetime(1970, 1, 1)  # If it belongs, then we equate the date to the beginning, since the variable
        # contains only time data
    elif type(data) == str:  # Checking if a variable belongs data to text
        data = datetime.strptime(date, "%d.%m.%Y")  # If it belongs, then convert to datetime
    else:
        data = data  # If not, then leave it unchanged.
    data = (data - datetime(1970, 1, 1)).total_seconds()  # Subtracts the original date from the date and converts to
    # seconds
    return data  # Return data


# Function for data generation
def add_requirement(k, temp_data, data):
    name = 'x' + str(k)  # Name for column
    temp_data = temp_data.rename(columns={temp_data.columns[0]: 'part', temp_data.columns[1]: name})  # Rename columns
    temp_data.loc[-1] = ['time', date]  # Add first row 'time' that contains time on seconds
    temp_data.index = temp_data.index + 1
    temp_data = temp_data.sort_index()
    temp_data = temp_data.groupby('part').sum().reset_index()  # Grouping the details
    if len(data) == 0:  # If data is null
        data = temp_data  # then we equate temp_data
    else:
        data = data.merge(temp_data, on='part', how='outer', indicator=False).fillna(0)  # Otherwise merge temp_data and
        # data

    return data  # Return data


# We start a cycle that will go through all the files, select the necessary data

requirement = pd.DataFrame()  # Creat null dataframe
# Since the data is filled in differently, having examined the file manually, the places were identified, and where the
# order date is located, and where the order data is located. Get the required rows and columns with "iloc".
for i in range(len(pyname)):  # Let's loop through the list of files and get the data
    temp = pd.read_excel(pyname[i], 'Требование').fillna(0)
    if temp.iloc[8, 5] != 0:
        date = temp.iloc[8, 5]
    elif temp.iloc[6, 5] != 0:
        date = temp.iloc[6, 5]
    elif temp.iloc[6, 7] != 0:
        date = temp.iloc[6, 7]
    else:
        date = temp.columns[7]
    date = modify_time(date)
    temp = temp.iloc[11:, [2, 6]]
    requirement = add_requirement(i, temp, requirement)

# Repeat all procedures for the second data sheet
for j in range(len(pyname)):
    temp = pd.read_excel(pyname[j], 'Накладная склада').fillna(0)
    if temp.iloc[6, 7] == 0:
        date = temp.columns[7]
    else:
        date = temp.iloc[6, 7]
    date = modify_time(date)
    temp = temp.iloc[127:250, [1, 8]]
    requirement = add_requirement(j, temp, requirement)
frequency = Counter(requirement.iloc[0])  # We create a dictionary that contains the order date and the number of times
# this date is repeated

for i in frequency:  # Creat cycle for find columns with the same date
    if frequency[i] > 1:  # If the frequency of the date is greater than one
        filters = (requirement == i).any()  # Create variable 'filters' for find columns with the same date
        name = (requirement.loc[:, filters]).columns  # Name of columns with the same date
        requirement[i] = requirement[name].sum(axis=1)  # The sum of duplicates
        requirement.loc[0, i] = i  # Rewrite the value of the time, as it has also developed
        requirement = requirement.drop(name, axis=1)  # Delete duplicated columns

name_time = requirement.values[:1]  # Get the list with the value, which contains the date of the orders placed
requirement.columns = name_time[0]  # Rename all columns as date
requirement = requirement.drop([0], axis=0)  # Delete row 'time'
requirement.rename(columns={'time': 'part'}, inplace=True)  # Rename columns

requirement.to_csv('data.csv')  # Save dataframe csv
requirement.to_excel('data.xlsx')  # Save dataframe excel
print("Finish!")
