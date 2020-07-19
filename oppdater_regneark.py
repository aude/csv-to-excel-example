# --------------------
# import tools that we will use
# --------------------

import csv
from openpyxl import load_workbook

# --------------------
# read data from CSV into the "data_from_csv" variable
#
# then we will have the data from the CSV available here inside this Python program
# --------------------

data_from_csv = []
with open("eksempeldata.csv", "r") as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        data_from_csv.append(row)

# uncomment the next line to see the data from CSV. could be useful for debugging
# print(data_from_csv)

# --------------------
# load up the Excel file
#
# then we can start changing the Excel file using Python
# --------------------

# found out how to do this on this website: https://openpyxl.readthedocs.io/en/stable/

# load Excel file
workbook = load_workbook(filename="regneark.xlsx")

# grab the active worksheet
worksheet = workbook.active

# --------------------
# change the Excel data
#
# the goal is to add the data in "data_from_csv" to the Excel spreadsheet
#
# NB: This only changes the Excel data in the *Python object*. So only in memory.
#     If we want to save the changed Excel data to disk, we need to save to file manually.
# --------------------

# if we want to add data, we can add it under the existing data
# so one approach could be to:
#
# 1. find out where the existing data stops
# 2. add the new data after
#
# doing that here

row_number = None

# find first cell under B2 that does not have a value
for row_number in range(3, 10000):

    row_name = "B" + str(row_number)
    # print(row_name)

    row = worksheet[row_name]
    # print(row.value)

    if row.value is None:
        break

# print(row_name)

# add data
for csv_row in data_from_csv:

    worksheet["B" + str(row_number)] = csv_row["weight"]
    worksheet["C" + str(row_number)] = csv_row["timestamp"]

    row_number += 1

# --------------------
# Save to file
# --------------------

# save to a new file, to avoid ruining the old file if there are bugs in this program
workbook.save("regneark-updated.xlsx")
