# @author Shaan Lehal 24 May 2021

# importing necessary libraries
import xlrd
import xlsxwriter

# Give the location of the file -- Mine is meme.xlsx
loc = ("meme.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)



# initialize the list that will store all of the viable email addresses
list_of_email_addresses = []


# iterate over every school in the spreadsheet

for x in range(4, 2531):

    # boolean necessary to prevent overcounting
    bool = False

    # the integers in the list correspond to the columns in the spreadsheet
    # that contain data about specific grades -- 37-39 denotes grades 6-8
    # inclusive

    for y in [37, 38, 39]:

        # the value in the spreadsheet will be 1 if the school contains students
        # in the grade specified
        if (sheet.cell_value(x, y) == 1.0 and bool == False):

            bool = True
            # append the email address to the list
            list_of_email_addresses.append(sheet.cell_value(x, 10))


# prints the list of addresses
# print(list_of_email_addresses)


# exporting the list to a new excel spreadsheet


workbook = xlsxwriter.Workbook('Example2.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell
# Rows and columns are zero indexed
row = 0
column = 0


# iterating through content list
for item in list_of_email_addresses:

    # write operation
    worksheet.write(row, column, item)

    # incrementing the value of row by one
    # with each iteraton
    row += 1

workbook.close()
