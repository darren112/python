import openpyxl

# create a new workbook
workbook = openpyxl.Workbook()

# create a new worksheet named "Answer Sheet"
worksheet = workbook.active
worksheet.title = "Answer Sheet"

# add the Selection column header
worksheet['A1'] = 'Selection'

# add the answer choices for each selection
for i in range(1, 101):
    # add the selection number
    worksheet.cell(row=i+1, column=1, value=i)

    # add the answer dropdown validation
    if i >= 7 and i <= 31:
        dv = openpyxl.worksheet.datavalidation.DataValidation(
            type="list", formula1='"A,B,C"', allow_blank=True)
    else:
        dv = openpyxl.worksheet.datavalidation.DataValidation(
            type="list", formula1='"A,B,C,D"', allow_blank=True)
    worksheet.add_data_validation(dv)
    dv.add(f"B{i+1}")

# save the workbook
workbook.save('answer_sheet.xlsx')
