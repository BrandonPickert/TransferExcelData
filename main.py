import pandas as pd
import openpyxl

TM_List = "C:\\Users\\bpickert\\PycharmProjects\\TransferExcelData\\TM_List.xlsx"
carb = "C:\\Users\\bpickert\\PycharmProjects\\TransferExcelData\\CARB.xlsx"


def delete_data(file_path, sheet_name):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    last_row = sheet.max_row
    delete_range = "B6:BW" + str(last_row)
    for row in sheet[delete_range]:
        for cell in row:
            cell.value = None
    # sheet[delete_range].clear()
    workbook.save(file_path)
    print("Data deletion complete.")

sheet_name = "Sheet1"
delete_data(carb, sheet_name)

df = pd.read_excel(TM_List, skiprows=range(0, 15), index_col=False, header=0)
print(df)

writer = pd.ExcelWriter(carb, engine='openpyxl', mode='a', if_sheet_exists='overlay', date_format='MM/DD/YYYY', datetime_format='MM/DD/YYYY')
df.to_excel(writer, sheet_name='Sheet1', startcol=1, startrow=5, index=False, header=False)
workbook = writer.book
worksheet = writer.sheets[sheet_name]
# format1 = workbook.add_format({"num_format": "MM/DD/YYYY"})
# worksheet.set_column('T:U',None,format1)
writer.close()

print("Done")