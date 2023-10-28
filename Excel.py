import traceback
from Error_log import write_error
import openpyxl
import datetime
from openpyxl.styles import Font


def Excel(production, batch, shift, a):
    try:
        wb = openpyxl.load_workbook("C:\\Vision\\Netra502\\Report_template.xltx")
        worksheet = wb['Sheet1']
        list1 = batch
        print("list1:" + f'{list1}')
        list2 = production[0]  # shift production data
        print("list2:" + f'{list2}')
        list3 = production[1]  # batch production data
        print("list3:" + f'{list3}')
        list5 = production[2]
        print(list5)
        list4 = int(list1[2]) - int(list3[2])  # Remaining qty
        # Machine name
        worksheet['B2'].value = list1[4]
        # Article
        worksheet['B3'].value = list1[0]
        # Change the value to the desired font size
        font = Font(size=8)
        worksheet['B3'].font = font
        # bath number
        worksheet['B4'].value = list1[1]
        # shift
        worksheet['E4'].value = shift
        # batch qty
        worksheet['E7'].value = list1[2]
        # Remaining qty
        worksheet['E8'].value = list4
        # Batch data
        # date & time
        # worksheet['E3'].value = list1[3]
        # produce
        worksheet['B6'].value = list3[1]
        # Accept
        worksheet['B7'].value = list3[2]
        # Reject
        worksheet['B8'].value = list3[3]
        # Cam 1 Rejection
        worksheet['B9'].value = list3[4]
        # Cam 2 Rejection
        worksheet['B10'].value = list3[5]
        # Cam 3 Rejection
        worksheet['B11'].value = list3[6]
        # Cam 4 Rejection
        worksheet['B12'].value = list3[7]
        # Cam 5 Rejection
        worksheet['B13'].value = list3[8]
        # start date batch
        worksheet['E6'].value = list1[3]
        # yield %
        # worksheet['E9'].value = list3[3]

        # shift data
        # produce
        worksheet['B15'].value = list2[1]
        # Accept
        worksheet['B16'].value = list2[2]
        # Reject
        worksheet['B17'].value = list2[3]
        # Cam 1 Rejection
        worksheet['B18'].value = list2[4]
        # Cam 2 Rejection
        worksheet['B19'].value = list2[5]
        # Cam 3 Rejection
        worksheet['B20'].value = list2[6]
        # Cam 4 Rejection
        worksheet['B21'].value = list2[7]
        # Cam 5 Rejection
        worksheet['B22'].value = list2[8]
        # start date batch
        worksheet['E16'].value = a
        # yield %
        # worksheet['E17'].value = list2[9]
        # date
        current_date = datetime.date.today()
        current_date1 = current_date.strftime("%d-%m-%Y")
        # print(current_date1)
        path = f'D:\\SGD Report\\Shift Report\\Report Excel\\{current_date1 + " " + shift}.xltx'
        print(path)
        wb.save(path)
        wb.close()
        print('Excel Completed')
        return path
    except Exception as e:
        error_message = traceback.format_exc()  # Get the detailed error message
        write_error(error_message)
        # print(error_message)
        exit()
