import traceback
from Error_log import write_error
from win32com import client
import datetime
import os


def Excel_PDF(path, shift):
    try:
        app = client.Dispatch("Excel.Application")
        app.Interactive = False
        app.visible = False
        app.DisplayAlerts = False  # Disable alerts

        # Excel File path
        # date
        current_date = datetime.date.today()
        current_date1 = current_date.strftime("%d-%m-%Y")
        # print(current_date1)
        save_path = f'D:\\SGD Report\\Shift Report\\Report PDF\\{current_date1} {shift}.pdf'
        workbook = app.Workbooks.Open(path)
        workbook.SaveAs(save_path, FileFormat=57)  # FileFormat=57 saves as PDF
        workbook.Close()
        # Close the Excel application
        app.Quit()
        print("PDF Completed")
        # Delete the Excel file
        os.remove(path)
    except Exception as e:
        error_message = traceback.format_exc()  # Get the detailed error message
        write_error(error_message)
        # print(error_message)
        exit()
