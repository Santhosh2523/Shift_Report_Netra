import datetime
from Batchno import batch_details
from Excel import Excel
from PDF import Excel_PDF
from ShiftA import shiftA
from ShiftB import shiftB
from ShiftC import shiftC
import time

# PATH TO DATABASE
path = "C:\\Vision\\Netra502\\Report_template.xltx"
while True:
    print("program start")
    # current date and time
    current_datetime = datetime.datetime.now().strftime('%d-%m-%Y %I:%M:%S %p')
    print(current_datetime)
    # date
    current_date = datetime.date.today()
    current_date1 = current_date.strftime("%d-%m-%Y")
    # print(current_date1)
    # timing
    A = f'{current_date1}' + " 02:00:00 PM"
    A1 = f'{current_date1}' + " 02:05:00 PM"
    B = f'{current_date1}' + " 10:00:00 PM"
    B1 = f'{current_date1}' + " 10:05:00 PM"
    C = f'{current_date1}' + " 06:00:00 AM"
    C1 = f'{current_date1}' + " 06:05:00 AM"
    # current_datetime = "17-10-2023 02:00:00 PM"

    if (A <= current_datetime <= A1) or (B <= current_datetime <= B1) or (C <= current_datetime <= C1):
        # calling batch details from ini file
        batch = batch_details()
        # print(batch)
        # batch number getting
        batch_number = batch[1]
        # print(batch_number)

        if A <= current_datetime <= A1:
            shift = "A-Shift Report"
            print(shift)
            # Getting data form database function
            production = shiftA(batch_number)
            # Excel Sheet Calling Function and Data to Excel function
            pdf = Excel(production, batch, shift, A)
            # Excel to PDF Conversion
            Excel_PDF(pdf, shift)
            print("Shift A Report generated")
            error_message = "Shift A Report generated Error"
            time.sleep(300)
            # exit()

        elif B <= current_datetime <= B1:
            shift = "B-Shift Report"
            print(shift)
            # Getting data form database function
            production = shiftB(batch_number)
            # Excel Sheet Calling Function and Data to Excel function
            pdf = Excel(production, batch, shift, B)
            # Excel to PDF Conversion
            Excel_PDF(pdf, shift)
            print("Shift B Report generated")
            error_message = "Shift B Report generated"
            time.sleep(300)
            # exit()

        elif C <= current_datetime <= C1:
            shift = "C-Shift Report"
            print(shift)
            # Getting data form database function
            production = shiftC(batch_number)
            # Excel Sheet Calling Function and Data to Excel function
            pdf = Excel(production, batch, shift, C)
            # Excel to PDF Conversion
            Excel_PDF(pdf, shift)
            print("Shift C Report generated")
            error_message = "Shift C Report generated"
            time.sleep(300)
            # exit()

        else:
            error_message = "Error in the sub function or time"
            print(error_message)
            exit()

        # writing error in text file
        file_path = "D:\\SGD Report\\Error log\\error_log.txt"
        with open(file_path, 'a') as file:
            file.write(f'{current_datetime}: {error_message}')
            file.close()
    else:
        print("Continue running sleep 3 mins")
        time.sleep(120)

# sandy changed on the 28/10/2023
