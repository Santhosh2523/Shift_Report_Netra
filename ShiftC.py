import traceback
import pyodbc
import datetime
from datetime import timedelta
from Error_log import write_error


def shiftC(batch_no):
    try:
        con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Vision\Netra502\netrav2.mdb;'
        conn = pyodbc.connect(con_string)
        print("Connected to Database")
        # print(batch_no)
        cur = conn.cursor()
        # date
        current_date = datetime.date.today()
        current_date1 = current_date.strftime("%d-%m-%Y")
        # print(current_date1)
        yesterday = current_date - timedelta(days=1)
        # print(yesterday)
        yesterday1 = yesterday.strftime("%d-%m-%Y")
        # print(yesterday1)
        C = f'{yesterday1}' + " 10:00:00 PM"
        # print(A)
        C1 = f'{yesterday1}' + " 10:01:00 PM"
        # print(A1)
        query1 = f"select * from {batch_no} where ldate >= ? and ldate <= ?"
        # print(query1)
        cur.execute(query1, (C, C1))
        result1 = cur.fetchone()
        # print(result1)
        if result1:
            _, _, _, _, _, _, _, _, _, _ = result1
        else:
            print("Error")

        C2 = f'{current_date1}' + " 06:00:00 AM"
        # print(A)
        C21 = f'{current_date1}' + " 06:01:00 AM"
        # print(A1)
        query2 = f"select * from {batch_no} where ldate >= ? and ldate <= ?"
        # print(query1)
        cur.execute(query2, (C2, C21))
        result2 = cur.fetchone()
        # print(result2)
        if result2:
            _, _, _, _, _, _, _, _, _, _ = result2
        else:
            print("Error")

        # print(result1)
        # print(result2)
        list1 = []
        for i in range(len(result1)):
            list1.append(result2[i] - result1[i])

        # print(list1)
        return list1, result2,
    except Exception as e:
        error_message = traceback.format_exc()  # Get the detailed error message
        write_error(error_message)
        # print(error_message)
        exit()
