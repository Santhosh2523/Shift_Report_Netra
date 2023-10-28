import traceback
import pyodbc
import datetime
from Error_log import write_error


def shiftB(batch_no):
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
        B = f'{current_date1}' + " 02:00:00 PM"
        # print(A)
        B1 = f'{current_date1}' + " 02:01:00 PM"
        # print(A1)
        query1 = f"select * from {batch_no} where ldate >= ? and ldate <= ?"
        # print(query1)
        cur.execute(query1, (B, B1))
        result1 = cur.fetchone()
        # print(result1)
        if result1:
            _, var2, var3, var4, var5, var6, var7, var8, var9, var10 = result1
        else:
            print("Error")

        B2 = f'{current_date1}' + " 10:00:00 PM"
        # print(A)
        B21 = f'{current_date1}' + " 10:01:00 PM"
        # print(A1)
        query2 = f"select * from {batch_no} where ldate >= ? and ldate <= ?"
        # print(query1)
        cur.execute(query2, (B2, B21))
        result2 = cur.fetchone()
        # print(result2)
        if result2:
            _, var12, var13, var14, var15, var16, var17, var18, var19, var20 = result2
        else:
            print("Error")

        # print(result1)
        # print(result2)
        list1 = []
        for i in range(len(result1)):
            list1.append(result2[i] - result1[i])

        # print(list1)
        return list1, result2

    except Exception as e:
        error_message = traceback.format_exc()  # Get the detailed error message
        write_error(error_message)
        # print(error_message)
        exit()
