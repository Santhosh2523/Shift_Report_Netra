import datetime
import pyodbc

batch_no = 12345678

con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Vision\Netra502\netrav2.mdb;'
conn = pyodbc.connect(con_string)
print("Connected to Database")
# print(batch_no)
cur = conn.cursor()
# date
current_date = datetime.date.today()
current_date1 = current_date.strftime("%d-%m-%Y")
# print(current_date1)
A = "08-06-2021 06:00:00 AM"
# print(A)
A1 = "08-06-2021 02:00:00 PM"
# print(A1)
query = f"select TOP 1 * from {batch_no} where ldate >= ? and ldate <= ? order by ldate ASC"
print(query)
cur.execute(query, (A, A1))
result = cur.fetchone()
print(result)

query1 = f"select TOP 1 * from {batch_no} where ldate >= ? and ldate <= ? order by ldate DESC"
print(query1)
cur.execute(query1, (A, A1))
result1 = cur.fetchone()
print(result1)
if result1:
     _, var2, var3, var4, var5, var6, var7, var8, var9, var10 = result1
else:
    print("Error")

