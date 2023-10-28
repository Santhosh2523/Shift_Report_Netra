import datetime


def write_error(error_message1):
    current_time = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    file_name = "D:\\SGD Report\\Error log\\error_log.txt"
    with open(file_name, "a") as file1:
        file1.write("[" + current_time + "] " + error_message1 + "\n")
        file1.close()
