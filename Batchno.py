import configparser
import traceback
from Error_log import write_error


def batch_details():
    try:
        # create a config object
        config = configparser.ConfigParser()

        # path
        path = 'C:\\Vision\\Netra502\\Product details.ini'

        # read the INI file
        config.read(path)

        # get value from file
        article1 = config.get('Production data', 'article')
        article = article1.strip('"')
        # print(article)
        batch_no11 = config.get('Production data', 'batch no')
        batch_no = batch_no11.strip('"')
        # print(batch_no)
        batch_qty = config.get('Production data', 'planned qty')
        # print(batch_qty)
        datetime11 = config.get('Production data', 'datetime')
        datetime1 = datetime11.strip('"')
        # print(datetime1)
        Machine_name = config.get('Production data', 'Machine name')
        # print(Machine_name)
        batch = [article, batch_no, batch_qty, datetime1, Machine_name]
        # print(batch_details)
        print("Batchno Completed")
        # print(batch)
        return batch

    except Exception as e:
        error_message = traceback.format_exc()  # Get the detailed error message
        write_error(error_message)
        # print(error_message)
        exit()
