import os
import xlrd
from scripts import writeAllData, writeServices, writeProduct


def main():
    cur_path = os.path.dirname(__file__)
    input_path = os.path.relpath("../input/data.xlsm", cur_path)
    data = {"service": None, "name": None, "day": None, "month": None, "year": None, "cost": None, "product": None,
            "quantity": None}
    workbook = xlrd.open_workbook(input_path)
    worksheet = workbook.sheet_by_name("Service_data")
    data_values = []
    for rows in range(1, worksheet.nrows):
        data_values.append(worksheet.cell_value(rows, 1))
    data["service"] = str(data_values[0])
    data["name"] = str(data_values[1])
    data["day"] = int(data_values[2])
    data["month"] = int(data_values[3])
    data["year"] = int(data_values[4])
    data["cost"] = float(data_values[5])
    data["cost"] = str(data["cost"]).replace(".", ",")
    data["product"] = str(data_values[6])
    data["quantity"] = int(data_values[7])
    print(data)

    if data["service"] is not None and data["product"] is not None:
        writeAllData.createDocAll(data)
    elif data["service"] is None and data["product"] is not None:
        writeProduct.createDocProducts(data)
    elif data["service"] is not None and data["product"] is None:
        writeServices.createDocServices(data)
    elif data["service"] is None and data["product"] is None:
        print("ERROR: No service or product specified")


main()
