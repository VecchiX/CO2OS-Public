import openpyxl

print("\n\n CONFIG(EXCEL)取得処理 Ver.2024.04.10.磯部\n")

dictionary = {}

def Get_Info(sheet_name, file_name="CONFIG.xlsx"):

    status = ""

    try:
        #Excelからの情報取得
        wb = openpyxl.load_workbook(file_name, read_only=True, data_only=True)
        ws = wb[sheet_name]

        status = "OPEN"

        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            key, data = row

            if key is not None:
                if data is None:
                    dictionary[key] = ''
                else:
                    dictionary[key] = data

        wb.close()
        status = "SUCCESS"

    except:
        if status == "OPEN":
            wb.close()

        return False

    finally:

        return True

