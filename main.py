from os import system, path
from gc import collect
from pandas import read_csv

class ws_handle():
    def __init__(self, wb_dir, wb_name, ws_name):
        from openpyxl import load_workbook
        wb = load_workbook(path.join(wb_dir, wb_name), data_only=True)
        ws = wb[ws_name]
        del wb
        print(ws)

def main():
    def extract_data(pd_data):
        return pd_data["implement"], \
        pd_data["workbook_dir"], \
        pd_data["workbook_name"], \
        pd_data["worksheet_name"]

    config_df = read_csv("config.csv", index_col="no")

    i = 0
    while True:
        try:
            imp, wb_dir, wb_name, ws_name = extract_data(config_df.iloc[i])
            #print(imp, wb_dir, wb_name, ws_name)
            if ("yes" == imp):
                ws = ws_handle(wb_dir, wb_name, ws_name)
            i += 1
        except:
            break

if __name__ == "__main__":
    system("cls")
    main()
