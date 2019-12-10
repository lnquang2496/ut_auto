from os import system
from gc import collect
import pandas as pd

def ws_handle(wb_dir, wb_name, ws_name):
    pass

def main():
    def extract_data(pd_data):
        return pd_data["implement"], \
        pd_data["workbook_dir"], \
        pd_data["workbook_name"], \
        pd_data["worksheet_name"]

    config_df = pd.read_csv("config.csv", index_col="no")

    i = 0
    while True:
        try:
            imp, wb_dir, wb_name, ws_name = extract_data(config_df.iloc[i])
            #print(imp, wb_dir, wb_name, ws_name)
            if ("yes" == imp):
                ws_handle(wb_dir, wb_name, ws_name)
            i += 1
        except:
            break
    collect()

if __name__ == "__main__":
    system("clear")
    main()
