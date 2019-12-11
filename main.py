from os import system, path
from gc import collect
from pandas import read_csv

class ws_handle():
    def __init__(self, wb_dir, wb_name, ws_name):
        from openpyxl import load_workbook
        wb = load_workbook(path.join(wb_dir, wb_name), data_only=True)
        self.ws = wb[ws_name]
        del wb
        collect()
        #print(self.ws)

    def coor_find_cell(self, target, match_case:bool=False, find_all:bool=False):
        return_value = list()
        # If target is string type
        if isinstance(target, str):
            for rows in self.ws:
                for cell in rows:
                    target_found = False
                    if match_case:
                        target_found = target in cell.value
                    else:
                        target_found = target == cell.value
                    if target_found:
                        del target_found
                        not_merged_cell = True
                        for merged_cell in ws.merged_cells:
                            if (cell.coordinate in merged_cell):
                                temp = (merged_cell.bounds[0],
                                        merged_cell.bounds[1],
                                        merged_cell.bounds[2],
                                        merged_cell.bounds[3])
                                not_merged_cell = False
                                if find_all:
                                    return_value.append(temp)
                                else:
                                    return tuple(temp)
                        if not_merged_cell:
                            temp = (cell.column,
                                    cell.row,
                                    cell.column,
                                    cell.row)
                            if find_all:
                                return_value.append(temp)
                            else:
                                return tuple(temp)
        # If target is list type, list format [column, row]
        elif isinstance(target, list):
            # list  = [col, row]
            coor = ws.cell(column=target[0], row=target[1]).coordinate
            not_merged_cell = True
            for merged_cell in ws.merged_cells:
                if (coor in merged_cell):
                    temp = {merged_cell.bounds[0],
                            merged_cell.bounds[1],
                            merged_cell.bounds[2],
                            merged_cell.bounds[3]}
                    not_merged_cell = False
                    return tuple(temp)
            if not_merged_cell:
                temp = {target[0],
                        target[1],
                        target[0],
                        target[1]}
                return tuple(temp)
        return return_value

def main():
    def extract_data(pd_data):
        return pd_data["implement"], \
        pd_data["workbook_dir"], \
        pd_data["workbook_name"], \
        pd_data["worksheet_name"]

    config_df = read_csv("config.csv", index_col="no")
    #print(config_df)

    i = 0
    while True:
        try:
            imp, wb_dir, wb_name, ws_name = extract_data(config_df.iloc[i])
            #print(imp, wb_dir, wb_name, ws_name)
            if ("yes" == imp):
                ws_obj = ws_handle(wb_dir, wb_name, ws_name)
                #ws_obj.debug()
            i += 1
        except:
            break

if __name__ == "__main__":
    system("cls")
    main()
