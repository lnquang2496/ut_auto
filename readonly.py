from openpyxl import load_workbook

def get_worksheet_data(wb_name, ws_name):
	wb = load_workbook(wb_name, True, False, True, False)
	rows = tuple(wb[ws_name].rows)
	for row in rows:
		for col in row:
			pass
	pass

get_worksheet_data("d:/ut_auto/data/test.xlsx", "Sheet1")
