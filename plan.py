from openpyxl import load_workbook

def get_worksheet_data(wb_name, ws_name):
	wb = load_workbook(wb_name, True, False, True, False)
	rows = tuple(wb[ws_name].rows)

	def get_worksheet_pcl(rows):
		temp_x_start, temp_x_end = 999999, 999999
		temp_y_start, temp_y_end = 999999, 999999
		temp_rows = []
		for y, row in enumerate(rows):
			temp_row = []
			temp_empty_row = 0
			for x, col in enumerate(row):
				if "#" == col.value:
					temp_x_start = x
					temp_y_start = y
				if "Judgment" == col.value:
					temp_x_end = x
				if x >= temp_x_start and x <= temp_x_end:
					temp_row.append(str(col.value))
					if None == col.value:
						temp_empty_row += 1
			if temp_empty_row == (temp_x_end - temp_x_start + 1):
				temp_y_end = y
			if y >= temp_y_start and y < temp_y_end:
				temp_rows.append(temp_row)
		return tuple(temp_rows)
	
	rows = get_worksheet_pcl(rows)

	def get_testcase(rows):
		temp_testcase = []
		for y, row in enumerate(rows):
			for x, col in enumerate(row):
				if x == 0:
					if "-" in col:
						temp_testcase.append(col)
		print(len(temp_testcase))
		return temp_testcase

	def get_input_factor(rows):
		temp_x_start, temp_x_end = 999999, 999999
		temp_rows = []
		for y, row in enumerate(rows):
			temp_row = []
			for x, col in enumerate(row):
				if "Input factor" == col:
					temp_x_start = x
				if "Output element" == col:
					temp_x_end = x
				if x >= temp_x_start and x < temp_x_end:
					temp_row.append(col)
			temp_rows.append(temp_row)
		return tuple(temp_rows)
	
	get_input_factor(rows)

	def get_output_element(rows):
		temp_x_start, temp_x_end = 999999, 999999
		temp_rows = []
		for y, row in enumerate(rows):
			temp_row = []
			for x, col in enumerate(row):
				if "Output element" == col:
					temp_x_start = x
				if "Judgment" == col:
					temp_x_end = x
				if x >= temp_x_start and x < temp_x_end:
					temp_row.append(col)
			temp_rows.append(temp_row)
		return tuple(temp_rows)
	get_output_element(rows)
	pass

get_worksheet_data("d:/ut_auto/data/test.xlsx", "Sheet1")
