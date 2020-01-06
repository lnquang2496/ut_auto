from openpyxl import load_workbook

def ut_load_worksheet(workbook, worksheet):
	l_workbook = load_workbook(workbook, False, False, True, False)
	l_worksheet = l_workbook[worksheet]

	return l_worksheet

def ut_load_pcl_data_init_pcl_range():
	l_pcl_range = [
		1000000, # 0 - first column
		1000000, # 1 - first row
		2000000, # 2 - last column
		2000000, # 3 - last row
		0]       # 4 - blank cell count

	return l_pcl_range

def ut_load_pcl_data_get_pcl_range(pcl_range, cell_col, cell_row, cell):
	l_total_blank_cell = pcl_range[2] - pcl_range[0]
	l_cell_value = str(cell.value)

	if "#" == l_cell_value:
		pcl_range[0] = cell_col
		pcl_range[1] = cell_row
	elif "Judgment" == l_cell_value:
		pcl_range[2] = cell_col
	elif "None" == l_cell_value:
		pcl_range[4] += 1

	if l_total_blank_cell == pcl_range[4]:
		pcl_range[3] = cell_row - 1

	return pcl_range

def ut_load_pcl_data_get_cell_info(worksheet, pcl_range, cell_col, cell_row, cell):
	if (pcl_range[0] <= cell_col <= pcl_range[2]) and (pcl_range[1] <= cell_row <= pcl_range[3]):
		cell_col -= pcl_range[0]
		cell_row -= pcl_range[1]
		l_cell_info = {
			"fst_col" : cell_col,
			"lst_col" : -1,
			"fst_row" : cell_row,
			"lst_row" : -1,
			"value"   : str(cell.value)
		}

		l_merged_cells = worksheet.merged_cells
		for l_merged in l_merged_cells:
			if cell.coordinate in l_merged:
				l_cell_info["lst_col"] = cell_col + (l_merged.bounds[2] - l_merged.bounds[0])
				l_cell_info["lst_row"] = cell_row + (l_merged.bounds[3] - l_merged.bounds[1])
				break
		else:
			l_cell_info["lst_col"] = cell_col
			l_cell_info["lst_row"] = cell_row

		return l_cell_info
	else:
		return None

def ut_load_pcl_data(worksheet):
	l_pcl_range = ut_load_pcl_data_init_pcl_range()
	l_pcl_data = []
	for l_y, l_row in enumerate(worksheet.rows):
		l_pcl_data_row = []
		l_pcl_range[4] = 0
		for l_x, l_col in enumerate(l_row):
			l_pcl_range = ut_load_pcl_data_get_pcl_range(l_pcl_range, l_x, l_y, l_col)
			l_cell_info = ut_load_pcl_data_get_cell_info(worksheet, l_pcl_range, l_x, l_y, l_col)
			if l_cell_info:
				l_pcl_data_row.append(l_cell_info)
		if (l_y > l_pcl_range[3]):
			break
		if l_pcl_data_row:
			l_pcl_data.append(l_pcl_data_row)
	return l_pcl_data

l_worksheet = ut_load_worksheet("data/test.xlsx", "Sheet1")
l_pcl_data = ut_load_pcl_data(l_worksheet)
pass
