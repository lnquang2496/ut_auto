from openpyxl import load_workbook
# Number 1 - load the worksheet
def ut_load_worksheet(workbook, worksheet):
	l_workbook = load_workbook(workbook, False, False, True, False)
	l_worksheet = l_workbook[worksheet]

	return l_worksheet
# Number 2 - load the pcl data
def ut_load_pcl_data_init_pcl_range():
	l_pcl_range = [
		1000000,	# 0 - first column
		1000000,	# 1 - first row
		2000000,	# 2 - last column
		2000000,	# 3 - last row
		0]			# 4 - blank cell count

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
			"fst_col" : cell_col,		# Cell first column coordinate
			"lst_col" : -1,				# Cell last column coordinate
			"fst_row" : cell_row,		# Cell first row coordinate
			"lst_row" : -1,				# Cell last row coordinate
			"value"   : str(cell.value)	# Cell value
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
# Number 3 - load the pcl information
def ut_load_pcl_info(pcl_data):
	l_pcl_info = {
		"tc_fst_row"	: -1,		# Test case first row
		"tc_lst_row"	: -1,		# Test case last row
		"tc_ini_row"	: -1,		# Test case initial variable row
		"inp_avail"		: False,	# Input factor availability
		"inp_start"		: -1,		# Input factor start column
		"inp_end"		: -1,		# Input factor end column
		"out_avail"		: False,	# Output element availability
		"out_start"		: -1,		# Output element start column
		"out_end"		: -1,		# Output element end column
		"jud_col"		: -1,		# Judgment column
		"tc_list"		: []		# Test case list
	}

	for l_y, l_row in enumerate(pcl_data):
		for l_x, l_col in enumerate(l_row):
			if l_x == 0:
				if ("-" in l_col["value"]):
					if (l_pcl_info["tc_fst_row"] == -1):
						l_pcl_info["tc_fst_row"] = l_y
						l_pcl_info["tc_ini_row"] = l_y - 1
					l_pcl_info["tc_list"].append(l_col["value"])
			if l_y == 0:
				if "Input factor" in l_col["value"]:
					l_pcl_info["inp_avail"] = True
					l_pcl_info["inp_start"] = l_x
				elif "Output element" in l_col["value"]:
					l_pcl_info["out_avail"] = True
					l_pcl_info["out_start"] = l_x
					if l_pcl_info["inp_avail"]:
						l_pcl_info["inp_end"] = l_x - 1
				elif "Judgment" in l_col["value"]:
					l_pcl_info["jud_col"] = l_x
					if l_pcl_info["out_avail"]:
						l_pcl_info["out_end"] = l_x - 1
					else:
						l_pcl_info["inp_end"] = l_x - 1
	else:
		l_pcl_info["tc_lst_row"] = l_y

	return l_pcl_info
# Number 4 - load the test info
def is_pointer(value):
			with open("data/pointer.txt", "r") as f:
				for l in f:
					l = l.rstrip("\n")
					if l != "*":
						l += " "
					if l in value:
						return l
				else:
					return None

def is_structure(value):
	with open("data/structure.txt", "r") as f:
		for l in f:
			l = l.rstrip("\n")
			l += " "
			if l in value:
				return l
		else:
			return None

def is_array(value):
	if ("[" in value) and ("]" in value):
		return value[value.find("[") +1 : value.find("]")]
	else:
		return None

def load_argu_info_load_info_load_expect(name, pcl_data, pcl_info):
	l_return = []

	if pcl_info["out_avail"]:
		for l_index in range(pcl_info["out_start"], pcl_info["out_end"] + 1):
			l_cell = pcl_data[1][l_index].copy()
			if ("[a]" in l_cell["value"]) or ("[g]" in l_cell["value"]):
				l_cell_value = l_cell["value"].replace("[a]", "").replace("[g]", "").replace("*", "")

				if name == l_cell_value:
					if l_cell["lst_row"] + 1 < pcl_info["tc_ini_row"]:
						for l_index_2 in range(l_cell["fst_col"], l_cell["lst_col"] + 1):
							l_cell_2 = pcl_data[l_cell["lst_row"] + 1][l_index_2].copy()
							l_array = is_array(l_cell_2["value"])
							if l_array:
								l_return.append(f"{l_cell_value}[{l_array}]")
								continue
					else:
						l_return.append(l_cell_value)

	return l_return

def load_argu_info_load_info(cell, pcl_data, pcl_info, parent=[]):
	l_return = {
		"declare"	: "",
		"type"		: "",
		"name"		: "",
		"is_pointer"	: False,
		"is_structure"	: False,
		"is_array"		: False,
		"array"			: [],
		"parent"		: [],
		"child"			: [],
		"expect"		: []
	}

	l_cell_value = cell["value"]

	l_return["declare"] = l_cell_value

	if parent:
		l_return["parent"].append(parent)

	l_is_pointer = is_pointer(l_cell_value)
	l_is_structure = is_structure(l_cell_value)
	l_is_array = is_array(l_cell_value)

	if l_is_pointer:
		l_return["is_pointer"] = True

		if "*" == l_is_pointer:
			l_cell_value = l_cell_value.replace("*", "").replace("  ", " ")

	if l_is_structure:
		l_return["is_structure"] = True

	if l_is_array:
		l_return["is_array"] = True

	l_return["name"] = l_cell_value[l_cell_value.rfind(" ") + 1 : ].replace(" ", "")

	l_return["type"] = l_cell_value[ : l_cell_value.rfind(" ")]

	# Load expect
	l_return["expect"] = load_argu_info_load_info_load_expect(l_return["name"], pcl_data, pcl_info)

	if cell["lst_row"] + 1 < pcl_info["tc_ini_row"]:
		for l_index in range(cell["fst_col"], cell["lst_col"] + 1):
			l_cell = pcl_data[cell["lst_row"] + 1][l_index].copy()

			if l_is_pointer:
				l_array = is_array(l_cell["value"])
				if l_array:
					l_return["array"].append(l_array)
					continue

			if l_cell["value"] != "None":
				l_argument_info = load_argu_info_load_info(l_cell, pcl_data, pcl_info, l_return["name"])
				l_return["child"].append(l_argument_info)

	return l_return

def ut_load_test_info_load_argu_info(pcl_data, pcl_info):
	l_return = []
	if pcl_info["inp_avail"]:
		for l_index in range(pcl_info["inp_start"], pcl_info["inp_end"] + 1):
			l_cell = pcl_data[1][l_index].copy()
			if ("[a]" in l_cell["value"]) or ("[g]" in l_cell["value"]):
				l_cell["value"] = l_cell["value"].replace("[a]", "").replace("[g]", "")
				l_argument_info = load_argu_info_load_info(l_cell, pcl_data, pcl_info)
				l_return.append(l_argument_info)

	return l_return


def load_stub_info_load_info(cell, pcl_data, pcl_info, cur_col, stub_info):
	l_return = {
		"name"		: "",
		"ret_type"	: "",
		"ret_col"	: -1,
		"out_col"	: [],
		"exp_name"	: [],
		"exp_col"	: [],
		"count"		: 0
	}

	l_cell_value = cell["value"]
	l_return["ret_type"] = l_cell_value[ : l_cell_value.find(" ")]
	l_return["name"] = l_cell_value[l_cell_value.find(" ") : l_cell_value.find("(")]
	l_return["ret_col"] = cur_col

	if stub_info:
		for l_stub in stub_info:
			if l_return["name"] == l_stub["name"]:
				l_return["count"] = l_stub["count"] + 1

	l_cell = pcl_data[1][cur_col + 1].copy()
	if ("[f]" in l_cell["value"]):
		for l_index in range(l_cell["fst_col"], l_cell["lst_col"] + 1):
			l_cell_output = pcl_data[l_cell["lst_row"] + 1][l_index].copy()
			l_cell_output_value = l_cell_output["value"].replace("*", "")
			l_return["out_col"].append(l_index)

	if pcl_info["out_avail"]:
		l_temp_count = 0
		for l_index in range(pcl_info["out_start"], pcl_info["out_end"]):
			l_cell = pcl_data[1][l_index].copy()
			l_cell_value = l_cell["value"]
			if "[f]" in l_cell_value:
				l_temp_name = f" {l_return['name']}("
				if l_temp_name in l_cell_value:
					if l_temp_count == l_return["count"]:
						for l_check_index in range(l_cell["fst_col"], l_cell["lst_col"] + 1):
							l_cell_check = pcl_data[l_cell["lst_row"] + 1][l_check_index].copy()
							l_cell_check_value = l_cell_check["value"]
							# TODO
							
					l_temp_count += 1

	return l_return

def ut_load_test_info_load_stub_info(pcl_data, pcl_info):
	l_return = []

	if pcl_info["out_avail"]:
		for l_index in range(pcl_info["out_start"], pcl_info["out_end"] + 1):
			l_cell = pcl_data[1][l_index].copy()
			if ("[rt]" in l_cell["value"]):
				l_cell["value"] = l_cell["value"].replace("[rt]", "")
				l_stub_info = load_stub_info_load_info(l_cell, pcl_data, pcl_info, l_index, l_return)
				l_return.append(l_stub_info)

	return l_return

def ut_load_test_info_load_ret_val_info(pcl_data,pcl_info):
	l_return = {
		"ret_avail"	: False,
		"ret_col"	: -1
	}

	if pcl_info["out_avail"]:
		for l_index in range(pcl_info["out_start"], pcl_info["out_end"] + 1):
			l_cell = pcl_data[1][l_index].copy()
			if ("Return value" in l_cell["value"]):
				l_return["ret_avail"] = True
				l_return["ret_col"] = l_index

	return l_return

def ut_load_test_info(pcl_data, pcl_info):
	l_test_info = {
		"argument_info" : ut_load_test_info_load_argu_info(pcl_data, pcl_info),		# Test argument information
		"stubfunc_info" : ut_load_test_info_load_stub_info(pcl_data, pcl_info),		# Test stub function information
		"ret_val_info"	: ut_load_test_info_load_ret_val_info(pcl_data, pcl_info)	# Test return value
	}

	return l_test_info
# Main script
l_worksheet = ut_load_worksheet("data/test.xlsx", "Sheet1")
l_pcl_data = ut_load_pcl_data(l_worksheet)
l_pcl_info = ut_load_pcl_info(l_pcl_data)
l_test_info = ut_load_test_info(l_pcl_data, l_pcl_info)
pass
