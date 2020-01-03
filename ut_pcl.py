class cell:
	def __init__(self):
		self.first_col = -1
		self.last_col = -1
		self.row = -1
		self.value = None

def ut_getdata(ws_data):
	temp_x_start, temp_x_end = 999999, 999999
	temp_y_start, temp_y_end = 999999, 999999
	temp_ws_data = []
	for y, row in enumerate(ws_data):
		temp_row = []
		temp_empty_row = 0
		for x, col in enumerate(row):
			if "#" == col.value:
				temp_x_start = x
				temp_y_start = y
			if "Judgment" == col.value:
				temp_x_end = x
			if x >= temp_x_start and x <= temp_x_end:
				temp_obj = cell()
				if None == col.value:
					temp_empty_row += 1
					temp_obj.value = str(col.value)
					del temp_obj.first_col, temp_obj.last_col, temp_obj.row
				else:
					temp_obj.first_col = x - temp_x_start
					temp_obj.row = y - temp_y_start
					temp_obj.value = str(col.value)
				temp_row.append(temp_obj)
		if temp_empty_row == (temp_x_end - temp_x_start + 1):
			temp_y_end = y
		if y >= temp_y_start and y < temp_y_end:
			temp_ws_data.append(temp_row)
		elif y >= temp_y_end:
			break
	return tuple(temp_ws_data)

class pcl:
	def __init__(self):
		self.tcfirst_row = -1
		self.tclast_row = -1
		self.init_row = -1
		self.inp_avail = False
		self.inpfirst_col = -1
		self.inplast_col = -1
		self.out_avail = False
		self.outfirst_col = -1
		self.outlast_col = -1
		self.jud_col = -1

def ut_getpclinfo(ws_data):
	pclinfo = pcl()
	for y, row in enumerate(ws_data):
		for x, col in enumerate(row):
			col_val = str(col.value)
			if x == 0:
				if "-" in col_val:
					if pclinfo.tcfirst_row == -1:
						pclinfo.tcfirst_row = y
						pclinfo.init_row = y - 1
			if "Input factor" in col_val:
				pclinfo.inp_avail = True
				pclinfo.inpfirst_col = x
			elif "Output element" in col_val:
				pclinfo.out_avail = True
				pclinfo.outfirst_col = x
				if pclinfo.inp_avail:
					pclinfo.inplast_col = x
			elif "Judgment" in col_val:
				pclinfo.jud_col = x
				if pclinfo.out_avail:
					pclinfo.outlast_col = x
				else:
					pclinfo.inplast_col = x
	pclinfo.tclast_row = y
	return pclinfo

def ut_getlastcol(ws_data):
	temp_ws_data = []
	for y, row in enumerate(ws_data):
		temp_col = 0
		temp_row = []
		for x, col in enumerate(row):
			if col.value != "None":
				for i, temp_temp_row in enumerate(temp_row):
					if i == temp_col:
						temp_temp_row.last_col = x
						break
				temp_col = x
			else:
				if y != 0 or y != 1:
					temp_y = y
					while temp_y > 0:
						temp_upper_value = ws_data[temp_y - 1][x].value
						if temp_upper_value != "None":
							for i, temp_temp_row in enumerate(temp_row):
								if i == temp_col:
									temp_temp_row.last_col = x
									break
							temp_col = x
						temp_y -= 1
			temp_row.append(col)
		temp_ws_data.append(temp_row)
	return tuple(temp_ws_data)

class ut_pcl:
	def __init__(self, ws_data):
		self.data = ut_getlastcol(ut_getdata(ws_data))
		self.pclinfo = ut_getpclinfo(self.data)
		pass
