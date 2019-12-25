from openpyxl import load_workbook

class ws_info:
	def __init__(self, wb_name, ws_name):
		self.rows, self.merged_cells = self.get_worksheet_data(wb_name, ws_name)

	def get_worksheet_data(self, wb_name, ws_name):
		wb = load_workbook(wb_name, data_only=True)
		ws_data = wb[ws_name].rows
		flag = False
		ws_data_new = []
		for rows in ws_data:
			count = 0
			for cell in rows:
				if ("#" == cell.value):
					flag = True
				elif None == cell.value:
					count += 1
			if flag:
				ws_data_new.append(rows)
			if count == len(rows):
				flag = False
		return tuple(ws_data_new), tuple(wb[ws_name].merged_cells)

class current_cell:
	def __init__(self, ws, target):
		self.ws = ws
		self.value = None
		self.find_cell(target)

	def find_cell(self, target, match_case:bool=False):
		def mergecell_search(cell):
			for merged_cell in self.ws.merged_cells:
				if cell.coordinate in merged_cell.coord:
					temp = (merged_cell.bounds[0], merged_cell.bounds[1],
							merged_cell.bounds[2], merged_cell.bounds[3])
					break
			else:
				temp = (cell.column, cell.row, cell.column, cell.row)
			self.first_col = temp[0]
			self.first_row = temp[1]
			self.last_col = temp[2]
			self.last_row = temp[3]
			self.value = str(cell.value)

		if isinstance(target, str):
			for rows in self.ws.rows:
				for cell in rows:
					if ((target == str(cell.value)) if match_case else (target in str(cell.value))):
						mergecell_search(cell)
						return
		elif isinstance(target, list) or isinstance(target, tuple):
			for rows in self.ws.rows:
				for cell in rows:
					if cell.column == target[0] and cell.row == target[1]:
						mergecell_search(cell)
						return
		elif hasattr(target, "value"):
			self.first_col = target.first_col
			self.first_row = target.first_row
			self.last_col = target.last_col
			self.last_row = target.last_row
			self.value = target.value
			return
		del self.first_col, self.first_row, self.last_col, self.last_row
		self.value = None

	def coor_shift_up(self):
		temp = self.first_row - 1
		temp = (self.first_col, 1 if temp < 1 else temp)
		self.find_cell(temp)

	def coor_shift_down(self):
		temp = (self.first_col, self.last_row + 1)
		self.find_cell(temp)

	def coor_shift_left(self):
		temp = self.first_col - 1
		temp = (1 if temp < 1 else temp, self.first_row)
		self.find_cell(temp)

	def coor_shift_right(self):
		temp = (self.last_col + 1, self.first_row)
		self.find_cell(temp)

class tc_info:
	def __init__(self, ws, target):
		self.ws = ws
		self.tc_col = 0
		self.first_tc_row = 0
		self.last_tc_row = 0
		self.total_tc = 0
		self.get_info(target)

	def get_info(self, target):
		coor = current_cell(self.ws, target)
		self.tc_col = coor.first_col
		coor.coor_shift_down()
		self.first_tc_row = coor.first_row
		while(coor.value != "None"):
			coor.coor_shift_down()
		self.last_tc_row = coor.first_row - 1
		self.total_tc = self.last_tc_row - self.first_tc_row + 1

def main():
	ws = ws_info("d:/ut_auto/data/test.xlsx", "Sheet1")
	"""
	for rows in ws.rows:
		for cell in rows:
			print(cell.value, end=", ")
		print()
	"""
	tc = tc_info(ws, "#")
	ce = current_cell(ws, "#")
	print()

if __name__=="__main__":
	main()
