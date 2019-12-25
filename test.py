from openpyxl import load_workbook
from os.path import join
from time import time

xxx = time()

def get_worksheet_data(wb_dir, wb_name, ws_name):
	wb = load_workbook(join(wb_dir, wb_name), read_only=True, data_only=True)
	ws_data = wb[ws_name].rows
	flag = False
	ws_data_new = []
	for rows in ws_data:
		for cell in rows:
			if ("#" == cell.value):
				flag = True
		if flag:
			ws_data_new.append(rows)

	wb = load_workbook(join(wb_dir, wb_name), data_only=True)
	return tuple(ws_data_new), tuple(wb[ws_name].merged_cells)

ws_data, ws_mece = get_worksheet_data("d:/ut_auto/data", "test.xlsx", "Sheet2")

class current_cell():
	def __init__(self, target):
		self.first_col = 0
		self.first_row = 0
		self.last_col = 0
		self.last_row = 0
		self.value = None
		self.find_cell(target)
	def find_cell(self, target, match_case:bool=False):
		def mergecell_search(cell):
			for merged_cell in ws_mece:
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
			for rows in ws_data:
				for cell in rows:
					if ((target == str(cell.value)) if match_case else (target in str(cell.value))):
						mergecell_search(cell)
		elif isinstance(target, list) or isinstance(target, tuple):
			for rows in ws_data:
				for cell in rows:
					if hasattr(cell, "column"):
						if cell.column == target[0] and cell.row == target[1]:
							mergecell_search(cell)
		elif hasattr(target, "value"):
			self.first_col = target.first_col
			self.first_row = target.first_row
			self.last_col = target.last_col
			self.last_row = target.last_row
			self.value = target.value

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

obj = current_cell("a")
obj_1 = current_cell(obj)

xxx = (time() - xxx) * 1000
print(xxx, "ms")
