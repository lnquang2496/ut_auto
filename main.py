from os import system
from os.path import join
from gc import collect
from pandas import read_csv
from openpyxl import load_workbook
from time import time

class cell(object):
	def __init__(self, ws, target, match_case:bool=False):
		self.ws = ws
		self.first_col = 0
		self.first_row = 0
		self.last_col = 0
		self.last_row = 0
		self.value = "None"
		self.merged_cell = False
		self.find_cell(target, match_case)

	def find_cell(self, target, match_case:bool=False):
		def coor_merged_cell(cell):
			for merged_cell in self.ws.merged_cells:
				if (cell.coordinate in merged_cell):
					temp = (merged_cell.bounds[0], merged_cell.bounds[1],
							merged_cell.bounds[2], merged_cell.bounds[3])
					self.merged_cell = True
					break
			else:
				temp = (cell.column, cell.row, cell.column, cell.row)
				self.merged_cell = False
			self.first_col = temp[0]
			self.first_row = temp[1]
			self.last_col = temp[2]
			self.last_row = temp[3]
			self.value = str(cell.value)

		if isinstance(target, str):
			for rows in self.ws:
				for cell in rows:
					cell_value = str(cell.value)
					if ((target == cell_value) if match_case else (target in cell_value)):
						coor_merged_cell(cell)
						return True
		elif isinstance(target, list) or isinstance(target, tuple):
			cell = self.ws.cell(column=target[0], row=target[1])
			coor_merged_cell(cell)
		else:
			print(f"target {target} is not str or list")
		return False

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

	def coor_print(self):
		print("--------------")
		print("first column : ", self.first_col)
		print("first row    : ", self.first_row)
		print("last column  : ", self.last_col)
		print("last row     : ", self.last_row)
		print("cell value   : ", self.value)
		print("merged cell  : ", self.merged_cell)
		print("--------------")

class testcase(object):
	def __init__(self, ws, target):
		self.ws = ws
		self.tc_col = 0
		self.first_tc_row = 0
		self.last_tc_row = 0
		self.total_tc = 0
		self.get_info(target)

	def get_info(self, target):
		coor = cell(self.ws, target)
		self.tc_col = coor.first_col
		coor.coor_shift_down()
		self.first_tc_row = coor.first_row
		while(coor.value != "None"):
			self.total_tc += 1
			coor.coor_shift_down()
		self.last_tc_row = coor.first_row - 1

	def testcase_print(self):
		print("--------------")
		print("tc column    : ", self.tc_col)
		print("first tc row : ", self.first_tc_row)
		print("last tc row  : ", self.last_tc_row)
		print("total tc     : ", self.total_tc)
		print("--------------")

def main():
	def ws_get(wb_dir, wb_name, ws_name):
		wb = load_workbook(join(wb_dir, wb_name), data_only=True)
		return wb[ws_name]
	def extract_data(x):
		return x["implement"], x["workbook_dir"], x["workbook_name"], x["worksheet_name"]

	config_df = read_csv("config.csv", index_col="no")

	i = 0
	while True:
		try:
			imp, wb_dir, wb_name, ws_name = extract_data(config_df.iloc[i])
			print(imp, wb_dir, wb_name, ws_name)
			if ("yes" == imp):
				ws = ws_get(wb_dir, wb_name, ws_name)
				coor_temp = cell(ws, "Input factor")
				tc_obj = testcase(ws, "#")
				tc_obj.testcase_print()
			i += 1
		except:
			break

if __name__ == "__main__":
	start = time()
	main()
	end = time()
	print("execute time : ", (end - start)*1000, "ms")
