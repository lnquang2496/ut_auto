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
		#self.merged_cell = False
		self.find_cell(target, match_case)

	def find_cell(self, target, match_case:bool=False):
		def coor_merged_cell(cell):
			for merged_cell in self.ws.merged_cells:
				if cell.coordinate in merged_cell.coord:
					temp = (merged_cell.bounds[0], merged_cell.bounds[1],
							merged_cell.bounds[2], merged_cell.bounds[3])
					#self.merged_cell = True
					break
			else:
				temp = (cell.column, cell.row, cell.column, cell.row)
				#self.merged_cell = False
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
		elif target.value:
			self.first_col = target.first_col
			self.first_row = target.first_row
			self.last_col = target.last_col
			self.last_row = target.last_row
			self.value = target.value
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
		print(self.first_col, self.first_row, self.last_col, self.last_col, self.value, sep="\t")

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
		return value[value.find("[") : value.find("]") + 1]
	else:
		return None

class cell_attribute:
	pass

def temp(obj):
	declare = obj.value.replace("[a]", "")
	pointer = is_pointer(declare)
	structure = is_structure(declare)
	array = is_array(declare)
	temp = declare
	if pointer:
		temp = declare.replace(pointer, "")
	if structure:
		temp = declare.replace(structure, "")
	if array:
		temp = declare.replace(array, "")
	print(temp)

class inoutput_handler(object):
	def __init__(self, ws):
		self.ws = ws
		self.list_obj = []
		obj_testcase = testcase(self.ws, "#")
		obj_input = cell(self.ws, "Input factor")
		obj_output = cell(self.ws, "Output element")
		self.get_info(obj_input)

	class obj_argument(object):
		def __init__(self, ws, value):
			self.declare = value
			self.type = ""
			self.name = ""
			self.pointer = False
			self.structure = False
			self.init_value = 0
			self.check_expected = False
			self.data_extract(value)

		def data_extract(self, value):
			temp_value = value
			if inoutput_handler.is_pointer(value):
				self.pointer = True
				temp_value = value.replace("*", "")
			self.type, self.name = temp_value.split()

	def get_info(self, obj):
		obj_1 = cell(self.ws, obj)
		obj_1.coor_shift_down()
		while obj_1.last_col <= obj.last_col:
			if "[a]" in obj_1.value:
				temp_value = obj_1.value.replace("[a]", "")
				obj_cell = self.obj_argument(self.ws, temp_value)
				obj_1.coor_print()
			obj_1.coor_shift_right()
		"""
		def inoutput_search(obj, row_testcase):
			obj_1 = cell(self.ws, obj)
			obj_1.coor_shift_down()
			while obj_1.last_col <= obj.last_col:
				process = True
				if is_pointer(obj_1.value):
					pass
				else:
					if is_structure(obj_1.value):
						inoutput_search(obj_1, row_testcase)
						process = False
				if process:
					obj_1.coor_print()
					print("\t\t\t\tvalue:\t", str(self.ws.cell(row_testcase, obj_1.first_col).value), sep="")
				obj_1.coor_shift_right()
			else:
				del obj_1

		obj_testcase = testcase(self.ws, "#")
		obj_input = cell(self.ws, "Input factor")
		for r in range(obj_testcase.first_tc_row, obj_testcase.last_tc_row + 1):
			inoutput_search(obj_input, r)
		obj_output = cell(self.ws, "Output element")
		for r in range(obj_testcase.first_tc_row, obj_testcase.last_tc_row + 1):
			inoutput_search(obj_output, r)
		"""

		"""
		def input_search(obj, row):
			def structure_search(obj, row):
				obj_temp = cell(self.ws, obj)
				obj_temp.coor_shift_down()
				while obj_temp.last_col <= obj.last_col:
					if is_pointer(obj_temp.value):
						pass
					else:
						if is_structure(obj_temp.value):
							structure_search(obj_temp, row)
					obj_temp.coor_shift_right()
				else:
					del obj_temp

			obj_temp = cell(self.ws, obj)
			obj_temp.coor_shift_down()
			while obj_temp.last_col <= obj.last_col:
				if "[a]" in obj_temp.value:
					if is_pointer(obj_temp.value):
						pass
					else:
						if is_structure(obj_temp.value):
							structure_search(obj_temp, row)
				elif "[rt]" in obj_temp.value:
					pass
				elif "[f]" in obj_temp.value:
					pass
				obj_temp.coor_shift_right()
			else:
				del obj_temp

		obj_testcase = testcase(self.ws, "#")
		obj_input = cell(self.ws, "Input factor")
		for r in range(obj_testcase.first_tc_row, obj_testcase.last_tc_row + 1):
			input_search(obj_input, r)
		"""


def main():
	def ws_get(wb_dir, wb_name, ws_name):
		wb = load_workbook(join(wb_dir, wb_name), data_only=True)
		return wb[ws_name]
	def extract_data(x):
		return x["implement"], x["workbook_dir"], x["workbook_name"], x["worksheet_name"]

	config_df = read_csv("data/config.csv", index_col="no")

	for i in range(config_df._AXIS_LEN):
		imp, wb_dir, wb_name, ws_name = extract_data(config_df.iloc[i])
		print(imp, wb_dir, wb_name, ws_name, sep="\t")
		if ("yes" == imp):
			ws = ws_get(wb_dir, wb_name, ws_name)
			#obj = inoutput_handler(ws)
			obj_input = cell(ws, "Input factor")
			obj_1 = cell(ws, obj_input)
			obj_1.coor_shift_down()
			while obj_1.last_col <= obj_input.last_col:
				if ("[a]" in obj_1.value):
					temp(obj_1)
				obj_1.coor_shift_right()

if __name__ == "__main__":
	start = time()
	main()
	end = time()
	print("execute time : ", (end - start)*1000, "ms")
