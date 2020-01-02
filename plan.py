from openpyxl import load_workbook
from time import time
import osal_cantata

class pcl:
	def __init__(self, wb_name, ws_name):
		def get_pcl_from_worksheet(rows):
			class cell:
				def __init__(self):
					self.col = -1
					self.col_last = -1
					self.row = -1
					self.value = None
				pass
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
						temp_obj = cell()
						temp_obj.col = x - temp_x_start
						temp_obj.row = y - temp_y_start
						temp_obj.value = str(col.value)

						temp_row.append(temp_obj)

						if None == col.value:
							temp_empty_row += 1
					
				if temp_empty_row == (temp_x_end - temp_x_start + 1):
					temp_y_end = y
				if y >= temp_y_start and y < temp_y_end:
					temp_rows.append(temp_row)
				elif y >= temp_y_end:
					break
			return tuple(temp_rows)

		def get_pcl_info(rows):
			class pcl_info:
				def __init__(self):
					self.tcfirst_row = -1
					self.tclast_row = -1
					self.initval_row = -1
					self.imp_available = False
					self.imp_first_col = -1
					self.imp_last_col = -1
					self.out_available = False
					self.out_first_col = -1
					self.out_last_col = -1
					self.judgment_col = -1
				pass

			obj_pcl_info = pcl_info()
			for y, row in enumerate(rows):
				for x, col in enumerate(row):
					if x == 0:
						if "-" in col.value:
							if obj_pcl_info.tcfirst_row == -1:
								obj_pcl_info.tcfirst_row = y
								obj_pcl_info.initval_row = y - 1
					if "Input factor" in col.value:
						obj_pcl_info.imp_available = True
						obj_pcl_info.imp_first_col = x
					if "Output element" in col.value:
						obj_pcl_info.out_available = True
						obj_pcl_info.out_first_col = x
						if obj_pcl_info.imp_available:
							obj_pcl_info.imp_last_col = x
					if "Judgment" in col.value:
						obj_pcl_info.judgment_col = x
						if obj_pcl_info.out_available:
							obj_pcl_info.out_last_col = x
						else:
							obj_pcl_info.imp_last_col = x
			obj_pcl_info.tclast_row = y
			return obj_pcl_info

		def get_last_col(rows):
			temp_rows = []
			for y, row in enumerate(rows):
				temp_col = 0
				temp_row = []
				for x, col in enumerate(row):
					if col.value != "None":
						temp_value = col.value
						for i, temp_temp_row in enumerate(temp_row):
							if i == temp_col:
								temp_temp_row.col_last = x
								break
						temp_col = x
					else:
						if y != 0 or y != 1:
							temp_y = y
							while temp_y > 0:
								temp_upper_value = rows[temp_y - 1][x].value
								if temp_upper_value != "None":
									for i, temp_temp_row in enumerate(temp_row):
										if i == temp_col:
											temp_temp_row.col_last = x
											break
									temp_col = x
								temp_y -= 1
							
					temp_row.append(col)
				temp_rows.append(temp_row)
			return temp_rows

		wb = load_workbook(wb_name, True, False, True, False)
		self.rows = get_last_col(get_pcl_from_worksheet(tuple(wb[ws_name].rows)))
		self.pcl_info = get_pcl_info(self.rows)
		self.pcl_input_factor = self.input_factor_handle()

	def input_factor_handle(self):
		class element_info:
			def __init__(self):
				self.cell_info = None
				self.declare = None
				self.type = None
				self.name = None
				self.is_pointer = False
				self.is_structure = False
				self.is_array = False
				self.parent = []
				self.child = None
				self.check_expected = []

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

		def get_element_info(initval_row, current_cell):
			def get_variable_info(current_cell, prefix="", parent:list=[]):
				temp_cell_value = current_cell.value.replace(prefix, "")
				obj_element_info = element_info()
				obj_element_info.cell_info = current_cell
				obj_element_info.declare = temp_cell_value

				pointer_found = is_pointer(temp_cell_value)
				structure_found = is_structure(temp_cell_value)
				array_found = is_array(temp_cell_value)

				if pointer_found:
					obj_element_info.is_pointer = True
					if "*" == pointer_found:
						temp_cell_value = temp_cell_value.replace("*", "")
				if structure_found:
					obj_element_info.is_structure = True

				obj_element_info.name = temp_cell_value[temp_cell_value.rfind(" ") + 1 :].replace(" ", "")
				obj_element_info.type = temp_cell_value[: temp_cell_value.rfind(" ")]
				obj_element_info.parent = parent

				none_count = 0
				temp_variable_info = []
				for col in range(current_cell.col, current_cell.col_last):
					target_row = current_cell.row + 1
					if target_row != initval_row:
						temp_cell = self.rows[target_row][col]
						if temp_cell.value != "None":
							if not obj_element_info.is_pointer:
								temp = []
								for data in obj_element_info.parent:
									temp.append(data)
								temp.append(obj_element_info.name)
								temp_parent = temp
							else:
								temp_parent = [obj_element_info.name]
							temp_variable_info.append(get_variable_info(temp_cell, parent=temp_parent))
						else:
							none_count += 1
					else:
						obj_element_info.child = None
						break
				else:
					obj_element_info.child = temp_variable_info

				if none_count == (current_cell.col_last - current_cell.col):
					obj_element_info.child = None

				return obj_element_info

			temp_prefix = "[a]" if "[a]" in current_cell.value else None
			if not temp_prefix:
				temp_prefix = "[g]" if "[g]" in current_cell.value else None
			if temp_prefix:
				obj_element_info = get_variable_info(current_cell, temp_prefix)
				return obj_element_info
			return None

		def get_extra_variable_info(current_cell, prefix, main_element_info:list=[]):
			temp_cell_value = current_cell.value.replace(prefix, "")
			temp_cell_value = temp_cell_value.replace("*", "") if "*" in temp_cell_value else temp_cell_value
			temp_find_blank = temp_cell_value.rfind(" ")
			while temp_find_blank == len(temp_cell_value) - 1:
				temp_cell_value = temp_cell_value[:-1]
				temp_find_blank = temp_cell_value.rfind(" ")
			if temp_find_blank != -1:
				temp_cell_value = temp_cell_value[temp_find_blank + 1 :].replace(" ", "")
			if main_element_info:
				for temp_element_info in main_element_info:
					if temp_element_info.name == temp_cell_value:
						for col in range (current_cell.col, current_cell.col_last):	
							temp_extra_cell_value = self.rows[current_cell.row + 1][col].value
							temp_element_info.check_expected.append(str(f"{temp_element_info.name}{temp_extra_cell_value}"))
			return main_element_info

		def get_extra_element_info(initval_row, current_cell, main_element_info:list=[]):
			temp_prefix = "[a]" if "[a]" in current_cell.value else None
			if not temp_prefix:
				temp_prefix = "[g]" if "[g]" in current_cell.value else None
			if temp_prefix:
				obj_element_info = get_extra_variable_info(current_cell, temp_prefix, main_element_info)
				return obj_element_info
			return main_element_info

		if self.pcl_info.imp_available:
			main_element_info = []
			# Input factor range handle
			for col in range(self.pcl_info.imp_first_col, self.pcl_info.imp_last_col):
				current_cell = self.rows[1][col]
				temp_element_info = get_element_info(self.pcl_info.initval_row, current_cell)
				if temp_element_info:
					main_element_info.append(temp_element_info)
			# Output element range handle
			for col in range(self.pcl_info.out_first_col, self.pcl_info.out_last_col):
				current_cell = self.rows[1][col]
				main_element_info = get_extra_element_info(self.pcl_info.initval_row, current_cell, main_element_info)
			return main_element_info

start = time()
obj_pcl = pcl("d:/ut_auto/data/test.xlsx", "Sheet1")
osal_cantata.export_test_program(obj_pcl)
end = time()
print((end - start) * 1000, "ms")
