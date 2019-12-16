from os import system
from os.path import join
from gc import collect
from pandas import read_csv
from openpyxl import load_workbook

class coor_handler(object):
	def __init__(self, ws, target, match_case:bool=False):
		self.ws = ws
		self.IsMergedCell = False
		coor_temp = self.coor_get(ws, target, match_case)
		if coor_temp != None:
			self.CoorAvailable = True
			self.assign_coor(coor_temp)
		else:
			self.CoorAvailable = False
		self.Value = self.get_cell_value(ws)
		print(self.Value, type(self.Value))

	def assign_coor(self, coor):
		self.FirstColumn = coor[0]
		self.FirstRow = coor[1]
		self.LastColumn = coor[2]
		self.LastRow = coor[3]

	def coor_get(self, target, match_case:bool=False):
		def coor_merged_cell(cell):
			for merged_cell in self.ws.merged_cells:
				if (cell.coordinate in merged_cell):
					temp = (merged_cell.bounds[0], merged_cell.bounds[1],
							merged_cell.bounds[2], merged_cell.bounds[3])
					self.IsMergedCell = True
					return temp
			else:
				temp = (cell.column, cell.row, cell.column, cell.row)
				self.IsMergedCell = False
				return temp

		if isinstance(target, str):
			for rows in self.ws:
				for cell in rows:
					cell_value = str(cell.value)
					if ((target == cell_value) if match_case else (target in cell_value)):
						return coor_merged_cell(cell)
		elif isinstance(target, list):
			cell = self.ws.cell(column=target[0], row=target[1])
			return coor_merged_cell(cell)
		else:
			print(f"target {target} is not str or list")
		return None

	def get_cell_value(self)->str:
		return str(self.ws.cell(column=self.FirstColumn, row=self.FirstRow).value)

	def coor_shift_up(self):
		temp = self.FirstRow - 1
		temp = [self.FirstColumn, 1 if temp < 1 else temp]
		temp = self.coor_get(temp)
		self.assign_coor(temp)

	def coor_shift_down(self):
		temp = [self.FirstColumn, self.LastRow + 1]
		temp = self.coor_get(temp)
		self.assign_coor(temp)

	def coor_shift_left(self):
		temp = self.FirstColumn - 1
		temp = [1 if temp < 1 else temp, self.FirstRow]
		temp = self.coor_get(temp)
		self.assign_coor(temp)

	def coor_shift_right(self):
		temp = [self.LastColumn + 1, self.FirstRow]
		temp = self.coor_get(temp)
		self.assign_coor(temp)

def main():
	def ws_get(wb_dir, wb_name, ws_name):
		wb = load_workbook(join(wb_dir, wb_name), data_only=True)
		return wb[ws_name]
	def extract_data(x):
		return x["implement"], x["workbook_dir"], x["workbook_name"], x["worksheet_name"]

	config_df = read_csv("config.csv", index_col="no")
	#print(config_df)

	i = 0
	while True:
		try:
			imp, wb_dir, wb_name, ws_name = extract_data(config_df.iloc[i])
			print(imp, wb_dir, wb_name, ws_name)
			if ("yes" == imp):
				ws = ws_get(wb_dir, wb_name, ws_name)
				coor_temp = coor_handler(ws, "Input factor")
				if (coor_temp.CoorAvailable):
					print(coor_temp.FirstColumn, coor_temp.FirstRow)
					print(coor_temp.LastColumn, coor_temp.LastRow)
				else:
					print("coor not available")
				coor_temp = coor_handler(ws, [1, 1])
				print(coor_temp.FirstColumn, coor_temp.FirstRow)
				print(coor_temp.LastColumn, coor_temp.LastRow)
			i += 1
		except:
			break

if __name__ == "__main__":
	main()
