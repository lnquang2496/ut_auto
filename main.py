from openpyxl import load_workbook
from time import time
from ut_pcl import ut_pcl

def main():
	wb = load_workbook("d:/ut_auto/data/test.xlsx", True, False, True, False)
	ws = wb["Sheet1"]
	# for i in range(1000):
	# 	obj_ut_pcl = ut_pcl(ws.rows)
	obj_ut_pcl = ut_pcl(ws.rows)
	pass

if __name__ == "__main__":
	start = time()
	main()
	end = time()
	print("Program finish within:", (end - start)*1000, "ms")
