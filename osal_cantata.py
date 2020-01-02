def export_test_program(pcl):
	def check_type(type):
		if "uint" in type:
			return "CHECK_U_INT"
		elif "int" in type:
			return "CHECK_I_INT"
		elif "bool" in type:
			return "CHECK_BOOLEAN"
		elif "addr_t" in type:
			return "CHECK_ADDRESS"

		return "CHECK_S_INT"

	def element_child_handle(child):
		if child:
			for element in child:
				if element.is_structure and not element.is_pointer and not element.parent:
					tab = "\t\t"
					f_expected_return_data_declarations.write(f"{tab}{element.type} {element.name};\n")
				if not element.child:
					tab = "\t"
					f_data_declarations.write(f"{tab}{element.declare};\n")
					if element.is_pointer:
						tab = "\t\t"
						f_expected_return_data_declarations.write(f"{tab}{element.type} {element.name};\n")
						tab = "\t\t\t"
						f_external_data_initialize.write(f"{tab}if (CURRENT_TEST.{element.name} != NULL) {{\n{tab}\tCURRENT_TEST.{element.name} = &{element.name};\n{tab}}}\n")
					if element.parent:
						tab = "\t\t\t"
						temp = ""
						for data in element.parent:
							temp += f"{data}."
						f_external_data_initialize.write(f"{tab}{temp}{element.name} = CURRENT_TEST.{element.name};\n")
				else:
					element_child_handle(element.child)
				
				for expect in element.check_expected:
					tab = "/t/t"
					temp_declare = expect.replace("[", "_").replace("]", "")
					f_expected_data_declarations.write(f"{tab}{element.type} expected_{temp_declare};\n")
					tab = "/t/t/t"
					check = check_type(element.type)
					f_expected_data_check.write(f"{tab}if (CURRENT_TEST.expected_{temp_declare} != UTS_DONTCARE) {{")
					f_expected_data_check.write(f"{tab}/t{check}({element.name}{expect}, CURRENT_TEST.expected_{temp_declare});\n")
					f_expected_data_check.write(f"{tab}}}\n")

	f_data_declarations = open("data_declarations.txt", "w")
	f_expected_return_data_declarations = open("expected_return_data_declarations.txt", "w")
	f_external_data_initialize = open("external_data_initialize.txt", "w")
	f_expected_data_declarations = open("f_expected_data_declarations.txt", "w")
	f_expected_data_check = open("expected_data_check.txt", "w")
	element_child_handle(pcl.pcl_input_factor)
	f_data_declarations.close()
	f_expected_return_data_declarations.close()
	f_external_data_initialize.close()
	f_expected_data_declarations.close()
	f_expected_data_check.close()
