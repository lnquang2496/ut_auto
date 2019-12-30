def export_test_program(pcl):
	def element_child_handle(child):
		for element in child:
			print(element.name)
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
					f_external_data_initialize.write(f"{tab}{element.parent}.{element.name} = CURRENT_TEST.{element.name};\n")
			else:
				element_child_handle(element.child)

	f_data_declarations = open("data_declarations.txt", "w")
	f_expected_return_data_declarations = open("expected_return_data_declarations.txt", "w")
	f_external_data_initialize = open("external_data_initialize.txt", "w")
	element_child_handle(pcl.pcl_input_factor)

	f_data_declarations.close()
	f_expected_return_data_declarations.close()
	f_external_data_initialize.close()
