def export_test_program(pcl):
	f_data_declarations = open("data_declarations.txt", "w")
	f_expected_return_data_declarations = open("expected_return_data_declarations.txt", "w")
	f_external_data_initialize = open("external_data_initialize.txt", "w")
	for element in pcl.pcl_input_factor:
		if not element.parent and not element.child:
			f_data_declarations.write(f"{element.declare};\n")
			if element.is_pointer:
				f_expected_return_data_declarations.write(f"{element.type} {element.name};")

	f_data_declarations.close()
	f_expected_return_data_declarations.close()
	f_external_data_initialize.close()
