# v1.0 - 27/08/2020 - First draft regular buckets
# v2.0 - 28/08/2020 - Include not mapped logic + fix residuals
# v3.0 - 29/08/2020 - Fetch dynamically input files as args
##############################################################################
# Import required libraries
import sys, getopt
import os, time
import lxml.etree as xml_lib
import pandas as pd_lib
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-xls_in", help="contains the path to the excel file.")
parser.add_argument("-xml_in", help="contains the path to the original XML file.")
parser.add_argument("-xml_out", help="contains the path to the target XML file.")

args = parser.parse_args()

if args.xls_in is not None and args.xml_in is not None and args.xml_out is not None:
    print('I am using the following Excel {}'.format(args.xls_in))
    print('I am using the following XML as origin {}'.format(args.xml_in))
    print('I am writing the results in the following XML {}'.format(args.xml_out))
    xls_in=format(args.xls_in)
    xml_in=format(args.xml_in)
    xml_out=format(args.xml_out)
else:
    # Prompt user to input path
    print('At least one of the required args (xls_in, xml_in, xml_out) is missing')
    sys.exit(1)

#xls_in = "StaticData_Update_Draft.xlsx"
#xml_in = "bim_matrices_1.6.0.xml"

# Init output file
#xml_out = "output"

# Change me to input excel file location
# Reading excel file
xls_path = os.path.abspath(xls_in + ".xlsx")
xls = pd_lib.ExcelFile(xls_path)

# Loading the ************ALL*********** tab in memory
ALL_tab = pd_lib.read_excel(xls, 'ALL')

# Change me to input xml file location
# Reading xml file
xml_path = os.path.abspath(xml_in + ".xml")
xml_file = xml_lib.parse(xml_path)

#Init not mapped lists
worst_value_X = {}
worst_value_Y = {}

for line in ALL_tab.values:
	matrix_name = line[0]
	x_axis = line[1]
	y_axis = line[2]
	z_axis = line[3]
	value = str(line[4])

	# Find the matrix within the xml file
	matrix = xml_file.find('.//matrix[@name="{}"]'.format(matrix_name))

	# init worst value for every axis
	# check on z_axis not required
	if (matrix_name, x_axis) not in worst_value_X.keys():
		worst_value_X[(matrix_name, x_axis)] = '0'
	if (matrix_name, y_axis) not in worst_value_Y.keys():
		worst_value_Y[(matrix_name, y_axis)] = '0'
		# Init Boolean for ResidualCheck - Oherwise not mapped values will take the maximum in all cases
		ResidualChecked = 0


	# Map the corresponding line
	# Check whether we need 3 dimensions or not
	# if z_axis is None : ---  Not working because the blank cell is taken as "nan"
	if pd_lib.isna(z_axis):
		cell = matrix.find('.//cells/cell[@x_axis="{}"][@y_axis="{}"]'.format(x_axis, y_axis))
	else:
		cell = matrix.find('.//cells/cell[@x_axis="{}"][@y_axis="{}"][@z_axis="{}"]'.format(x_axis, y_axis, z_axis))

	if cell is not None:
		print('Updating matrix: {} | x_axis:{} | y_axis:{} |z_axis: {} with value {}'.format(matrix_name, x_axis, y_axis, z_axis, value))
		cell.attrib['value'] = value.encode("utf-8").decode("utf-8")
	if float(value) > float(worst_value_X[(matrix_name, x_axis)]):
		worst_value_X[(matrix_name, x_axis)]=float(value)
	if x_axis == "Residual":
		# Pay attention to the check on x_axis / impact on y_axis, otherwise, will get "Residual" and hence lost for final check
		worst_value_Y[(matrix_name, y_axis)] = float(value)
		ResidualChecked = 1
	if float(value) > float(worst_value_Y[(matrix_name, y_axis)]) and ResidualChecked == 0:
		worst_value_Y[(matrix_name, y_axis)]=float(value)

# Here we take care of the 'Not Mapped' values
for line in ALL_tab.values:
	matrix_name = line[0]
	x_axis = line[1]
	y_axis = line[2]

	matrix = xml_file.find('.//matrix[@name="{}"]'.format(matrix_name))
	cell_x = matrix.find('.//cells/cell[@x_axis="{}"][@y_axis="{}"]'.format(x_axis, 'Not mapped'))
	cell_y = matrix.find('.//cells/cell[@x_axis="{}"][@y_axis="{}"]'.format('Not mapped', y_axis))

	# Test confirms 'Not mapped' exist in matrix in specified dimension
	if cell_x is not None:
		# Pay attention to the logic = using the transpose to match the worst (or residual)
		#print(matrix_name, x_axis, worst_value_Y[(matrix_name, x_axis)])
		cell_x.attrib['value'] = str(worst_value_X[(matrix_name, x_axis)])
	if cell_y is not None:
		#print(matrix_name, y_axis, worst_value_X[(matrix_name, y_axis)])
		# Pay attention to the logic = using the transpose to match the worst (or residual)
		#worst_value_Y[(matrix_name, x_axis)]
		cell_y.attrib['value'] = str(worst_value_Y[(matrix_name, y_axis)])


# Change me to output xml file location
# The parameter is to get the first line of the xml, otherwise lost
output_path = os.path.abspath(xml_out)
timestamp = time.strftime("%Y%m%d_%H%M%S")
xml_file.write(output_path + "_" + timestamp + ".xml", xml_declaration=True)
#print(worst_value_X)
#print("-------------")
#print(worst_value_Y)
