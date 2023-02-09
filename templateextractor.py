from docx import Document
from docx.shared import Inches, Length
import olefile
import argparse

parser = argparse.ArgumentParser(description='extract template information from a word template')
parser.add_argument('filename',
                    help='the filename of the word doc template')
parser.add_argument('--output',
                    help='the filename to write the output data to (optional)')
args = parser.parse_args()



document = Document(args.filename)

output = ""

def gen_output_line(name, value):
	return str(name) + "=" + str(value) + "\n"

def gen_output_data_line(name, value):
	return gen_output_line(name, round(value, 3))

vertical_tablespace=0

output += gen_output_line("unit", "inches")


sections = document.sections
for section in sections:
	# print(section.start_type)
	# print(section.page_height, section.page_width)
	# section.orientation
	output += gen_output_data_line("page_height", section.page_height.inches)
	output += gen_output_data_line("page_width", section.page_width.inches)
	output += gen_output_data_line("left_margin", section.left_margin.inches)
	output += gen_output_data_line("right_margin", section.right_margin.inches)
	output += gen_output_data_line("top_margin", section.top_margin.inches)
	output += gen_output_data_line("bottom_margin", section.bottom_margin.inches)

	vertical_tablespace = Length(section.page_height - section.top_margin - section.bottom_margin).inches



for table in document.tables:
	output += gen_output_data_line("column_count", len(table.columns))
	output += gen_output_data_line("row_count", len(table.rows))
	# print(table.autofit)
	first_row = table.row_cells(0)
	col_widths = [cell.width.inches for cell in first_row]

	output += gen_output_data_line("col_spacing", min(col_widths))
	output += gen_output_data_line("label_width", max(col_widths))


	# calculate cell height, cuz i guess thats not a thing
	
	# first_col = table.column_cells(0)
	# for cell in first_col:
	# 	print(cell.width.inches)
	# section.orientation
	# print(section.page_height.inches, section.page_width.inches)


# print(olefile.isOleFile(args.filename))
# with  olefile.OleFileIO(args.filename) as ole:
# 	print(ole.listdir())
# 	if ole.exists('worddocument'):
# 		print("This is a Word document.")
# 		doc = ole.openstream('worddocument')
# 		data = doc.read()
# 		# print(data)
# 		print('1Table')
# 		doc = ole.openstream('1Table')
# 		doc.listdir()
# 		data = doc.read()
# 		print(data)
# 		print('Data')
# 		doc = ole.openstream('Data')
# 		data = doc.read()
# 		print(data)
		
# 		if ole.exists('macros/vba'):
# 			print("This document seems to contain VBA macros.")
# 	# meta = ole.get_metadata()
# 	# meta.dump()
	output += gen_output_data_line("label_height_estimate", vertical_tablespace/len(table.rows))


if args.output:
	with open(args.output, "w") as outfile:
		outfile.writelines(output.split("\n"))
else:
	print(output)