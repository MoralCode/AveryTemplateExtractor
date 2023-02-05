from docx import Document
from docx.shared import Inches
import olefile
import argparse

parser = argparse.ArgumentParser(description='extract template information from a word template')
parser.add_argument('filename',
                    help='the filename of the word doc template')
args = parser.parse_args()



document = Document(args.filename)
print(document)


sections = document.sections
for section in sections:
	print(section.start_type)
	print(section.page_height, section.page_width)
	# section.orientation
	print(section.page_height.inches, section.page_width.inches)


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