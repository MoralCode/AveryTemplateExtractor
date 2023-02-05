from docx import Document
from docx.shared import Inches
import argparse
import argparse

parser = argparse.ArgumentParser(description='extract template information from a word template')
parser.add_argument('filename',
                    help='the filename of the word doc template')
args = parser.parse_args()



document = Document(args.filename)