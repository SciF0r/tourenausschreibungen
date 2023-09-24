import sys
from datetime import datetime
from docx_creator_addresses import DocxCreator
from address_parser import AddressParser

if len(sys.argv) != 2:
    print('Usage: {0} path_to_xlsx_file'.format(sys.argv[0]))
    sys.exit(1)

file_path = sys.argv[1]

print('Getting addresses from {0}'.format(file_path))
parser = AddressParser(file_path)
addresses = parser.parse_for_year_program()

export_file_name = '{0}_adressen.docx'.format(datetime.now().strftime('%Y-%m-%d_%H_%M_%S'))
docx_creator = DocxCreator(addresses, export_file_name)
docx_creator.create()
