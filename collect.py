import sys
from datetime import datetime
from docx_creator import DocxCreator
from tour_cleaner import TourCleaner
from tour_parser import TourParser

if len(sys.argv) != 2:
    print('Usage: {0} path_to_xlsx_file'.format(sys.argv[0]))
    sys.exit(1)

file_path = sys.argv[1]

print('Getting tours from {0}'.format(file_path))
parser = TourParser(file_path)
tours = parser.parse_for_year_program()

cleaner = TourCleaner(tours)
with open('rules.txt', 'r') as rules_file:
    rules = [line.split('/', 2) for line in rules_file.read().splitlines()]
cleaned_tours = cleaner.clean(rules)

export_file_name = '{0}_jahresprogramm_{1}.docx'.format(datetime.now().strftime('%Y-%m-%d_%H_%M_%S'), '{0}')
docx_creator = DocxCreator(cleaned_tours, export_file_name)
docx_creator.create()
