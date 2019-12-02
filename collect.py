import sys
from datetime import datetime
from docx_creator import DocxCreator
from tour_cleaner import TourCleaner
from tour_parser import TourParser

if len(sys.argv) != 5:
    print('Usage: {0} start_date end_date username password (date format: dd.mm.YYYY)'.format(sys.argv[0]))
    sys.exit(1)

start_date = sys.argv[1]
end_date = sys.argv[2]
username = sys.argv[3]
password = sys.argv[4]
url = 'https://ssl.dropnet.ch/sac-aarau/aktivitaeten/kalender/'

print('Getting tours from {0}, {1} to {2}'.format(url, start_date, end_date))
parser = TourParser(url, start_date, end_date, username, password)
tours = parser.parse()

cleaner = TourCleaner(tours)
with open('rules.txt', 'r') as rules_file:
    rules = [line.split('/', 2) for line in rules_file.read().splitlines()]
cleaned_tours = cleaner.clean(rules)

export_file_name = '{0}_tourenausschreibungen_{1}_{2}.docx'.format(
        datetime.now().strftime('%Y-%m-%d_%H:%M:%S'),
        start_date,
        end_date)
docx_creator = DocxCreator(cleaned_tours, export_file_name)
docx_creator.create()
