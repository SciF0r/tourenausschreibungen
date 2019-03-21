from datetime import datetime
from docx_creator import DocxCreator
from tour_parser import TourParser

start_date = '01.04.2019'
end_date = '30.06.2019'

parser = TourParser('https://www.sac-aarau.ch/aktivitaeten/kalender/', start_date, end_date)
tours = parser.parse()

export_file_name = '{0}_tourenausschreibungen_{1}_{2}.docx'.format(
        datetime.now().strftime('%Y-%m-%d_%H:%M:%S'),
        start_date,
        end_date)
docx_creator = DocxCreator(tours, export_file_name)
docx_creator.create()
