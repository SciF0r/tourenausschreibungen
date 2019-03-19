from tour_parser import TourParser

parser = TourParser('https://www.sac-aarau.ch/aktivitaeten/kalender/', '01.04.2019', '30.06.2019')
parser.parse()
