import locale
import pandas as pd
from docx import Document
from tour_parser import TourParser

class DocxCreatorRoteKarte:
    """Create a docx file from given tours"""
    __template = 'tourenausschreibungen_template.docx'
    __style = 'Standard_Ausschreibungen'
    __tours_events = []
    __tours_sektion = []
    __tours_fabe = []
    __tours_kibe = []
    __tours_jo = []

    def __init__(self, tours, file_name):
        self.__file_name = file_name
        self.__tours = tours
        self.__document = Document(self.__template)
        locale.setlocale(locale.LC_ALL, 'de_CH')

    def create(self):
        for row_count, row in self.__tours.iterrows():
            groups = row[TourParser.COL_GROUP].split('|')
            if 'Events' in groups:
                self.__add_tour(row, self.__tours_events)
            if 'Sektion' in groups:
                self.__add_tour(row, self.__tours_sektion)
            if 'Familienbergsteigen' in groups:
                self.__add_tour(row, self.__tours_fabe)
            if 'Kinderbergsteigen' in groups:
                self.__add_tour(row, self.__tours_kibe)
            if 'Jugendorganisation' in groups:
                self.__add_tour(row, self.__tours_jo)
        self.__write_document()

    def __add_tour(self, row, tours_group):
        tour = []
        self.__add_activity(row, tour)
        self.__add_guide(row, tour)
        self.__add_requirements(row, tour)
        self.__add_metrics(row, tour)
        self.__add_route(row, tour)
        self.__add_hospitality(row, tour)
        self.__add_costs(row, tour)
        self.__add_execution(row, tour)
        self.__add_description(row, tour)
        self.__add_additional_information(row, tour)
        self.__add_gear(row, tour)
        self.__add_registration(row, tour)
        tours_group.append(tour)

    def __add_activity(self, row, tour):
        tour.append((self.__get_date(row), row[TourParser.COL_ACTIVITY]))

    def __add_guide(self, row, tour):
        tour.append((row[TourParser.COL_TOUR_TYPE_LONG], row[TourParser.COL_GUIDE]))

    def __add_requirements(self, row, tour):
        requirements = self.__get_requirements(row)
        if requirements:
            tour.append('Anforderungen', requirements)

    def __add_metrics(self, row, tour):
        self.__append_if_not_empty(row, tour, TourParser.COL_METRICS, 'Auf-/Abstieg, MZ')

    def __add_route(self, row, tour):
        self.__append_if_not_empty(row, tour, TourParser.COL_ROUTE, 'Reiseroute')

    def __add_hospitality(self, row, tour):
        self.__append_if_not_empty(row, tour, TourParser.COL_HOSPITALITY, 'Unterk./Verpfl.')

    def __add_costs(self, row, tour):
        costs = self.__get_costs(row)
        if costs:
            tour.append('Kosten', costs)

    def __add_execution(self, row, tour):
        self.__append_if_not_empty(row, tour, TourParser.COL_EXECUTION, 'Durchführung')

    def __add_description(self, row, tour):
        self.__append_if_not_empty(row, tour, TourParser.COL_ROUTE_DESCRIPTION, 'Route / Details')

    def __add_additional_information(self, row, tour):
        self.__append_if_not_empty(row, tour, TourParser.COL_ADDITIONAL_INFORMATION, 'Zusatzinfo')

    def __add_gear(self, row, tour):
        self.__append_if_not_empty(row, tour, TourParser.COL_GEAR, 'Ausrüstung')

    def __add_registration(self, row, tour):
        registration = self.__get_registration(row)
        if registration:
            tour.append('Anmeldung', registration)

    def __get_date(self, row):
        start_date = row[TourParser.COL_START_DATE]
        end_date = row[TourParser.COL_END_DATE]
        if pd.isna(end_date):
            return start_date.strftime('%d.%m.%Y')
        else:
            return '{0}-{1}'.format(start_date.strftime('%d'), end_date.strftime('%d.%m.%y'))

    def __get_requirements(self, row):
        conditions = []
        cond_req = row[TourParser.COL_COND_REQ]
        if not pd.isna(cond_req):
            conditions.append(cond_req)
        cond_tech = row[TourParser.COL_TECH_REQ]
        if not pd.isna(cond_tech):
            conditions.append(cond_tech)
        return ', '.join(conditions)

    def __get_costs(self, row):
        costs = row[TourParser.COL_COSTS]
        costs_long = row[TourParser.COL_COSTS_LONG]
        costs_string = ''
        if not pd.isna(costs):
            costs_string = 'Fr. {0}.--'

    def __append_if_not_empty(self, row, tour, field, key):
        value = row[field]
        if value:
            tour.append(key, value)

    def __write_document(self):
        document.add_paragraph(line, self.__style)
        self.__document.save(self.__file_name)
