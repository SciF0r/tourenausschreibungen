import pandas as pd

class TourParser(object):
    """Parse tours retrieved from a DropTours xlsx export file"""
    COL_GROUP = 'Gruppe'
    COL_START_DATE = 'Startdatum'
    COL_END_DATE = 'Enddatum'
    COL_DURATION = 'Dauer'
    COL_TOUR_TYPE = 'Tourtyp'
    COL_TOUR_TYPE_LONG = 'Tourtyp lang'
    COL_GUIDE = 'Tourenleiter 1'
    COL_METRICS = 'Auf-, Abstieg/Marschzeit'
    COL_ROUTE = 'Reiseroute'
    COL_HOSPITALITY = 'Unterkunft / Verpflegung'
    COL_COSTS = 'Kosten'
    COL_COSTS_LONG = 'Kosten lang'
    COL_EXECUTION = 'Durchführungskontakt'
    COL_MEETING_POINT = 'Treffpunkt'
    COL_DESCRIPTION = 'Routenbeschreibung'
    COL_ADDITIONAL_INFORMATION = 'Zusatzinfo'
    COL_GEAR = 'Ausrüstung'
    COL_REGISTRATION_MEANS = 'Anmeldung'
    COL_REGISTRATION_START_DATE = 'Anmeldestart'
    COL_REGISTRATION_END_DATE = 'Anmeldeschluss'
    COL_COND_REQ = 'Kond. Anforderungen'
    COL_TECH_REQ = 'Techn. Anforderungen'
    COL_ACTIVITY = 'Aktivität'
    COL_FIRST_NAME = 'Vorname'
    COL_LAST_NAME = 'Name'
    COL_STATUS = 'Status'
    COL_PROCESS_STATUS = 'Prozessstatus'
    STATUS_CANCELLED = 'abgesagt'
    STATUS_FULL = 'ausgebucht'
    PROCESS_STATUS_APPROVED = 'Tour bewilligt & publiziert'
    EVENT_TYPES = ['Ftn', 'Div']


    def __init__(self, file_path):
        self.__cols_year_program = [
                self.COL_PROCESS_STATUS,
                self.COL_GROUP,
                self.COL_START_DATE,
                self.COL_DURATION,
                self.COL_TOUR_TYPE,
                self.COL_COND_REQ,
                self.COL_TECH_REQ,
                self.COL_ACTIVITY,
                self.COL_FIRST_NAME,
                self.COL_LAST_NAME,
        ]
        self.__cols_rote_karte = [
                self.COL_PROCESS_STATUS,
                self.COL_GROUP,
                self.COL_STATUS,
                self.COL_START_DATE,
                self.COL_END_DATE,
                self.COL_ACTIVITY,
                self.COL_TOUR_TYPE_LONG,
                self.COL_GUIDE,
                self.COL_COND_REQ,
                self.COL_TECH_REQ,
                self.COL_METRICS,
                self.COL_ROUTE,
                self.COL_HOSPITALITY,
                self.COL_COSTS,
                self.COL_COSTS_LONG,
                self.COL_EXECUTION,
                self.COL_MEETING_POINT,
                self.COL_DESCRIPTION,
                self.COL_ADDITIONAL_INFORMATION,
                self.COL_GEAR,
                self.COL_REGISTRATION_MEANS,
                self.COL_REGISTRATION_START_DATE,
                self.COL_REGISTRATION_END_DATE,
        ]
        self.__read_file(file_path)

    def parse(self, parser_type):
        """Parse the tours table and return a tours object"""
        self.__data[self.COL_GROUP] = self.__data.apply(lambda tour: self.__get_real_group(tour[self.COL_GROUP], tour[self.COL_TOUR_TYPE]), axis = 1)
        self.__data.sort_values([self.COL_START_DATE, self.COL_TOUR_TYPE], inplace = True)
        if parser_type == 'jahresprogramm':
            return pd.DataFrame(self.__data, columns = self.__cols_year_program)
        elif parser_type == 'rotekarte':
            return pd.DataFrame(self.__data, columns = self.__cols_rote_karte)

    def __read_file(self, file_path):
        """Store an object with the xls file content"""
        data = pd.read_excel(file_path, parse_dates=True, date_format='%Y-%m-%d')
        self.__data = data

    def __get_real_group(self, group, tour_type):
        if tour_type in self.EVENT_TYPES:
            return 'Events'
        return group
