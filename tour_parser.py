import pandas as pd

class TourParser(object):
    """Parse tours retrieved from a DropTours xlsx export file"""
    COL_GROUP = 'Gruppe'
    COL_START_DATE = 'Startdatum'
    COL_DURATION = 'Dauer'
    COL_TOUR_TYPE = 'Tourtyp'
    COL_COND_REQ = 'Kond. Anforderungen'
    COL_TECH_REQ = 'Techn. Anforderungen'
    COL_ACTIVITY = 'Aktivität'
    COL_FIRST_NAME = 'Vorname'
    COL_LAST_NAME = 'Name'


    def __init__(self, file_path):
        self.__cols_year_program = [
                self.COL_GROUP,
                self.COL_START_DATE,
                self.COL_DURATION,
                self.COL_TOUR_TYPE,
                self.COL_COND_REQ,
                self.COL_TECH_REQ,
                self.COL_ACTIVITY,
                self.COL_FIRST_NAME,
                self.COL_LAST_NAME
        ]
        self.__read_file(file_path)

    def parse_for_year_program(self):
        """Parse the tours table and return a tours object"""
        self.__data[self.COL_GROUP] = self.__data.apply(lambda tour: self.__get_real_group(tour[self.COL_GROUP], tour[self.COL_TOUR_TYPE]), axis = 1)
        self.__data.sort_values([self.COL_GROUP, self.COL_START_DATE, self.COL_TOUR_TYPE], inplace = True)
        return pd.DataFrame(self.__data, columns = self.__cols_year_program)

    def __read_file(self, file_path):
        """Store an object with the xls file content"""
        data = pd.read_csv(file_path, sep = ',')
        data[self.COL_START_DATE] = pd.to_datetime(data[self.COL_START_DATE], format = '%d.%m.%y')
        self.__data = data

    def __get_real_group(self, group, tour_type):
        if tour_type == 'Anl' or tour_type == 'Ftn' or tour_type == 'Div':
            return 'Events'
        return group
