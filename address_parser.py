import pandas as pd

class AddressParser(object):
    """Parse addresses retrieved from xls file prepared by someone(TM)"""
    COL_FUNCTION = 'Funktion'
    COL_COMMITTEE = 'Gremium'
    COL_FIRST_NAME = 'Vorname'
    COL_LAST_NAME = 'Name'
    COL_ADDRESS = 'Strasse, Nr.'
    COL_ZIP = 'PLZ'
    COL_TOWN = 'Ort'
    COL_EMAIL = 'E-Mail'
    COL_PHONE_PRIVATE = 'Telefon P'
    COL_PHONE_BUSINESS = 'Telefon G'
    COL_PHONE_MOBILE = 'Mobile'
    COL_TYPE_SOMMER1 = 'Sommer 1'
    COL_TYPE_SOMMER2 = 'Sommer 2'
    COL_TYPE_WINTER1 = 'Winter 1'
    COL_TYPE_WINTER2 = 'Winter 2'
    COL_TYPE_SPORTCLIMBING = 'Sport-\nklettern'
    COL_TYPE_SNOWSHOETOURS = 'Schnee-\nschuhtouren'
    COL_TYPE_MOUNTAINHIKING = 'Berg-\nwandern'
    COL_TYPE_ALPINEHIKING = 'Alpin-\nwandern'

    def __init__(self, file_path):
        self.__cols_addresses = [
                self.COL_FUNCTION,
                self.COL_COMMITTEE,
                self.COL_FIRST_NAME,
                self.COL_LAST_NAME,
                self.COL_ADDRESS,
                self.COL_ZIP,
                self.COL_TOWN,
                self.COL_EMAIL,
                self.COL_PHONE_PRIVATE,
                self.COL_PHONE_BUSINESS,
                self.COL_PHONE_MOBILE,
                self.COL_TYPE_SOMMER1,
                self.COL_TYPE_SOMMER2,
                self.COL_TYPE_WINTER1,
                self.COL_TYPE_WINTER2,
                self.COL_TYPE_SPORTCLIMBING,
                self.COL_TYPE_SNOWSHOETOURS,
                self.COL_TYPE_MOUNTAINHIKING,
                self.COL_TYPE_ALPINEHIKING
        ]
        self.__read_file(file_path)

    def parse_for_year_program(self):
        """Parse the addresses table and return an addresses object"""
        return pd.DataFrame(self.__data, columns = self.__cols_addresses)

    def __read_file(self, file_path):
        """Store an object with the xls file content"""
        data = pd.read_csv(file_path, sep = ',')
        data[self.COL_ZIP] = data[self.COL_ZIP].astype(pd.Int64Dtype())
        self.__data = data
