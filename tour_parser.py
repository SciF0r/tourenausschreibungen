import bs4
import json
import requests
import ssl

class TourParser(object):
    """Parse tours retrieved from the SAC Aarau tour page"""

    def __init__(self, url, startDate, endDate, username, password):
        self.__session = requests.Session()
        self.__login(username, password)
        self.__read_page(url, startDate, endDate)

    def parse(self):
        """Parse the tours table and return a tours object"""
        self.__current_title = ''
        self.__tours = {}
        self.__current_tour = []
        table = self.__page.find('table', class_='program')
        for row in table.find_all('tr'):
            self.__process_row(row)
        self.__tours[self.__current_title].append(self.__current_tour)
        return self.__tours;

    def __login(self, username, password):
        data = {
            'site': 'sac-aarau',
            'username': username,
            'password': password
        }
        self.__session.post('https://ssl.dropnet.ch/manager/login', verify=False, data=data)

    def __read_page(self, url, start_date, end_date):
        """Store a BeautifulSoup object of the given url"""
        data = {
            'start': start_date,
            'end': end_date,
            'published': 'on'
        }
        r = self.__session.post(url, verify=False, data=data)
        self.__page = bs4.BeautifulSoup(r.text, 'html5lib')

    def __process_row(self, row):
        """Process a row and update the tours object with its contents"""
        cells = row.find_all('td')
        if len(cells) == 1:
            return
        if len(cells) > 2:
            print('More than two cells in row, this should not happen!')
            return
        left_cell, right_cell = cells
        if not left_cell.string and not right_cell.string:
            return
        skip_row = self.__process_cell_class(left_cell) or self.__process_cell_class(right_cell)
        if skip_row:
            return
        self.__add_key_value_tuple_to_current_tour(left_cell.get_text('\n'), right_cell.get_text('\n'))

    def __process_cell_class(self, cell):
        """Process a cell according to its css class"""
        if not 'class' in cell.attrs:
            return
        if len(cell['class']) > 1:
            print('More than one css class for cell "{0}"'.format(cell.string))
            return False
        cell_type = cell['class'][0]
        if cell_type == 'title':
            self.__current_title = cell.string
            self.__tours[self.__current_title] = []
            self.__current_tour = []
            return True
        if cell_type == 'datum':
            if self.__current_tour != []:
                self.__tours[self.__current_title].append(self.__current_tour)
            self.__current_tour = []
            return False
        return False

    def __add_key_value_tuple_to_current_tour(self, key, value):
        """Process a cell if no css class is set"""
        self.__current_tour.append((key, value))
