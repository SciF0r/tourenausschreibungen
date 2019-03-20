import bs4
import ssl
import urllib

class TourParser(object):
    """Parse tours retrieved from the SAC Aarau tour page"""

    def __init__(self, url, startDate, endDate):
        self.__read_page(url, startDate, endDate)

    def parse(self):
        """Parse the tours table and return a tours object"""
        self.__tours = {}
        table = self.__page.find('table', class_='program')
        for cell in table.find_all('td'):
            self.__process_cell(cell)
        return self.__tours;

    def __read_page(self, url, start_date, end_date):
        """Store a BeautifulSoup object of the given url"""
        data = {
            'start': start_date,
            'end': end_date,
            'published': 'on'
        }
        with urllib.request.urlopen(
            url,
            context=ssl._create_unverified_context(),
            data=urllib.parse.urlencode(data).encode("utf-8")
        ) as page:
            self.__page = bs4.BeautifulSoup(page, 'html5lib')

    def __process_cell(self, cell):
        """Process a cell and update the tours object with its contents"""
        if 'class' in cell.attrs:
            cellClasses = cell['class']
            if len(cellClasses) > 1:
                print('More than one css class, ignoring cell')
                return
            if len(cellClasses) == 1:
                cellType = cellClasses[0]
                print(cellType)
        else:
            print('no class')
