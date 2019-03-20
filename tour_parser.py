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
        for row in table.find_all('tr'):
            self.__process_row(row)
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

    def __process_row(self, row):
        """Process a row and return a tours object with its contents"""
        left, right = row.find_all('td')
        print(left.prettify())
        print(right.prettify())