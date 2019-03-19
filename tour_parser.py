import bs4
import ssl
import urllib

class TourParser(object):
    """Parse tours retrieved from the SAC Aarau tour page"""

    def __init__(self, url, startDate, endDate):
        self.tours = {}
        self.__read_page(url, startDate, endDate)

    def parse(self):
        print(self.__page.prettify())

    def __read_page(self, url, start_date, end_date):
        """Returns a BeautifulSoup object of the given url"""
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