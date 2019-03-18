import bs4
import ssl
import urllib.request as req

def get_html(url):
    """Returns a BeautifulSoup object of the given url"""
    page = req.urlopen(url, context=ssl._create_unverified_context())
    return bs4.BeautifulSoup(page, 'html5lib')
    
page = get_html('https://www.sac-aarau.ch/aktivitaeten/kalender/')
print(page.prettify());