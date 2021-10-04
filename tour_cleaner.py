import pandas as pd
import re

class TourCleaner(object):
    """Clean a tours object"""
    
    def __init__(self, tours):
        self.__tours = tours

    def clean(self, rules):
        for pattern, replacement in rules:
            self.__apply_rule(pattern, replacement)
        return self.__tours

    def __apply_rule(self, pattern, replacement):
        for row in range(self.__tours.shape[0]):
            for column in range(self.__tours.shape[1]):
                cell_value = self.__tours.iat[row, column]
                if (type(cell_value) is not str):
                    continue
                self.__tours.iat[row, column] = re.sub(pattern, replacement, cell_value)
