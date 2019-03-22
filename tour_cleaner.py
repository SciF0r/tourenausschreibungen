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
        print('replacing all {0} with {1}'.format(pattern, replacement))
        for tour_type, tours_of_type in self.__tours.items():
            new_tours = []
            for tour in tours_of_type:
                new_tour = []
                for line in tour:
                    left, right = line
                    if (left is not None):
                        left = re.sub(pattern, replacement, left)
                    if (right is not None):
                        right = re.sub(pattern, replacement, right)
                    new_tour.append((left, right))
                new_tours.append(new_tour)

            self.__tours[tour_type] = new_tours

