import sys
from datetime import datetime
from docx_creator_adressen import DocxCreatorAdressen
from docx_creator_jahresprogramm import DocxCreatorJahresprogramm
from docx_creator_rotekarte import DocxCreatorRoteKarte
from docx_creator_rotekarteneu import DocxCreatorRoteKarteNeu
from tour_cleaner import TourCleaner
from address_parser import AddressParser
from tour_parser import TourParser

if len(sys.argv) != 3:
    print('Usage: {0} adressen|jahresprogramm|rotekarte path_to_xlsx_file'.format(sys.argv[0]))
    sys.exit(1)

parser_type = sys.argv[1].lower()
file_path = sys.argv[2]

print('Getting tours from {0}'.format(file_path))

if parser_type not in ['adressen', 'jahresprogramm', 'rotekarte', 'rotekarteneu']:
    print('Usage: {0} adressen|jahresprogramm|rotekarte path_to_xlsx_file'.format(sys.argv[0]))
    sys.exit(1)

if parser_type == 'adressen':
    parser = AddressParser(file_path)
    addresses = parser.parse_for_year_program()

    export_file_name = '{0}_adressen.docx'.format(datetime.now().strftime('%Y-%m-%d_%H_%M_%S'))
    docx_creator = DocxCreatorAdressen(addresses, export_file_name)
else:
    parser = TourParser(file_path)
    tours = parser.parse(parser_type)
    tours = tours[tours[TourParser.COL_PROCESS_STATUS] == TourParser.PROCESS_STATUS_APPROVED]
    cleaner = TourCleaner(tours)
    with open('rules.txt', 'r') as rules_file:
        rules = [line.split('/', 2) for line in rules_file.read().splitlines()]
    cleaned_tours = cleaner.clean(rules)

    if parser_type == 'jahresprogramm':
        export_file_name = '{0}_{1}_{2}.docx'.format(datetime.now().strftime('%Y-%m-%d_%H_%M_%S'), parser_type, '{0}')
        docx_creator = DocxCreatorJahresprogramm(cleaned_tours, export_file_name)
    elif parser_type == 'rotekarteneu':
        export_file_name = '{0}_{1}_{2}.docx'.format(datetime.now().strftime('%Y-%m-%d_%H_%M_%S'), parser_type, '{0}')
        docx_creator = DocxCreatorRoteKarteNeu(cleaned_tours, export_file_name)
    else:
        export_file_name = '{0}_{1}.docx'.format(datetime.now().strftime('%Y-%m-%d_%H_%M_%S'), parser_type)
        docx_creator = DocxCreatorRoteKarte(cleaned_tours, export_file_name)

docx_creator.create()
