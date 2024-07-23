from parser import Parser
import openpyxl

zmanim_parser = Parser()
zmanim_parser.fill_plaque1()
zmanim_parser.fill_plaque2()
zmanim_parser.fill_plaque3()
zmanim_parser.fill_plaque4()
zmanim_parser.to_csv()