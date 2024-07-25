from parser import Parser

# pulls a excel located in the __init__ method
# creates a csv file in the path located in the to_csv method, csv file must be called Zmanim
# creates a excel in the path located in the to_excel method for any needed customization

zmanim_parser = Parser()
zmanim_parser.fill_plaque1()
zmanim_parser.fill_plaque2()
zmanim_parser.fill_plaque3()
zmanim_parser.fill_plaque4()
zmanim_parser.fill_bottomtext()
zmanim_parser.fill_dafyomitext()
zmanim_parser.fill_parshah()
zmanim_parser.fill_perek()
zmanim_parser.to_csv()
zmanim_parser.to_excel()