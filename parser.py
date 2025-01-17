import pandas as pd
import os
import sys

current_directory = os.path.dirname(sys.argv[0]) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))


class Parser:
    time_only = {'Alos72', 'TalisDefault', 'Sunrise', 'ShemaMA72', 'ShemaMA72fix', 'ShemaGro', 'ShachrisGro', 'Midday',
                 'MinchaGroLechumra', 'KetanaGro', 'PlagGro', 'Candles', 'Sunset', 'Tzes3Stars', 'Tzes72'}

    def __init__(self):
        self.new_df = None
        file_path = current_directory + '\\badzman.xlsx'
        old_df = pd.read_excel(file_path)
        self.new_df = pd.DataFrame(
            columns=['WkDay', 'CivilDate', 'JewishDate', 'HolidayHebrew', 'Plaque1', 'Plaque2', 'Plaque3', 'Plaque4',
                     'bottomtext', 'extrabottomtext', 'HolidayEnglish', 'ParshaHebrew', 'ParshaEnglish', 'DafYomi',
                     'Omer', 'Alos72', 'TalisDefault', 'Sunrise', 'ShemaMA72', 'ShemaMA72fix', 'ShemaGro',
                     'ShachrisGro', 'Midday', 'MinchaGroLechumra', 'KetanaGro', 'PlagGro', 'Candles', 'Sunset',
                     'Tzes3Stars', 'Tzes72'])

        for column in self.new_df.columns:
            if column in old_df.columns:
                if column == 'CivilDate':
                    old_df['CivilDate'] = pd.to_datetime(old_df['CivilDate'], format='%m/%d/%Y')
                    self.new_df['CivilDate'] = old_df['CivilDate'].apply(
                        lambda x: f"{x.month}/{x.day}/{str(x.year)[-2:]}")
                elif column in self.__class__.time_only:
                    # Format time without seconds
                    self.new_df[column] = old_df[column].apply(
                        lambda x: x.strftime('%I:%M %p').lstrip('0') if pd.notnull(x) else x)
                else:
                    self.new_df[column] = old_df[column]

    def fill_plaque1(self):
        erev_pesach = self.new_df[self.new_df['HolidayEnglish'] == 'Erev Pesach'].index[0]
        shmini_atzeres = self.new_df[self.new_df['HolidayEnglish'] == 'Shmini Atzeres'].index[0]
        self.new_df.loc[:erev_pesach, 'Plaque1'] = 'משיב הרוח'
        self.new_df.loc[erev_pesach + 1:shmini_atzeres, 'Plaque1'] = 'ותן ברכה'
        self.new_df.loc[shmini_atzeres + 1:, 'Plaque1'] = 'משיב הרוח'

    def fill_plaque2(self):
        erev_pesach = self.new_df[self.new_df['HolidayEnglish'] == 'Erev Pesach'].index[0]
        self.new_df.loc[:erev_pesach, 'Plaque2'] = 'ותן תל ומטר'

        simchas_torah = self.new_df[self.new_df['HolidayEnglish'] == 'Simchas Torah'].index[0]
        self.new_df.loc[simchas_torah:, 'Plaque2'] = 'ותן ברכה'

        # start vesen tal umatar on december 5th
        vesen_tal_umatar = self.new_df[self.new_df['CivilDate'] == '12/5/24'].index[0]
        self.new_df.loc[vesen_tal_umatar:, 'Plaque2'] = 'ותן טל ומטר'

        rosh_hashanah = self.new_df[self.new_df['HolidayEnglish'] == 'Rosh Hashanah'].index[0]
        # yom kippur is always 10 days after first day of rosh hashanah
        self.new_df.loc[rosh_hashanah:rosh_hashanah + 10, 'Plaque2'] = 'המלך הקדוש'

        # isru chag after pesach
        isru_chag_pesach = self.new_df[self.new_df['HolidayEnglish'] == 'Isru Chag'].index[0]
        rosh_chodesh_eyar = self.new_df[self.new_df['HolidayHebrew'] == 'ראש חודש אייר'].index[0]
        self.new_df.loc[isru_chag_pesach + 1:rosh_chodesh_eyar - 1, 'Plaque2'] = "א" + "\u05F4" + "א " + "תחנון"
        # no tachanun in between yom kippur and succos
        self.new_df.loc[rosh_hashanah + 10:rosh_hashanah + 15, 'Plaque2'] = "א" + "\u05F4" + "א " + "תחנון"

        yaaleh_veyuvo = {'Rosh Chodesh', 'Pesach', 'Chol HaMoed', 'Erev Shvii shel Pesach', 'Shvii shel Pesach',
                         'Acharon shel Pesach', 'Shavuos', 'Sukkos', 'Shabbos Chol HaMoed', 'Hoshanah Rabbah',
                         'Shmini Atzeres', 'Simchas Torah'}
        no_tachanun = {'Pesach Sheini', 'Lag BaOmer', 'Isru Chag', 'Tu B׳Shvat', 'Shushan Purim'}

        al_hanisim = {'Purim', 'Chanukah', 'Shabbos Chanukah'}

        for index, row in self.new_df.iterrows():
            if row['HolidayEnglish'] in yaaleh_veyuvo:
                if pd.isna(row['Plaque2']):
                    self.new_df.loc[index, 'Plaque2'] = 'יעלה ויבא'
                elif pd.isna(row['Plaque3']):
                    self.new_df.loc[index, 'Plaque3'] = 'יעלה ויבא'
                elif pd.isna(row['Plaque4']):
                    self.new_df.loc[index, 'Plaque3'] = 'יעלה ויבא'
            if row['HolidayEnglish'] in no_tachanun:
                if pd.isna(row['Plaque2']):
                    self.new_df.loc[index, 'Plaque2'] = "א" + "\u05F4" + "א " + "תחנון"
                elif pd.isna(row['Plaque3']):
                    self.new_df.loc[index, 'Plaque3'] = "א" + "\u05F4" + "א " + "תחנון"
            if row['HolidayEnglish'] in al_hanisim:
                if pd.isna(row['Plaque2']):
                    self.new_df.loc[index, 'Plaque2'] = 'על הנסים'
                elif pd.isna(row['Plaque3']):
                    self.new_df.loc[index, 'Plaque3'] = 'על הנסים'

    def fill_plaque3(self):
        # say david adonai ori for 50 days starting after elul
        david_adonai_ori = self.new_df[self.new_df['HolidayHebrew'] == 'ראש חודש אלול'].index[1]
        self.new_df.loc[david_adonai_ori + 1:david_adonai_ori + 50, 'Plaque3'] = 'לדוד ה׳ אורי'

    def fill_plaque4(self):
        row_number = self.new_df.index[self.new_df['HolidayHebrew'] == 'ראש חודש אלול'].tolist()[1]
        self.new_df.at[row_number, 'Plaque4'] = 'לדוד ה׳ אורי'

        for index, row in self.new_df.iterrows():
            if pd.notna(row['Candles']):
                self.new_df.loc[index, 'Plaque4'] = ('\u05D4\u05D3\u05DC\u05F4\u05E0' + ' '
                                                     + row['Candles'].replace('PM', ''))

    def fill_dafyomitext(self):
        self.new_df['DafYomi'] = self.new_df['DafYomi'].apply(lambda x: f"דף יומי: {x}")

    def fill_perek(self):
        # Start after Pesach, end the Shabbos before Rosh Hashanah
        isru_chag_pesach = self.new_df[self.new_df['HolidayEnglish'] == 'Isru Chag'].index[0]
        rosh_hashanah_index = self.new_df[self.new_df['HolidayEnglish'] == 'Rosh Hashanah'].index[0]

        # Find the last Shabbos before Rosh Hashanah
        last_shabbos = self.new_df[(self.new_df.index < rosh_hashanah_index) & (self.new_df['WkDay'] == 'Sha')].index[
            -1]

        perakim = {0: 'א', 1: 'ב', 2: 'ג', 3: 'ד', 4: 'ה', 5: 'ו'}
        num_perakim = len(perakim)
        counter = 0

        # Find all Shabbats between Isru Chag Pesach and last Shabbat before Rosh Hashanah
        shabbat_indexes = self.new_df[(self.new_df.index >= isru_chag_pesach) & (self.new_df.index <= last_shabbos) & (
                self.new_df['WkDay'] == 'Sha')].index
        num_shabbats = len(shabbat_indexes)

        # Determine if doubling up is necessary
        doubling_up = False

        for index in self.new_df.index:
            if index > last_shabbos:
                break
            if index < isru_chag_pesach:
                continue

            if num_shabbats - counter == (num_perakim - counter % num_perakim) // 2:
                doubling_up = True

            # Calculate which chapter to assign
            text = 'פרק ' + perakim[counter % num_perakim] + '\u05F3'
            if doubling_up:
                text = text + '-' + perakim[counter % num_perakim + 1] + '\u05F3'

            if self.new_df.at[index, 'WkDay'] == 'Sha':
                counter += 2 if doubling_up else 1

            # Determine the correct plaque column to fill
            plaque_col = 'Plaque3' if pd.isna(self.new_df.at[index, 'Plaque3']) else 'Plaque2'
            self.new_df.at[index, plaque_col] = text

    # def fill_perek(self):
    #     # start after pesach, end the shabbos before rosh hashanah
    #     isru_chag_pesach = self.new_df[self.new_df['HolidayEnglish'] == 'Isru Chag'].index[0]
    #     rosh_hashanah_index = self.new_df[self.new_df['HolidayEnglish'] == 'Rosh Hashanah'].index[0]
    #
    #     # Find the last Shabbos before Rosh Hashanah
    #     last_shabbos = self.new_df[(self.new_df.index < rosh_hashanah_index) & (self.new_df['WkDay'] == 'Sha')].index[
    #         -1]
    #
    #     perakim, counter = {0: 'א', 1: 'ב', 2: 'ג', 3: 'ד', 4: 'ה', 5: 'ו'}, 0
    #
    #     for index in self.new_df.index:
    #         if index > last_shabbos:
    #             break
    #         if index < isru_chag_pesach:
    #             continue
    #
    #         text = 'פרק ' + perakim[counter % len(perakim)] + '\u05F3'
    #         if self.new_df.at[index, 'WkDay'] == 'Sha':
    #             counter += 1
    #
    #         plaque_col = 'Plaque3' if pd.isna(self.new_df.at[index, 'Plaque3']) else 'Plaque2'
    #         self.new_df.at[index, plaque_col] = text
    #
    #     # df_reversed = self.new_df['Plaque2'][::-1].copy()
    #
    #     # for index in df_reversed.index:
    #     #     if counter % len(perakim) != 0:
    #     #         break
    #     #     df_reversed.at[index] += ('-' + perakim[counter % len(perakim)] + '\u05F3')
    #     #     if self.new_df.at[index, 'WkDay'] == 'Sha':
    #     #         counter += 1

    def fill_parshah(self):
        current_parshah = None
        df_reversed = self.new_df['ParshaHebrew'][::-1].copy()
        for index in df_reversed.index:
            if pd.notna(df_reversed[index]):
                current_parshah = df_reversed[index]
            if current_parshah:
                df_reversed[index] = current_parshah
        self.new_df['ParshaHebrew'] = df_reversed[::-1].apply(lambda x: 'פרשת ' + str(x) if pd.notna(x) else x)

    def fill_bottomtext(self):
        hebrew_phrases = [
            "הַיּוֹם יוֹם אֶחָד לָעֹמֶר׃",
            "הַיּוֹם שְׁנֵי יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁלֹשָׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם אַרְבָּעָה יָמִים לָעֹמֶר׃",
            "הַיּוֹם חֲמִשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שִׁשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שִׁבְעָה יָמִים שֶׁהֵם שָׁבוּעַ אֶחָד לָעֹמֶר׃",
            "הַיּוֹם שְׁמוֹנָה יָמִים שֶׁהֵם שָׁבוּעַ אֶחָד וְיוֹם אֶחָד לָעֹמֶר׃",
            "הַיּוֹם תִּשְׁעָה יָמִים שֶׁהֵם שָׁבוּעַ אֶחָד וּשְׁנֵי יָמִים לָעֹמֶר׃",
            "הַיּוֹם עֲשָׂרָה יָמִים שֶׁהֵם שָׁבוּעַ אֶחָד וּשְׁלֹשָׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם אַחַד עָשָׂר יוֹם שֶׁהֵם שָׁבוּעַ אֶחָד וְאַרְבָּעָה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁנֵים עָשָׂר יוֹם שֶׁהֵם שָׁבוּעַ אֶחָד וַחֲמִשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁלֹשָׁה עָשָׂר יוֹם שֶׁהֵם שָׁבוּעַ אֶחָד וְשִׁשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם אַרְבָּעָה עָשָׂר יוֹם שֶׁהֵם שְׁנֵי שָׁבוּעוֹת לָעֹמֶר׃",
            "הַיּוֹם חֲמִשָּׁה עָשָׂר יוֹם שֶׁהֵם שְׁנֵי שָׁבוּעוֹת וְיוֹם אֶחָד לָעֹמֶר׃",
            "הַיּוֹם שִׁשָּׁה עָשָׂר יוֹם שֶׁהֵם שְׁנֵי שָׁבוּעוֹת וּשְׁנֵי יָמִים לָעֹמֶר׃",
            "הַיּוֹם שִׁבְעָה עָשָׂר יוֹם שֶׁהֵם שְׁנֵי שָׁבוּעוֹת וּשְׁלֹשָׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁמוֹנָה עָשָׂר יוֹם שֶׁהֵם שְׁנֵי שָׁבוּעוֹת וְאַרְבָּעָה יָמִים לָעֹמֶר׃",
            "הַיּוֹם תִּשְׁעָה עָשָׂר יוֹם שֶׁהֵם שְׁנֵי שָׁבוּעוֹת וַחֲמִשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם עֶשְׂרִים יוֹם שֶׁהֵם שְׁנֵי שָׁבוּעוֹת וְשִׁשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם אֶחָד וְעֶשְׂרִים יוֹם שֶׁהֵם שְׁלֹשָׁה שָׁבוּעוֹת לָעֹמֶר׃",
            "הַיּוֹם שְׁנַֽיִם וְעֶשְׂרִים יוֹם שֶׁהֵם שְׁלֹשָׁה שָׁבוּעוֹת וְיוֹם אֶחָד לָעֹמֶר׃",
            "הַיּוֹם שְׁלֹשָׁה וְעֶשְׂרִים יוֹם שֶׁהֵם שְׁלֹשָׁה שָׁבוּעוֹת וּשְׁנֵי יָמִים לָעֹמֶר׃",
            "הַיּוֹם אַרְבָּעָה וְעֶשְׂרִים יוֹם שֶׁהֵם שְׁלֹשָׁה שָׁבוּעוֹת וּשְׁלֹשָׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם חֲמִשָּׁה וְעֶשְׂרִים יוֹם שֶׁהֵם שְׁלֹשָׁה שָׁבוּעוֹת וְאַרְבָּעָה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שִׁשָּׁה וְעֶשְׂרִים יוֹם שֶׁהֵם שְׁלֹשָׁה שָׁבוּעוֹת וַחֲמִשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שִׁבְעָה וְעֶשְׂרִים יוֹם שֶׁהֵם שְׁלֹשָׁה שָׁבוּעוֹת וְשִׁשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁמוֹנָה וְעֶשְׂרִים יוֹם שֶׁהֵם אַרְבָּעָה שָׁבוּעוֹת לָעֹמֶר׃",
            "הַיּוֹם תִּשְׁעָה וְעֶשְׂרִים יוֹם שֶׁהֵם אַרְבָּעָה שָׁבוּעוֹת וְיוֹם אֶחָד לָעֹמֶר׃",
            "הַיּוֹם שְׁלֹשִׁים יוֹם שֶׁהֵם אַרְבָּעָה שָׁבוּעוֹת וּשְׁנֵי יָמִים לָעֹמֶר׃",
            "הַיּוֹם אֶחָד וּשְׁלֹשִׁים יוֹם שֶׁהֵם אַרְבָּעָה שָׁבוּעוֹת וּשְׁלֹשָׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁנַֽיִם וּשְׁלֹשִׁים יוֹם שֶׁהֵם אַרְבָּעָה שָׁבוּעוֹת וְאַרְבָּעָה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁלֹשָׁה וּשְׁלֹשִׁים יוֹם שֶׁהֵם אַרְבָּעָה שָׁבוּעוֹת וַחֲמִשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם אַרְבָּעָה וּשְׁלֹשִׁים יוֹם שֶׁהֵם אַרְבָּעָה שָׁבוּעוֹת וְשִׁשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם חֲמִשָּׁה וּשְׁלֹשִׁים יוֹם שֶׁהֵם חֲמִשָּׁה שָׁבוּעוֹת לָעֹמֶר׃",
            "הַיּוֹם שִׁשָּׁה וּשְׁלֹשִׁים יוֹם שֶׁהֵם חֲמִשָּׁה שָׁבוּעוֹת וְיוֹם אֶחָד לָעֹמֶר׃",
            "הַיּוֹם שִׁבְעָה וּשְׁלֹשִׁים יוֹם שֶׁהֵם חֲמִשָּׁה שָׁבוּעוֹת וּשְׁנֵי יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁמוֹנָה וּשְׁלֹשִׁים יוֹם שֶׁהֵם חֲמִשָּׁה שָׁבוּעוֹת וּשְׁלֹשָׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם תִּשְׁעָה וּשְׁלֹשִׁים יוֹם שֶׁהֵם חֲמִשָּׁה שָׁבוּעוֹת וְאַרְבָּעָה יָמִים לָעֹמֶר׃",
            "הַיּוֹם אַרְבָּעִים יוֹם שֶׁהֵם חֲמִשָּׁה שָׁבוּעוֹת וַחֲמִשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם אֶחָד וְאַרְבָּעִים יוֹם שֶׁהֵם חֲמִשָּׁה שָׁבוּעוֹת וְשִׁשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁנַֽיִם וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁשָּׁה שָׁבוּעוֹת לָעֹמֶר׃",
            "הַיּוֹם שְׁלֹשָׁה וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁשָּׁה שָׁבוּעוֹת וְיוֹם אֶחָד לָעֹמֶר׃",
            "הַיּוֹם אַרְבָּעָה וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁשָּׁה שָׁבוּעוֹת וּשְׁנֵי יָמִים לָעֹמֶר׃",
            "הַיּוֹם חֲמִשָּׁה וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁשָּׁה שָׁבוּעוֹת וּשְׁלֹשָׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שִׁשָּׁה וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁשָּׁה שָׁבוּעוֹת וְאַרְבָּעָה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שִׁבְעָה וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁשָּׁה שָׁבוּעוֹת וַחֲמִשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם שְׁמוֹנָה וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁשָּׁה שָׁבוּעוֹת וְשִׁשָּׁה יָמִים לָעֹמֶר׃",
            "הַיּוֹם תִּשְׁעָה וְאַרְבָּעִים יוֹם שֶׁהֵם שִׁבְעָה שָׁבוּעוֹת לָעֹמֶר׃"
        ]

        for index, row in self.new_df.iterrows():
            if pd.notna(row['Omer']):
                self.new_df.loc[index, 'bottomtext'] = hebrew_phrases[int(row['Omer']) - 1]
        self.new_df['bottomtext'] = self.new_df['bottomtext'].fillna('שויתי ה׳ לנגדי תמיד')

    # fill in for the friday and shabbos before every rosh chodesh

    def to_csv(self):
        # Save the DataFrame to a CSV file
        self.new_df.to_csv(current_directory + '\\Zmanim.csv',
                           index=False)

    def to_excel(self):
        self.new_df.to_excel(current_directory + '\\zmanexcel.xlsx',
                             index=False)
