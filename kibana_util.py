import csv
import time
from datetime import datetime
from gsheet_util import GSheet_Util

month_dict = {
    'Jan': 1,
    'Feb': 2,
    'Mar': 3,
    'Apr': 4,
    'May': 5,
    'Jun': 6,
    'Jul': 7,
    'Aug': 8,
    'Sep': 9,
    'Oct': 10,
    'Nov': 11,
    'Dec': 12
}

# Order of keys to write in the Cartographers worksheet
key_order = ['Approved', 'Corrected', 'CriticalApproved', 'Disapproved', 'Duplicate', 'Irrelevant', 'NoData', 'Seen',
                'HighlyApproved', 'NotSeen']

# Dictionary of users in Stalin and their tabs in the worksheet
editor_dict = {
    'Nastaran.Sabouri': 'Nastaran_Stalin',
    'mahta.bakhshipour': 'Mahta_Stalin',
    'fattaneh.kia': 'Fati_Stalin',
    'niloofar.khosravi': 'Niloofar_Stalin',
    'elham.nasimzadeh': 'Elham_Stalin'
}


class Reporter:
    def __init__(self, keys):
        self._keys = keys
        self._notSeen_idx = keys.index('NotSeen')
        self._report_list = []
        self._gsheet = GSheet_Util(sheet_name='Cartographers Worksheet',
                         auth_file='../gsheet_auth/model-journal-343507-ff222c010398.json')

    def add_report(self, report, date):
        self._report_list.append((report, date))

    def write_to_sheet(self, sheet_tab):
        report_len = len(self._report_list) - 1
        self._report_list.reverse()
        l = []
        for i in range(report_len):
            diff = []
            for j in range(10):
                if j == self._notSeen_idx:
                    diff.append(int(self._report_list[i+1][0][j]))
                else:
                    diff.append(int(self._report_list[i+1][0][j]) - int(self._report_list[i][0][j]))

            # create dict from key in input excel and diff as value
            res = dict(zip(self._keys, diff))
            dict_out = {}
            for k in key_order:
                dict_out[k] = res[k]

            # create date from month name and day
            dd, tt = self._report_list[i + 1][1].split('@')
            date = f"{datetime.strptime(dd.strip(), '%b %d, %Y').date()}"
            diff.insert(0, date)

            # create data to insert row in the order of key_order list
            row = [date]
            sum = 0
            for k in dict_out:
                if k != 'NotSeen':
                    sum += dict_out[k]
                else:
                    row.append(sum)
                row.append(dict_out[k])

            l.append(row)
        self._gsheet.write_to_sheet(sheet_tab=sheet_tab, value=l, single_row=False)


def extract_excel(input_csv):
    """
    Method to read CSV generated from Kibana panel and write result to the Cartographers sheet.
    :param input_csv: the path of generated CSV by Kibana
    :return:
    """
    with open(input_csv, 'r') as reader:

        id_col = '_id'
        editor_col = 'editor'
        timestamp_col = 'timestamp'
        lines = csv.reader(reader, delimiter=',', quotechar='"')

        idx = 0
        reporters = {}
        for line in lines:

            idx += 1
            if idx == 1:
                id_idx = line.index(id_col)
                editor_idx = line.index(editor_col)
                timestamp_idx = line.index(timestamp_col)
                keys = line[0:id_idx]
                continue

            editor = line[editor_idx]
            if editor not in reporters:
                reporter = Reporter(keys)
                reporter.add_report(line[0:id_idx], line[timestamp_idx])
                reporters[editor] = reporter
            else:
                reporters[editor].add_report(line[0:id_idx], line[timestamp_idx])

        cnt = 0
        for k in reporters:
            cnt += 1
            if cnt > 60:
                time.sleep(100)
                cnt = 0
            reporters[k].write_to_sheet(sheet_tab=editor_dict[k])


extract_excel('input_csv.csv')