#-*- coding:utf-8 -*-

import os
import re
import xlrd
from xlrd.sheet import Cell

cp = re.compile('\[(.*?)\]')


class DataReport():

    def __init__(self, filedir, stat_date):
        self.filedir = filedir
        self.stat_date = stat_date

    def get_data(self):
        data = []

        workbook = xlrd.open_workbook(self.filedir)
        sheet_names = workbook._sheet_names

        for sheet in sheet_names:
            s = workbook.sheet_by_name(sheet)
            n_rows = s.nrows
            n_cols = s.ncols

            for i_row in range(1, n_rows):
                row_data = []
                for i_col in range(n_cols):
                    value = None

                    if i_col == 0:
                        value = Cell(1, s._cell_values[i_row][i_col])
                        value = "%s"%value
                        value = value.split(':')[1]
                        value = "'%s"%int(float(value))
                    else:
                        value = s.cell_value(i_row, i_col)
                    row_data.append(value)
                    if i_col == 0:
                        row_data.append(self.stat_date)

                row_data = map(str, row_data)
                row_data_str = '\t'.join(row_data)

                data.append(row_data_str)

        return data


def main():
    data_prefix = './data'
    files = os.listdir(data_prefix)
    files = filter(lambda x: x.endswith('.xls'), files)

    w = open('after_merge.csv', 'w')

    # headers need to be filled
    headers = []
    headers_str = '\t'.join([])

    w.write('%s\n'%headers_str)

    for f in files:
        dates = cp.findall(f)[0]
        report = DataReport(data_prefix+'/'+f, dates)

        for i in report.get_data():
            w.write('%s\n'%i)

if __name__ == '__main__':
    main()
