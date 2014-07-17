# -*- coding:utf-8 -*-

import os
import sys
import xlwt

class Excel(object):

    def __init__(self, font_name='Times New Roman', bold=False):
        self.font_name = font_name
        self.bold = bold
        self.create_wb()

    @property
    def font(self):
        font = xlwt.Font()
        font.name = self.font_name
        font.bold = self.bold
        return font

    @property
    def style(self):
        style = xlwt.XFStyle()
        style.font = self.font
        return style

    def create_wb(self):
        self.wb = xlwt.Workbook()

    def check(self, save_path, filename, csvfile=[]):

        if not os.path.exists(save_path):
            os.makedirs(save_path)

        file_path = os.path.join(save_path, filename)
        if os.path.exists(file_path):
            print 'The file has already exists in the save_path, please change another.'
            sys.exit(0)

        if not csvfile:
            print "no csv file."
            sys.exit(0)

        return file_path

    def csv_to_xls(self, save_path, filename, csvfile=[]):
        
        file_path = self.check(save_path, filename, csvfile=csvfile)

        for index, csv in enumerate(csvfile):
            with open(csv) as f:
                name = f.name.split('/')[-1]
                ws = self.wb.add_sheet(name)
                for rindex, line in enumerate(f.readlines()):
                    row = line.strip().split(",")
                    for cindex, cn in enumerate(row):
                        ws.write(rindex, cindex, cn)
        
        self.wb.save(file_path)

if __name__ == '__main__':
    obj = Excel()
    files = ["/Users/ZJN/other/a.csv", "/Users/ZJN/other/b.csv"]
    obj.csv_to_xls('/Users/ZJN/other', 'new1.xls', csvfile=files)
