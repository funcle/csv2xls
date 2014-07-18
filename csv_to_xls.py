# -*- coding:utf-8 -*-

import os
import sys
import xlwt

class Excel(object):

    def __init__(self, **kwargs):
        self.create_font(**kwargs)
        self.create_style()
        self.create_wb()
    
    def create_font(self, **kwargs):
        self.font = xlwt.Font()
        self.font.__dict__.update(kwargs)

    def create_style(self):
        self.style = xlwt.XFStyle()
        self.style.font = self.font

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
                        ws.write(rindex, cindex, cn, self.style)
        
        self.wb.save(file_path)

if __name__ == '__main__':
    di = dict(shadow=False, bold=False, name="Arial", height=240)
    obj = Excel(**di)
    files = ["/Users/ZJN/other/a.csv", "/Users/ZJN/other/b.csv"]
    obj.csv_to_xls('/Users/ZJN/other', 'new4.xls', csvfile=files)
