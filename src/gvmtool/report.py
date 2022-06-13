#!/usr/bin/env python3
import pandas as pd
import sys
import os
import datetime
import openpyxl
import zipfile
from shutil import rmtree


PWD = os.getcwd()
DATA_SHEET = 'rawdata'
TEMPLATE_FILENAME = f'{os.path.dirname(__file__)}/../data/template.xlsm'
TMP_FILENAME = f'{PWD}/tmp.xlsx'
ZF_TMP_PATH = f'{PWD}/_x_tmp'
OUTFILE_EXT = 'xlsm'


class Report(object):
    def __init__(self):
        self.df = None # data frame

    def load_csv_file(self, filename):
        print(f'[|] {filename} file')
        _b = pd.read_csv(filename)
        self.df = _b if self.df is None else pd.concat([self.df, _b])

    def load_csv_files(self, path):
        os.chdir(path)
        files = list(filter(lambda x: x.find('.csv') != -1, os.listdir()))

        for fn in files:
            if os.path.isfile(fn) is False:
                print(f'{fn} not found')
                sys.exit(2)
            self.load_csv_file(fn)
    
    def save_to_xlsm_file(self, filename):
        with pd.ExcelWriter(TMP_FILENAME, engine='openpyxl') as wt:
            wb = openpyxl.load_workbook(TEMPLATE_FILENAME, keep_vba=True)
            wt.book = wb
            wt.sheets = dict((ws.title, ws) for ws in wb.worksheets)
            wt.vba_archive = wb.vba_archive

            self.df.to_excel(
                wt, sheet_name=DATA_SHEET, index=False, header=False, startrow=2
            )
            wt.save()

        with zipfile.ZipFile(TMP_FILENAME, 'r') as zf:
            zf.extractall(ZF_TMP_PATH)

        with zipfile.ZipFile(TEMPLATE_FILENAME, 'r') as zf:
            extracts = (
                '[Content_Types].xml',
                'xl/_rels/workbook.xml.rels',
                'xl/vbaProject.bin',
                'xl/sharedStrings.xml'
            )
            for fn in extracts:
                zf.extract(fn, path=ZF_TMP_PATH)

        with zipfile.ZipFile(filename, 'w') as zf:
            os.chdir(ZF_TMP_PATH)
            for root, dirs, files in os.walk('./'):
                for file in files:
                    zf.write(os.path.join(root, file))

        #clean
        rmtree(ZF_TMP_PATH)
        os.remove(f'{PWD}/tmp.xlsx')

    def merge(self):
        print(f'[*] sort dataframe')
        self.df = self.df.sort_values(
            by=['CVSS', 'NVT Name', 'IP', 'Port'],
            ascending=[False, True, True, True]
        ).reset_index(drop=True)
        self.df.insert(0, 'IDX', self.df.index + 1)

        headers = self.df.columns.values
        col_len, row_len = len(headers), len(self.df)
        print(f'[+] {row_len} x {col_len} table was loaded')


def merge_report(path, out):
    print(f'[+] load csv files and merge')
    report = Report()
    report.load_csv_files(path)
    
    report.merge()

    outfilename = f'{PWD}/'
    if out is not None:
        outfilename += f'{out}.{OUTFILE_EXT}'
    else:
        outfilename += datetime.date.today().strftime('%Y%m%d')
        outfilename += '-Monthly VA.xlsm'
        

    print(f'[*] Save report : {outfilename}')
    report.save_to_xlsm_file(outfilename)


# Backup Code
# with pd.ExcelWriter(f'{PWD}/test.xlsm', engine='openpyxl') as wt:
#     wb = openpyxl.load_workbook(TEMPLATE_FILENAME, keep_vba=True)
#     wt.book = wb
#     wt.sheets = dict((ws.title, ws) for ws in wb.worksheets)
#     wt.vba_archive = wb.vba_archive

#     sorted.to_excel(wt, sheet_name=DATA_SHEET, index=False, header=False, startrow=2)
#     wt.save()
