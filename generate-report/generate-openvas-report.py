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
TEMPLATE_FILENAME = f'{PWD}/template.xlsm'
TMP_FILENAME = f'{PWD}/tmp.xlsx'
ZF_TMP_PATH = f'{PWD}/_x_tmp'


def save(filename, df):
    """Save DataFrame to the file
    """

    # Write temp excel file
    with pd.ExcelWriter(TMP_FILENAME, engine='openpyxl') as wt:
        wb = openpyxl.load_workbook(TEMPLATE_FILENAME, keep_vba=True)
        wt.book = wb
        wt.sheets = dict((ws.title, ws) for ws in wb.worksheets)
        wt.vba_archive = wb.vba_archive

        df.to_excel(
            wt, sheet_name=DATA_SHEET, index=False, header=False, startrow=2
        )
        wt.save()

    # Decompress temp file that we stored above.
    with zipfile.ZipFile(TMP_FILENAME, 'r') as zf:
        zf.extractall(ZF_TMP_PATH)

    # Decompress the template.xlsm file to temp directory, not template dir
    # This work was needed to make VB macro work without error.
    with zipfile.ZipFile(TEMPLATE_FILENAME, 'r') as zf:
        extracts = (
            '[Content_Types].xml',
            'xl/_rels/workbook.xml.rels',
            'xl/vbaProject.bin',
            'xl/sharedStrings.xml'
        )
        for fn in extracts:
            zf.extract(fn, path=ZF_TMP_PATH)

    # Compress temp dir to the filename (output)
    with zipfile.ZipFile(filename, 'w') as zf:
        os.chdir(ZF_TMP_PATH)
        for root, dirs, files in os.walk('./'):
            for file in files:
                zf.write(os.path.join(root, file))

    # Clean working dirs
    rmtree(ZF_TMP_PATH)
    os.remove(f'{PWD}/tmp.xlsx')


def load_csvfiles():
    """Load csv files and return the concatenated DataFrame
    """
    os.chdir(sys.argv[1])
    files = list(filter(lambda x: x.find('.csv') != -1, os.listdir()))

    df = None
    for fn in files:
        if os.path.isfile(fn) is False:
            print(f'{fn} not found')
            sys.exit(2)
        print(f'[|] {fn} file')
        _b = pd.read_csv(fn)
        df = _b if df is None else pd.concat([df, _b])

    headers = df.columns.values
    col_len, row_len = len(headers), len(df)
    print(f'[+] {row_len} x {col_len} table was loaded')
    return df


def main():
    print(f'[+] load csv files and merge')
    df = load_csvfiles()

    print(f'[*] sort dataframe')
    df = df.sort_values(
        by=['CVSS', 'NVT Name', 'IP', 'Port'],
        ascending=[False, True, True, True]
    ).reset_index(drop=True)
    df.insert(0, 'IDX', df.index + 1)

    outfilename = datetime.date.today().strftime('%Y%m%d') + '-report.xlsm'
    print(f'[*] Save report : {outfilename}')
    save(f'{PWD}/{outfilename}', df)


if __name__ == '__main__':
    if len(sys.argv) == 2:
        main()
    else:
        pname = sys.argv[0].lstrip('./')
        print(f'Usage: {pname} <csv directory name>')
        print(f'ex) ./{pname} data  # data is a name of directory')
        sys.exit(-1)

