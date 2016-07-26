#!/usr/bin/env python
"""
Created on Thu Jul 14 15:22:58 2016

@author: Christopher.Rieve

Goal: Create a python program template to grab all excel files and various
      information that may prove helpful
ToDo:   - Make paths relative
        - Remove rows with zero columns and rows
        - Make right click menu option
        - Rewrite functionally
        - Add column headers in excel sheet
Notes: worksheet.write(0, 0, "I'm sheet number %d" % (i + 1)) <-- Could come in
       handy later.
"""
# pylint: disable=C0103
# pylint: disable=W0312

import os
import sqlite3

import readablebytes
import xlrd
import xlsxwriter

# Windows
rootdir = r'C:\Users\christopher.rieve\NERA\Projects\Excel Directory Summarizer'
# Mac
# rootdir = '/Users/chrisrieve/Dropbox/File
# Cabinet/Python/Projects/ExcelDirectorySummarizer'
os.chdir(rootdir)
# Mac
# data_folder = rootdir + '/source'
# Windows
data_folder = rootdir + r'\source'
data_files = []
for subdir, dirs, files in os.walk(data_folder):
    for f in files:
        if f.endswith('.xls'):
            data_files.append(os.path.join(subdir, f))

column_names = [
    ('Workbook', 'TEXT', 'Fill Name'),
    ('# sheets', 'INTEGER', 0),
    ('Sheetname', 'TEXT', 'Fill Sheet Name'),
    ('Rows', 'INTEGER', 0),
    ('Columns', 'INTEGER', 0),
    ('Workbook Size', 'REAL', 0),
    ('Useful(1-5)', 'INTEGER', 0),
    ('Description', 'TEXT', 'Insert Description')
]

# sqlite experimental stuff
sqlite_file = r'excel_info.sqlite'
table_name = r'excel_sheets'

new_field = [row[0] for row in column_names][0]
field_type = [row[1] for row in column_names][0]
default_val = [row[2] for row in column_names][0]

conn = sqlite3.connect(sqlite_file)
c = conn.cursor()
# Dropping table if it exisits
c.execute('DROP TABLE IF EXISTS excel_sheets')
# Creating a new table in sqlite
c.execute('CREATE TABLE {tn} ({nf} {ft})'
          .format(tn=table_name, nf=new_field, ft=field_type))
conn.commit()
conn.close()

# adding additional rows to sqlite db
conn = sqlite3.connect(sqlite_file)
c = conn.cursor()

for col in column_names[1:]:
    new_field = col[0]
    field_type = col[1]
    default_val = col[2]
    c.execute("ALTER TABLE {tn} ADD COLUMN '{cn}' {ct} DEFAULT '{df}'"
              .format(tn=table_name, cn=new_field, ct=field_type, df=default_val))
conn.commit()
conn.close()

for f in data_files:
    print f
    workbook = xlrd.open_workbook(f)
    sheet_names = workbook.sheet_names()
    # Time to add data to the table
    for sheet in sheet_names:
        print sheet
        conn = sqlite3.connect(sqlite_file)
        c = conn.cursor()
        sheet_object = workbook.sheet_by_name(sheet)
        sheet_rows = sheet_object.nrows
        sheet_cols = sheet_object.ncols
        workbook_name = os.path.split(f)[1]
        wb_size = readablebytes.humanize_bytes(os.path.getsize(f))
        num_sheets = len(sheet_names)
        sheet_name = sheet
        c.execute('''INSERT INTO excel_sheets('Workbook','# sheets',
                   'Sheetname','Rows','Columns','Workbook Size') VALUES(?,?,?,?,?,?)''',
                  (workbook_name, num_sheets, sheet_name, sheet_rows, sheet_cols, wb_size))
        conn.commit()
        conn.close()


# Exporting to excel file
conn = sqlite3.connect(sqlite_file)
c = conn.cursor()
workbook = xlsxwriter.Workbook('output/Summary of Files.xlsx')
worksheet = workbook.add_worksheet('Data Summary')
mysel = c.execute("SELECT * FROM excel_sheets")
for i, row in enumerate(mysel):
    print row
    for j, value in enumerate(row):
        worksheet.write(i, j, value)
workbook.close()
