# -*- coding: utf-8 -*-
"""
Created on Thu Jul 14 15:22:58 2016

@author: Christopher.Rieve

Goal: Create a python program template to grab all excel files and various
      information that may prove helpful 
"""

import os 
import pandas as pd 
import sqlite3
import readablebytes

rootdir = r'C:\Users\christopher.rieve\NERA\Projects\Excel Directory Summarizer'
os.chdir(rootdir)
data_folder = rootdir + '\source'
data_files = []
for subdir, dirs, files in os.walk(data_folder):
    for file in files:
        if file.endswith('.xls'):
            data_files.append(os.path.join(subdir,file))

xls = pd.ExcelFile(data_files[0])

column_names = [
('Workbook', 'TEXT', 'Fill Name'),
('# sheets', 'INTEGER', 0),
('Sheetname', 'TEXT', 'Fill Sheet Name'),
('rows', 'INTEGER', 0),
('Workbook size', 'REAL', 0),
('Useful(1-5)', 'INTEGER', 0),
('description', 'TEXT', 'Insert Description')
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
c.execute('CREATE TABLE {tn} ({nf} {ft})'\
        .format (tn=table_name, nf=new_field, ft=field_type))
conn.commit()
conn.close()

# adding additional rows to sqlite db 
conn = sqlite3.connect(sqlite_file)
c = conn.cursor()

i = 0
for col in column_names[1:]:
    new_field = col[0]
    field_type = col[1]
    default_val = col[2]
    c.execute("ALTER TABLE {tn} ADD COLUMN '{cn}' {ct} DEFAULT '{df}'"\
            .format(tn=table_name, cn=new_field, ct=field_type, df=default_val))
conn.commit()
conn.close()    

# Time to add data to the table
conn = sqlite3.connect(sqlite_file)
c = conn.cursor()
for sheet in xls.sheet_names:
    print sheet
    try:
        df = xls.parse(sheet)
        sheet_rows = len(df.index)
#        will need to change data_file[0] to a reference to current wb
        workbook = os.path.split(data_files[0])[1]
        wb_size = readablebytes.humanize_bytes(os.path.getsize(data_files[0]))
        num_sheets = len(xls.sheet_names)
        sheet_name = sheet
        c.execute('''INSERT INTO excel_sheets([c[0] for c in column_names[:5]])
                    VALUES(?,?,?,?,?)''',(workbook, num_sheets, sheet_name,
                    sheet_rows, wb_size))
    except:
        continue
conn.commit()
conn.close()
    
xls.sheet_names

named_sheets = xls.sheet_names
num_sheets = 0
for sheet in named_sheets:
    print sheet
    try:    
        df = xls.parse(sheet)
        num_sheets += 1
    except:
        continue
print num_sheets 