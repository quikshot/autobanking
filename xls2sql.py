#
#  Input xls file and insert into sql database
# 
#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlrd
import sqlite3 as lite
import sys
import string
from optparse import OptionParser
import re


parser = OptionParser()
parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

(options, args) = parser.parse_args()

kutxaInitialRow=8

#TODO: Refactor enum:
kutxaDate=0
kutxaConcept=1
kutxaRealDate=2
kutxaQuantity=3
kutxaBalance=4

#to something like:
kutxaRows= ['date','concept','realDate','currentBalance']


con = None

#----------------------------------------------------------------------
def process_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
 
    # print number of sheets
    #print book.nsheets
 
    # print sheet names
    #print book.sheet_names()
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
 
    # read a row
    #print first_sheet.row_values(0)
 
    # read a cell
    #cell = first_sheet.cell(0,0)
    #print cell
    #print cell.value
 
    emptyRow=first_sheet.cell(kutxaInitialRow,0)
    print "row=",emptyRow
    row=kutxaInitialRow
    print "number of entries: ",first_sheet.nrows-kutxaInitialRow
    con = lite.connect('test.db')

    with con:
    
        cur = con.cursor()
        cur.execute('SELECT SQLITE_VERSION()')
        data = cur.fetchone()
        #print "SQLite version: %s" % data

        #Create Database
        cur.execute("CREATE TABLE IF NOT EXISTS kutxapersonal(Date Text, Concept TEXT, RealDate TEXT, Quantity REAL not null, Balance REAL not null, Cathegory TEXT, Ignore INT, unique(Date, Quantity,Balance) ON CONFLICT IGNORE )")

    while row < first_sheet.nrows:
        # read a row slice
        entryDate     = first_sheet.cell(row, kutxaDate).value
        entryConcept  = first_sheet.cell(row, kutxaConcept).value
        entryRealDate = first_sheet.cell(row, kutxaRealDate).value
        entryQuantity = first_sheet.cell(row, kutxaQuantity).value
        entryBalance  = first_sheet.cell(row, kutxaBalance).value
        row=row+1

        #build insert query
        sqlquery="INSERT INTO kutxapersonal VALUES ('"
        sqlquery += entryDate
        sqlquery += "', '"
        sqlquery += entryConcept.replace("'"," ")
        sqlquery += "', '"
        sqlquery += entryRealDate 
        sqlquery += "', "
        sqlquery += str(entryQuantity)
        sqlquery += ", "
        sqlquery += str(entryBalance)
        sqlquery += ", 'None', 0)"

        #print sqlquery
        cur.execute(sqlquery)

    con.commit()
    con.close()
    print "Processed ",row-kutxaInitialRow," entries from bank. Non duplicates added to database"

#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "./kutxa.xls"
    process_file(options.filename)


