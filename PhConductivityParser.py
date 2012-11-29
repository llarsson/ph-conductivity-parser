#!/usr/bin/env python
#
# Parses pH and conductivity records from a particular proprietary 
# format and outputs it in Microsoft Excel format.
#
# Author: Lars Larsson (larsson.work@gmail.com)
# Copyright 2012
#

import xlwt
import csv
import copy
import sys

import UnicodeCSVTools

class Record:
    """A record from the type of file output by the equipment"""
    def __init__(self):
        """Creates a new instance with all fields set to None"""
        self.number = None
        self.comment = None
        self.start_time = None
        self.rx = None
        self.result = None
        self.unit = None
        self.name = None

    def as_list(self):
        """Represents this record as a list of values in the order 
        expected in the target file"""
        return [
                self.number,
                self.comment,
                self.start_time,
                self.rx,
                # result is a number, Excel needs it treated as such
                # warning: this could introduce subtle bugs, due to
                # handling of floats
                float(self.result),
                self.unit,
                self.name
                ]

def excel_exporter(records, filename):
    """Exports the list of records as an Excel file named filename."""

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet 1")

    # Styles for cells
    general = xlwt.easyxf()
    number = xlwt.easyxf(num_format_str='0.000')

    # Very fragile mapping between cell headers and contents formatting
    headers = ["No", "Comment / ID", "Start time", "Rx", "Result", 
            "Unit", "Name", "LabID", "File"]
    types = [general, general, general, general, number, general, general, number]

    # Write out the headers
    for column, header in enumerate(headers):
        sheet.write(0, column, header)

    lab_id = 1 # record counter
    for row, record in enumerate(records):
        # We will have to add +1 to the row index due to the headers
        for column, value in enumerate(record.as_list()):
            sheet.write(row+1, column, value, types[column])
        sheet.write(row+1, 7, lab_id, number)
        sheet.write(row+1, 8, xlwt.Formula("MID(CELL(\"filename\"),SEARCH(\"[\",CELL(\"filename\"))+1,SEARCH(\"]\",CELL(\"filename\"))-SEARCH(\"[\",CELL(\"filename\"))-1)))"))
        lab_id += 1
    workbook.save(filename)

def record_from_row(row, base_record):
    """Creates a new record based on a base record from the given row. 
    Fields are first copied from the base record and then given unique 
    values found in the row, if there are such values."""

    record = copy.copy(base_record)

    if row[0]:
        record.number = row[0]
    if row[1]:
        record.comment = row[1]
    if row[2]:
        record.start_time = row[2]
    if row[3]:
        record.rx = row[3]
    if row[4]:
        record.result = row[4]
    if row[5]:
        record.unit = row[5]
    if row[6]:
        record.name = row[6]

    return record

def records_from_file(filename):
    """Reads records from the specified file, returning a list of Record 
    elements."""

    with open(filename) as csvfile:
        record_reader = UnicodeCSVTools.UnicodeReader(csvfile, encoding="utf-16", delimiter='\t')
        record_reader.next() # skip first line, it just has headers
        records = []
        previous_record = Record() # empty record at first, will be overwritten
        for row in record_reader:
            if not row:
                break # skip empty rows (there can be one at the end of the file)
            record = record_from_row(row, previous_record)
            records.append(record)
            previous_record = record
        return records

if __name__ == "__main__":
    files = sys.argv[1:]
    if not files:
        print "Usage: list the names of all the files to convert as parameters."
        sys.exit(1)
    for file in files:
        records = records_from_file(file)
        excel_exporter(records, file + ".xls")
