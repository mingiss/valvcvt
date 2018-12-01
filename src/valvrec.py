#!/usr/bin/python
# coding=UTF-8

"""
valvrec.py

Extracts data segments from multiple `Microsoft Excel 2003 XML` files in particular valve classificator format
without binding them to particular coordinates on the worksheet.

Using:
    python valvrec.py input.xml output.xml
"""

__author__ = "Mindaugas Piešina"
__version__ = "0.0.1"
__email__ = "mpiesina@netscape.net"
__status__ = "Prototype"

import sys
import copy

try:
    from lxml import etree
except ImportError:
    print ('no lxml')
    import xml.etree.ElementTree as etree

from xlstree import XlsTree


# ----------------------------------
class InCellValue:
    def __init__(self):
        self.value = ''
        self.is_heading = False
        self.colspan = 0
        self.rowspan = 0

# ----------------------------------
class ValvRecTree(XlsTree):
    '''xlsx.xml tools localized to valvrec'''

    materials = ['Stainless Steel', 'Steel', 'Brass']

    pat_wdt = 3 # the width of data pattern to be searched

    def __init__(self):

        out_data = []   # array of output rows -- arrays of string cells


    def calc_max_row_len(self):
        # calculate max row length
        max_row_len = 0
        for row_data in self.in_data:
            row_len = len(row_data)
            if (row_len > max_row_len):
                max_row_len = row_len
        return max_row_len


    def scan_in_data(self, table):
        '''scanning input data from table node to self.in_data'''

        # array of input rows -- arrays of tupplets, each conforming to layout InCellLayout
        self.in_data = []

        for row in table.xpath('xmlns:Row', namespaces = XlsTree.ns_xsl):
            row_data = []
            for cell in row.xpath('xmlns:Cell', namespaces = XlsTree.ns_xsl):
                cell_data = InCellValue()
                cell_data.value = ''.join(cell.xpath('.//text()'))
                colsp = cell.get(XlsTree.ns_pref + 'MergeAcross')
                if (colsp):
                    cell_data.colspan = int(colsp)
                if (cell_data.colspan > 0):
                    cell_data.is_heading = True
                rowsp = cell.get(XlsTree.ns_pref + 'MergeDown')
                if (rowsp):
                    cell_data.rowspan = int(rowsp)
                row_data.append(cell_data)
            self.in_data.append(row_data)

        # insert colspans
        for row_data in self.in_data:
            col_ix = 0
            while (col_ix < len(row_data)):
                cell_data = row_data[col_ix]
                if (cell_data.colspan > 0):
                    for ii in range(0, cell_data.colspan):
                        row_data.insert(col_ix + 1, InCellValue())
                    cell_data.colspan = 0
                col_ix = col_ix + 1

        # spread headings through rowspans
        max_row_len = self.calc_max_row_len()
        col_ix = 0
        while (col_ix < max_row_len):
            for row_ix in range(0, len(self.in_data)):
                if (len(self.in_data[row_ix]) > col_ix):
                    cell_data = self.in_data[row_ix][col_ix]
                    if (cell_data.rowspan > 0):
                        new_cell = InCellValue()
                        src_row_ix = row_ix - 1
                        if ((src_row_ix >= 0) and (len(self.in_data[src_row_ix]) > col_ix)):
                            new_cell = copy.copy(self.in_data[src_row_ix][col_ix])
                        new_cell.colspan = 0
                        new_cell.rowspan = 0
                        for ii in range(0, cell_data.rowspan):
                            new_row_ix = row_ix + ii + 1
                            if (len(self.in_data) > new_row_ix):

                                # appending empty cells in case, if len(self.in_data[new_row_ix]) < col_ix
                                for ii in range(len(self.in_data[new_row_ix]), col_ix):
                                    self.in_data[new_row_ix].append(InCellValue())

                                self.in_data[new_row_ix].insert(col_ix, copy.copy(new_cell))
                        self.in_data[row_ix][col_ix] = copy.copy(new_cell)
                        max_row_len = self.calc_max_row_len()
                        cell_data.rowspan = 0
            col_ix = col_ix + 1

        # align rows to have equal lengths and
        # additional columns at the end for data pattern being searched not to exceed the lengths
        max_row_len = self.calc_max_row_len()
        for row_data in self.in_data:
            for ii in range(len(row_data), max_row_len + ValvRecTree.pat_wdt - 1):
                row_data.append(InCellValue())


        print('----------------------------------')
        for row_data in self.in_data:
            for cell_data in row_data:
                print(cell_data.value + ',\t', end = '')
                # print(str(cell_data.colspan) + ',\t', end = '')
                # print(str(cell_data.rowspan) + ',\t', end = '')
                # print(str(cell_data.is_heading) + ',\t', end = '')
            print()


    def process_table(self, table):
        '''
        processing of one input worksheet table
        table -- lxml.etree node, containing ordinary input worksheet table
        '''

        self.scan_in_data(table)


    def process_in_file(self, in_fname):
        '''processing of one xlsx file'''

        if (not self.load(in_fname)):
            print("Error: " + tree.last_error)
            sys.exit(1)
        print('Processing file: ' + self.fname)

        self.trim()

        for table in self.dom.xpath('//xmlns:Table', namespaces = XlsTree.ns_xsl):
            self.process_table(table)


    def write_csv(self, out_fname):
        pass


# ----------------------------------
def main():

    if (len(sys.argv) < 3):
        print('Error: Give input and output file names as parameters')
        sys.exit(2)

    in_flist_fname = sys.argv[1]
    out_fname = sys.argv[2]

    try:
        with open(in_flist_fname) as flist:
            in_fnames = flist.read().splitlines()
    except OSError as err:
        self.last_error = 'Unable to open file %s (%s)' % (in_flist, err)
        sys.exit(1)

    # ----------------------------------
    tree = ValvRecTree()

    for in_fname in in_fnames:
        if (in_fname != 'orig/README.txt'):
            tree.process_in_file(in_fname)

    tree.write_csv(out_fname)


if __name__ == "__main__":
    main()
