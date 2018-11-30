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
        self.colspan = 1
        self.rowspan = 1

# ----------------------------------
class ValvRecTree(XlsTree):
    '''xlsx.xml tools localized to valvrec'''

    materials = ['Stainless Steel', 'Steel', 'Brass']

    def __init__(self):

        out_data = []   # array of output rows -- arrays of string cells


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
                if (cell_data.colspan > 1):
                    cell_data.is_heading = True
                rowsp = cell.get(XlsTree.ns_pref + 'MergeDown')
                if (rowsp):
                    cell_data.rowspan = int(rowsp)
                row_data.append(cell_data)
            self.in_data.append(row_data)

        # insert colspans
        for row_data in self.in_data:
            for cell_data in row_data:
                for ii in range(1, cell_data.colspan):
                    row_data.insert(row_data.index(cell_data) + 1, InCellValue())
                cell_data.colspan = 1


        print('----------------------------------')
        for row_data in self.in_data:
            for cell_data in row_data:
                # print(cell_data.value + ',\t', end = '')
                print(str(cell_data.colspan) + ',\t', end = '')
                # print(str(cell_data.rowspan) + ',\t', end = '')
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