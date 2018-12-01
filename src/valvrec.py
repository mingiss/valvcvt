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
    '''Element of the ValvRecTree.in_data'''
    def __init__(self):
        self.value = ''
        self.is_heading = False
        self.colspan = 0
        self.rowspan = 0

class SegHeadingCol:
    '''One column of data segment headings of certain attribute type'''

    # attributes of classification in order their titles appear in front of the output rows
    # values of the dictionary -- lists of keywords for attribute class recognition
    class_attribs = \
    { \
        'Kategorie':                [], \
        'Familie':                  [], \

         # keywords of this attribute should be exact values of all possible materials,
         # because in some files they are provided as the suffices of worksheet names
         # should be arranged in increasing ambiguity order
        'Material':                 ['Stainless Steel', 'Steel', 'Brass'], \

        'Bauform':                  [], \
        'Serie/Verbindungstyp':     [], \
        'Metrisch/UNC':             [] \
    }

    def __init__(self):
        self.attrib = '' # key to class_attribs when recognized
        self.values = [] # heading values itself, the amount should correspond to the height of data segment

class DataSeg:
    pat_wdt = 3 # the width of data segment pattern to be searched

    seg_head = ['Type', 'PN [bar]', 'filename [.stp]']

    def __init__(self):
        self.xx = 0
        self.yy = 0
        self.length = 1 # number of rows in the segment
        self.headings = [] # list of SegHeadingCol's

    def location(self):
        '''Formats segment coordinates in Excell notation'''
        shift = len(SegHeadingCol.class_attribs)
        return (chr(ord('A') + self.xx - shift) + str(self.yy + 1) + ':' + chr(ord('A') + self.xx + DataSeg.pat_wdt - 1 - shift) + str(self.yy + self.length))


# ----------------------------------
class ValvRecTree(XlsTree):
    '''xlsx.xml tools localized to valvrec'''

    def __init__(self):
        self.out_data = []   # array of output rows -- arrays of string cells


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

        # align rows to have equal lengths
        self.in_data.append([]) # appending empty row for data pattern search not to exceed amount of rows
        max_row_len = self.calc_max_row_len()
        for row_data in self.in_data:
            for ii in range(len(row_data), max_row_len + \
                    # additional columns at the end for data pattern being searched not to exceed the lengths
                    DataSeg.pat_wdt - 1):
                row_data.append(InCellValue())
            # additional starting columns for attribute heading searching in case of absence of them
            for ii in range(0, len(SegHeadingCol.class_attribs)):
                row_data.insert(0, InCellValue())

        # print('----------------------------------')
        # for row_data in self.in_data:
        #    for cell_data in row_data:
        #        print(cell_data.value + ',\t', end = '')
        #        print(str(cell_data.colspan) + ',\t', end = '')
        #        print(str(cell_data.rowspan) + ',\t', end = '')
        #        print(str(cell_data.is_heading) + ',\t', end = '')
        #    print()


    def search_data_pattern(self, prev_seg):
        '''
        Searches for the next 3 x N data segment
        Current file table data should be read to self.in_data using scan_in_data()
        parameter prev_seg -- previous data segment of the type DataSeg
        returns newly found DataSeg object or None in case of the last one
        '''

        max_row_len = self.calc_max_row_len()
        for col_ix in range (prev_seg.xx, max_row_len):
            for row_ix in range(prev_seg.yy + prev_seg.length, len(self.in_data)):
                in_cell = self.in_data[row_ix][col_ix]
                if (in_cell.value and (not in_cell.is_heading)):
                    new_seg = DataSeg()
                    new_seg.xx = col_ix
                    new_seg.yy = row_ix

                    for seg_row_ix in range(row_ix + 1, len(self.in_data)):
                        found = False
                        for ii in range(0, DataSeg.pat_wdt):
                            next_cell = self.in_data[seg_row_ix][col_ix + ii]
                            if (next_cell.value and (not next_cell.is_heading)):
                                found = True
                                break
                        if (found):
                            new_seg.length += 1
                        else:
                            break

                    has_head = True
                    for ii in range(0, DataSeg.pat_wdt):
                        if (self.in_data[row_ix][col_ix + ii].value != DataSeg.seg_head[ii]):
                            has_head = False
                            break
                    if (has_head):
                        new_seg.yy += 1
                        new_seg.length -= 1
                    else:
                        print ('Error: Data segment {} has no heading {}'.format(new_seg.location(), DataSeg.seg_head))

                    # extracting headings
                    for ii in range(0, len(SegHeadingCol.class_attribs)):
                        cur_head = SegHeadingCol()
                        found = False
                        for seg_row_ix in range(new_seg.yy, new_seg.yy + new_seg.length):
                            cur_head.values.append(self.in_data[seg_row_ix][col_ix - 1 - ii].value)
                            found = (found or self.in_data[seg_row_ix][col_ix - 1 - ii].is_heading)
                        if (found):
                            new_seg.headings.append(cur_head)
                        else:
                            break

                    return new_seg

        return None


    def process_table(self, table):
        '''
        processing of one input worksheet table
        table -- lxml.etree node, containing ordinary input worksheet table
        '''

        self.scan_in_data(table)

        data_seg = DataSeg()
        while (data_seg):
            data_seg = self.search_data_pattern(data_seg)
            if (data_seg):
                print (data_seg.location())
                for head in data_seg.headings:
                    print(head.values)


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
