#!/usr/bin/python
# coding=UTF-8

"""
valvcvt.py

Concatenates tables of multiple `Microsoft Excel 2003 XML` files in particular valve classificator format
Spreads minor and major headings, sheet and file names to the start of the exported rows.

Using:
    python valvcvt.py input_list.txt output.xml csv

        input_list.txt -- text file with the list of input xml file names

        the third optional parameter -- output format
            at the moment could be `csv` (default) or `xml` (the latter for files in `Microsoft Excel 2003 XML` format)
"""

__author__ = "Mindaugas Piešina"
__version__ = "0.0.1"
__email__ = "mpiesina@netscape.net"
__status__ = "Prototype"

import sys
import codecs

try:
    from lxml import etree
except ImportError:
    print ('no lxml')
    import xml.etree.ElementTree as etree

from xlstree import XlsTree

# ----------------------------------
class valvtree(XlsTree):
    '''xlsx.xml tools localized to valvcvt'''

    out_fmt_selector = \
    { \
        'csv': (lambda self, out_fname, delim: XlsTree.export_csv(self, out_fname, delim)), \
        'xml': (lambda self, out_fname, delim: XlsTree.write(self, out_fname)) \
    }

    materials = ['Stainless Steel', 'Steel', 'Brass']

    def del_hats(self):
        '''removes heading rows after the last spanned heading'''
        for tab in self.dom.xpath('//xmlns:Table', namespaces = XlsTree.ns_xsl):
            prev_spanned = False
            for row in tab.xpath('xmlns:Row', namespaces = XlsTree.ns_xsl):
                row_spanned = False
                for cell in row.xpath('xmlns:Cell', namespaces = XlsTree.ns_xsl):
                    try:
                        span = int(cell.get(XlsTree.ns_pref + 'MergeAcross'))
                        if (span > 1):
                            row_spanned = True
                    except:
                        pass
                if (prev_spanned and (not row_spanned)):
                    # previous row had spanned cells -- this row is a heading -- deleting
                    print(self.fname + ': Removed heading: ' + ''.join(row.xpath('.//text()')))
                    row.getparent().remove(row)
                prev_spanned = row_spanned

    def insert_heads(self):
        '''inserts first level heads, if absent'''
        for tab in self.dom.xpath('//xmlns:Table', namespaces = XlsTree.ns_xsl):
            second_spanned = False
            for row in tab.xpath('xmlns:Row', namespaces = XlsTree.ns_xsl):
                cells = row.xpath('xmlns:Cell', namespaces = XlsTree.ns_xsl)
                if (len(cells) > 1):
                    cell = cells[1]
                    try:
                        span = int(cell.get(XlsTree.ns_pref + 'MergeAcross'))
                        if (span > 1):
                            second_spanned = True
                            break
                    except:
                        pass
            if (not second_spanned):
                # there were no second level headings -- the whole table is shifted left
                # for instance, in file 2-way ball valves flangeable with SAE connections.xlsx.xml (sheet KH-SAE Steel)
                # file Cartridge ball valves.xlsx.xml has no headings at all (the method should be applyed twice)
                # just moving the table to right by one cell
                for row in tab.xpath('xmlns:Row', namespaces = XlsTree.ns_xsl):
                    new_cell = etree.Element(XlsTree.ns_pref + 'Cell')
                    row.insert(0, new_cell)


    def spread_heads(self):
        '''spreads second level headings to first column of each section row'''
        for tab in self.dom.xpath('//xmlns:Table', namespaces = XlsTree.ns_xsl):
            heading = ''
            for row in tab.xpath('xmlns:Row', namespaces = XlsTree.ns_xsl):
                second_spanned = False
                cells = row.xpath('xmlns:Cell', namespaces = XlsTree.ns_xsl)
                if (len(cells) > 1):
                    cell = cells[1]
                    try:
                        span = int(cell.get(XlsTree.ns_pref + 'MergeAcross'))
                        if (span > 1):
                            cell_text = ''.join(cell.xpath('.//text()'))
                            heading = ''.join(cell.xpath('.//text()'))
                            second_spanned = True
                    except:
                        pass
                if (second_spanned):
                    row.getparent().remove(row)
                else:
                    new_cell = etree.Element(XlsTree.ns_pref + 'Cell')
                    cell_data = etree.Element(XlsTree.ns_pref + 'Data')
                    cell_data.set(XlsTree.ns_pref + 'Type', 'String')
                    cell_data.text = heading
                    new_cell.append(cell_data)
                    row.insert(0, new_cell)

    def spread_sheet_heads(self):
        '''
        spreads sheet headings to first column of each row in the sheet and
        material name parsed out of them into the second
        '''
        for ws in self.dom.xpath('//xmlns:Worksheet', namespaces = XlsTree.ns_xsl):
            heading = ws.get(XlsTree.ns_pref + 'Name')
            material = ''
            for mat in valvtree.materials:
                if (mat in heading):
                    material = mat
                    break
            for tab in ws.xpath('xmlns:Table', namespaces = XlsTree.ns_xsl):
                for row in tab.xpath('xmlns:Row', namespaces = XlsTree.ns_xsl):
                    new_cell = etree.Element(XlsTree.ns_pref + 'Cell')
                    cell_data = etree.Element(XlsTree.ns_pref + 'Data')
                    cell_data.set(XlsTree.ns_pref + 'Type', 'String')
                    cell_data.text = heading
                    new_cell.append(cell_data)
                    row.insert(0, new_cell)
                    new_cell = etree.Element(XlsTree.ns_pref + 'Cell')
                    cell_data = etree.Element(XlsTree.ns_pref + 'Data')
                    cell_data.set(XlsTree.ns_pref + 'Type', 'String')
                    cell_data.text = material
                    new_cell.append(cell_data)
                    row.insert(1, new_cell)

    def spread_fname(self):
        '''spreads file name to first column of each row in the sheet'''
        for tab in self.dom.xpath('//xmlns:Table', namespaces = XlsTree.ns_xsl):
            for row in tab.xpath('xmlns:Row', namespaces = XlsTree.ns_xsl):
                new_cell = etree.Element(XlsTree.ns_pref + 'Cell')
                cell_data = etree.Element(XlsTree.ns_pref + 'Data')
                cell_data.set(XlsTree.ns_pref + 'Type', 'String')
                cell_data.text = self.fname
                new_cell.append(cell_data)
                row.insert(0, new_cell)

    def process_valv(self):
        '''all headings processing of one xlsx file'''
        self.trim()

        # self.insert_heads() # for tables with second level headings solely
        # self.insert_heads() # for tables without headings
        self.del_hats()
        self.spread_heads() # spread second level heads to the relevant groups of rows
        self.spread_heads() # first level heads at the moment are shifted to the right as if being second level
        self.spread_sheet_heads()
        self.spread_fname()

        self.append_xlsx_sheet()


# ----------------------------------
def main():

    if (len(sys.argv) < 3):
        print('Error: Give input list and output file names as parameters')
        sys.exit(2)

    in_flist_fname = sys.argv[1]
    out_fname = sys.argv[2]
    out_fmt = 'csv'
    if (len(sys.argv) > 3):
        out_fmt = sys.argv[3]
        if (not out_fmt in valvtree.out_fmt_selector.keys()):
            print('Error: Unknown output file format: ' + out_fmt)
            sys.exit(2)

    try:
        with open(in_flist_fname) as flist:
            in_fnames = flist.read().splitlines()
    except OSError as err:
        self.last_error = 'Unable to open file %s (%s)' % (in_flist, err)
        sys.exit(1)

    # ----------------------------------
    tree = valvtree()

    if (not tree.load(in_fnames[0])):
        print("Error: " + tree.last_error)
        sys.exit(1)

    tree.process_valv()

    del in_fnames[0]
    for in_fname in in_fnames:
        if (in_fname != 'orig/README.txt'):
            add_tree = valvtree()

            if (not add_tree.load(in_fname)):
                print("Error: " + add_tree.last_error)
                sys.exit(1)

            add_tree.process_valv()

            tree.append_xlsx(add_tree)

    # ----------------------------------
    if (not valvtree.out_fmt_selector[out_fmt](tree, out_fname, ',')):
        print("Error: " + tree.last_error)
        sys.exit(1)

if __name__ == "__main__":
    main()
