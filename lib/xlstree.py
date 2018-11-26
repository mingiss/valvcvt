#!/usr/bin/python
# coding=UTF-8

"""
xlstree.py

Class to work with from OpenOffice Calc exported xml's
"""

__author__ = "Mindaugas Piešina"
__version__ = "0.0.1"
__email__ = "mpiesina@netscape.net"
__status__ = "Prototype"

try:
    from lxml import etree
except ImportError:
    print ('no lxml')
    import xml.etree.ElementTree as etree

# from lxml import objectify

ns = 'urn:schemas-microsoft-com:office:spreadsheet'
ns_pref = '{%s}' % ns
ns_xsl = {'xmlns': ns}

class xlstree:
    '''xlsx.xml tools class'''

    def __init__(self):
        self.last_error = ''
        self.fname = ''

        # objectify.deannotate(new_dom, cleanup_namespaces = True)


    def load(self, in_fname):
        '''loads self.dom from xml file, returns True in case of success'''

        self.fname = in_fname

        try:
            with open(in_fname, 'rb') as in_file:
                in_data = in_file.read()
        except OSError as err:
            self.last_error = 'Unable to open file "%s" (%s)' % (self.fname, err)
            return(False)

        try:
            self.dom = etree.fromstring(in_data)
        except Exception as err:
            self.last_error = 'Bad input file "%s" format (%s)' % (self.fname, err)
            return(False)

        return(True)


    def write(self, out_fname):
        '''writes self.dom to xml file, returns True in case of success'''

        try:
            out_data = etree.tostring(self.dom).decode()
        except Exception as err:
            self.last_error = 'Unable to convert data (%s)' % err
            return(False)

        try:
            with open(out_fname, 'w') as out_file:
                out_file.write('<?xml version="1.0" encoding="UTF-8"?>\n')
                out_file.write('<?mso-application progid="Excel.Sheet"?>')
                out_file.write(out_data)
        except Exception as err:
            self.last_error = 'Unable to write file %s (%s)' % (out_fname, err)
            return(False)

        return(True)


    def trim(self):
        '''delete empty cells of self.dom'''

        # delete empty cells at the ends of the rows
        for row in self.dom.xpath('//xmlns:Row', namespaces = ns_xsl):
            for cell in reversed(row.xpath('xmlns:Cell', namespaces = ns_xsl)):
                if (cell.tail or cell.text or list(cell)):
                    break
                row.remove(cell)

        # delete empty rows at the ends of the tables
        for tab in self.dom.xpath('//xmlns:Table', namespaces = ns_xsl):
            for row in reversed(tab.xpath('xmlns:Row', namespaces = ns_xsl)):
                if (row.tail or row.text or list(row)):
                    break
                tab.remove(row)

        # collect dictionary of used styles
        used_styles = {}
        for cell in self.dom.xpath('//xmlns:Cell', namespaces = ns_xsl):
            used_styles[cell.get(ns_pref + 'StyleID')] = True

        #for nod in self.dom.xpath('//*'):
        #    if (nod.get(ns_pref + 'StyleID') and (nod.tag != ns_pref + 'Cell')):
        #        print(nod.tag + ' ' + nod.get(ns_pref + 'StyleID'))

        # delete unused styles
        for sty in self.dom.xpath('//xmlns:Style', namespaces = ns_xsl):
            if (sty.get(ns_pref + 'ID') not in used_styles.keys()):
                sty.getparent().remove(sty)


    def del_empty_tables_sheets(self):

        for tab in self.dom.xpath('//xmlns:Table', namespaces = ns_xsl):
            if not (tab.tail or tab.text or tab.xpath('xmlns:Row', namespaces = ns_xsl)):
                tab.getparent().remove(tab)

        for sheet in self.dom.xpath('//xmlns:Worksheet', namespaces = ns_xsl):
            if not (sheet.tail or sheet.text or sheet.xpath('xmlns:Table', namespaces = ns_xsl)):
                sheet.getparent().remove(sheet)


    def append_xlsx_sheet(self):
        """ Add worksheets together"""

        tables = self.dom.xpath('//xmlns:Table', namespaces = ns_xsl)
        main = tables[0]
        del tables[0]

        for tbl in tables:
             for row in tbl.xpath('xmlns:Row', namespaces = ns_xsl):
                main.append(row)

        self.del_empty_tables_sheets()


    def append_xlsx(self, add_tree):
        """Add tables from seperate files together"""

        tables = self.dom.xpath('//xmlns:Table', namespaces = ns_xsl)
        main = tables[0]

        add_tree_tables = add_tree.dom.xpath('//xmlns:Table', namespaces = ns_xsl)
        table = add_tree_tables[0]

        for row in table.xpath('xmlns:Row', namespaces = ns_xsl):
            main.append(row)
