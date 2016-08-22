#!/usr/bin/env python

from collections import OrderedDict

from openpyxl import load_workbook
from openpyxl.styles import Font


class Excelerator(object):

    def __init__(self, file = None):
        # Set style variables for later use.
        self.bold = Font(bold=True)

        # If provided, parse file immediately.
        if file:
            self.excelerate(file)

    def append_data(self, sheet, data):
        sorted_data = sorted(data, key=lambda k: k['PART NUMBER'])

        sheet.append(sorted_data[0].keys())

        for row in sorted_data:
            sheet.append(row.values())

    def apply_styles(self):
        for i in range(1, len(self.workbook.worksheets)):
            dims = {}
            ws = self.workbook.worksheets[i]

            for cell in ws.rows[0]:
                cell.font = self.bold

            for j, row in enumerate(ws.rows):
                padding = 2 if j == 0 else 1

                for cell in row:
                    if cell.value:
                        dims[cell.column] = max((dims.get(cell.column, 0),
                            len(str(cell.value)) + padding, 4))

            for col, value in dims.items():
                ws.column_dimensions[col].width = value

    def create_headers_list(self, source_cells):
        row_number = str(self.find_first_row() + 1)
        headers_list = self.master_parts_sheet.iter_rows(
            'A' + row_number + ':Z' + row_number).next()

        return headers_list

    def create_parts_list(self, initial_row = 0):
        first_row_number = self.find_first_row(initial_row)
        last_row_number = self.find_last_row(first_row_number)

        first_row = 'A' + str(first_row_number + 1)
        last_row = 'Z' + str(last_row_number + 1)

        parts_list = list(
            self.master_parts_sheet.iter_rows(first_row + ':' + last_row))

        return parts_list, last_row_number

    def create_section_list(self, section, part_group):
        headers = [i.value for i in self.headers if i.value != None]

        def item(i, j):
            if j == len(headers):
                attr = ('PART GROUP', part_group)
            else:
                attr = (self.headers[j].value, section[i][j].value)

            return attr

        return [OrderedDict(item(i, j) for j in range(len(headers) + 1))
            for i in range(1, len(section))]

    def create_sheet(self, name):
        sheet = self.workbook.create_sheet()
        sheet.title = name

        return sheet

    def find_first_row(self, row = 0):
        cell_value = str()

        while str(cell_value) != 'QTY':
            cell_value = self.master_parts_sheet.rows[row][0].value
            row += 1

        return row - 1

    def find_last_row(self, row):
        cell_value = str()

        while cell_value != None:
            cell_value = self.master_parts_sheet.rows[row][0].value
            row += 1

        return row - 2

    def excelerate(self, file):
        # Load spreadsheet into Workbook object.
        self.workbook = load_workbook(file)

        # First spreadsheet should contain master parts list.
        self.master_parts_sheet = self.workbook.active

        # Iterate through master parts list and identify each section.
        fabricated_parts, last_row_number = self.create_parts_list()
        weldments, last_row_number = self.create_parts_list(
            last_row_number + 1)
        purchased_parts, last_row_number = self.create_parts_list(
            last_row_number + 1)

        # Create master list of headers.
        self.headers = self.create_headers_list(fabricated_parts)

        # Create lists of dictionarys for each section.
        fabricated_list = self.create_section_list(
            fabricated_parts, 'FabricatedParts')
        weldments_list = self.create_section_list(
            weldments, 'WeldmentParts')
        purchased_list = self.create_section_list(
            purchased_parts, 'PurchasedFabParts')

        master_parts_list = fabricated_list + weldments_list + purchased_list

        # Create Weld SFC Pick List sheet.
        weld_picklist_sheet = self.create_sheet('WELD SCF Pick List')
        self.append_data(weld_picklist_sheet, fabricated_list)

        # Create WELD BOM sheet.
        weld_bom_sheet = self.create_sheet('WELD BOM')
        weld_bom_data = [x for x in master_parts_list if str(x['WELDED']) ==
            'WELDED' and str(x['WELDMENT USED']) != 'SHIPPED LOOSE']
        self.append_data(weld_bom_sheet, weld_bom_data)

        # Create WELD LOOSE sheet.
        weld_loose_sheet = self.create_sheet('WELD LOOSE')
        weld_loose_data = [x for x in master_parts_list if str(x['WELDED']) ==
            'WELDED' and str(x['WELDMENT USED']) == 'SHIPPED LOOSE']
        self.append_data(weld_loose_sheet, weld_loose_data)

        # Create WELD Packing Slip sheet.
        weld_packing_sheet = self.create_sheet('WELD Packing Slip')
        weld_packing_data = [x for x in master_parts_list
            if str(x['WELDMENT USED']) == 'SHIPPED LOOSE']
        self.append_data(weld_packing_sheet, weld_packing_data)

        # Create FINISH Pick List sheet.
        finish_picklist_sheet = self.create_sheet('FINISH Pick List')
        finish_picklist_data = [x for x in master_parts_list
            if str(x['WELDMENT USED']) == 'SHIPPED LOOSE']
        self.append_data(finish_picklist_sheet, finish_picklist_data)

        # Apply styles.
        self.apply_styles()

        self.workbook.save("test_complete.xlsx")


#parser = Excelerator()
#parser.excelerate('test.xlsx')
Excelerator('test.xlsx')
