import copy
import os
from collections import OrderedDict
from datetime import datetime
from io import BytesIO

import xlrd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.writer.excel import save_virtual_workbook


class Excelerator(object):

    def __init__(self, workbook=str(), multiplier=5):
        self.filename = workbook
        self.multiplier = multiplier

        if not isinstance(self.filename, str):
            self.filename = workbook.filename

        self.filename_stripped, self.extension = os.path.splitext(self.filename)

        # Set style variables for later use.
        self.side = Side(border_style='thin')

        self.bold = Font(bold=True)
        self.border = Border(left=self.side, right=self.side,
                             top=self.side, bottom=self.side)
        self.date_format = 'mm/dd/yy'
        self.title_font = Font(size=18)

        # If provided, parse file immediately.
        if self.filename:
            self.excelerate(workbook)

    def add_column(self, name, parts_list, last=False):
        for part in parts_list:
            part.update({name: None})
            part.move_to_end(name, last=last)

    def append_data(self, data, sheet):
        sorted_data = sorted(data, key=lambda k: k['PART NUMBER'])

        # Append dictionary keys as spreadsheet headers.
        sheet.append(list(sorted_data[0]))

        for row in sorted_data:
            sheet.append(list(row.values()))

    def append_empty_row(self, sheet, number=1):
        for i in range(number):
            sheet.append([str()])

    def append_signature(self, prompt, sheet, date=True):
        signature_line = ': ___________________________'
        signature = str(prompt) + signature_line

        if date:
            signature += '           ' + 'Date' + signature_line

        self.append_empty_row(sheet, 2)
        sheet.append([None, signature])

        # Style signature cells.
        signature_row = sheet.max_row
        sheet.merge_cells(start_row=signature_row, start_column=2,
                          end_row=signature_row, end_column=5)

    def append_title(self, sheet, title=None):
        if not title:
            title_components = [
                self.filename_stripped.upper(),
                '–',
                sheet.title,
                '({n}x)'.format(n=self.multiplier)
            ]
            title = ' '.join(title_components)

        sheet.append([title])

        # Style sheet title.
        title_row = sheet.max_row
        sheet.merge_cells(start_row=title_row, start_column=1,
                          end_row=title_row, end_column=4)
        title_cell = sheet.cell(column=1, row=title_row)
        title_cell.font = self.title_font

        # Pad title with blank rows.
        self.append_empty_row(sheet, 2)

    def apply_styles(self, sheet, start_row=3):
        dims = {}

        for cell in sheet.rows[start_row]:
            cell.border = self.border
            cell.font = self.bold

        for i in range(3, len(sheet.rows)):
            padding = 2 if i == 0 else 1

            for cell in sheet.rows[i]:
                cell.border = self.border

                if cell.value:
                    dims[cell.column] = max((dims.get(cell.column, 0),
                        len(str(cell.value)) + padding, 4))

                    if isinstance(cell.value, datetime):
                        cell.number_format = self.date_format

        for col, value in dims.items():
            sheet.column_dimensions[col].width = value

    def convert_to_xlsx(self, file):
        # Read xls using xlrd.
        if isinstance(file, str):
            book = xlrd.open_workbook(file)
        else:
            book = xlrd.open_workbook(file_contents=file.stream.read())

        index = 0
        nrows, ncols = 0, 0

        while nrows * ncols == 0:
            sheet = book.sheet_by_index(index)
            nrows = sheet.nrows
            ncols = sheet.ncols
            index += 1

        # Prepare an xlsx workbook.
        workbook = Workbook()
        worksheet = workbook.active

        for row in range(0, nrows):
            for col in range(0, ncols):
                worksheet.cell(
                    row=row + 1, column=col + 1).value = sheet.cell_value(
                        row, col)

        return workbook

    def create_headers_list(self, source_cells):
        row_number = str(self.find_first_row() + 1)
        headers_generator = self.master_parts_sheet.iter_rows(
            'A' + row_number + ':Z' + row_number)
        headers_list = next(headers_generator)

        return headers_list

    def create_parts_list(self, initial_row=0):
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
            value = section[i][j].value

            if self.headers[j].value == 'QTY':
                value *= self.multiplier

            if j == len(headers):
                attr = ('PART GROUP', part_group)
            else:
                attr = (self.headers[j].value, value)

            return attr

        return [OrderedDict(item(i, j) for j in range(len(headers) + 1))
            for i in range(1, len(section))]

    def create_sheet(self, name):
        sheet = self.workbook.create_sheet()
        sheet.title = name

        return sheet

    def find_first_row(self, row=0):
        cell_value = str()

        while str(cell_value) != 'QTY':
            row += 1

            if row == self.master_parts_sheet.max_row:
                break

            cell_value = self.master_parts_sheet.cell(row=row, column=1).value

        return row - 1

    def find_last_row(self, row):
        cell_value = True

        while cell_value:
            row += 1

            if row == self.master_parts_sheet.max_row:
                break

            cell_value = self.master_parts_sheet.cell(row=row, column=1).value

        return row - 2

    def get_workbook(self):
        workbook = None

        if hasattr(self, 'workbook'):
            workbook = self.workbook

        return workbook

    def get_workbook_stream(self):
        workbook = None

        if hasattr(self, 'workbook'):
            wb_bytes = save_virtual_workbook(self.workbook)
            workbook = BytesIO(wb_bytes)

        return workbook

    def excelerate(self, file):
        # Load spreadsheet into Workbook object.
        if self.extension == '.xlsx':
            self.workbook = load_workbook(file)
        else:
            self.workbook = self.convert_to_xlsx(file)

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
        weld_picklist_data = copy.deepcopy(fabricated_list)
        self.add_column('RCVD', weld_picklist_data)

        self.append_title(weld_picklist_sheet)
        self.append_data(weld_picklist_data, weld_picklist_sheet)
        self.apply_styles(weld_picklist_sheet)

        self.append_signature('Picked By', weld_picklist_sheet)

        # Create WELD BOM sheet.
        weld_bom_sheet = self.create_sheet('WELD BOM')
        weld_bom_data = [x for x in copy.deepcopy(master_parts_list)
            if str(x['WELDED']) == 'WELDED'
            and str(x['WELDMENT USED']) != 'SHIPPED LOOSE']

        self.append_title(weld_bom_sheet)
        self.append_data(weld_bom_data, weld_bom_sheet)
        self.apply_styles(weld_bom_sheet)

        # Create WELD LOOSE sheet.
        weld_loose_sheet = self.create_sheet('WELD LOOSE')
        weld_loose_data = [x for x in copy.deepcopy(master_parts_list)
            if str(x['WELDED']) == 'WELDED'
            and str(x['WELDMENT USED']) == 'SHIPPED LOOSE']

        self.append_title(weld_loose_sheet)
        self.append_data(weld_loose_data, weld_loose_sheet)
        self.apply_styles(weld_loose_sheet)

        # Create WELD Packing Slip sheet.
        weld_packing_sheet = self.create_sheet('WELD Packing Slip')
        weld_packing_data = [x for x in copy.deepcopy(master_parts_list)
            if str(x['WELDMENT USED']) == 'SHIPPED LOOSE']
        self.add_column('PICKED', weld_packing_data)

        self.append_title(weld_packing_sheet)
        self.append_data(weld_packing_data, weld_packing_sheet)
        self.apply_styles(weld_packing_sheet)

        self.append_signature('Picked By', weld_packing_sheet)

        # Create FINISH Pick List sheet.
        finish_picklist_sheet = self.create_sheet('FINISH Pick List')
        finish_picklist_data = [x for x in copy.deepcopy(master_parts_list)
            if str(x['WELDMENT USED']) == 'SHIPPED LOOSE']
        self.add_column('RCVD', finish_picklist_data)

        self.append_title(finish_picklist_sheet)
        self.append_data(finish_picklist_data, finish_picklist_sheet)
        self.apply_styles(finish_picklist_sheet)

        self.append_signature('Picked By', finish_picklist_sheet)
        self.append_signature('Shipped By', finish_picklist_sheet)

        return self.workbook
