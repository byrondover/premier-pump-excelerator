#!/usr/bin/env python3

from excelerator import Excelerator

excelerator = Excelerator('test.xlsx')
workbook = excelerator.get_workbook()

workbook.save('test_complete.xlsx')
