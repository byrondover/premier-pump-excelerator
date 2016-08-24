#!/usr/bin/env python3

import cProfile

from excelerator import Excelerator

cProfile.run('excelerator = Excelerator("test.xls")', sort='cumtime')

workbook = excelerator.get_workbook()

workbook.save('test_complete.xlsx')
