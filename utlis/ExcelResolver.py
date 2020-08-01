# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 15:27
import os
import xlrd
from openpyxl.utils import column_index_from_string


class ExcelResolver:
    """docstring for InputResolver."""
    is_change = None

    def __new__(cls, *args, **kwargs):
        if cls.is_change == None:
            cls.is_change = super(ExcelResolver, cls).__new__(cls, *args, **kwargs)

        return cls.is_change


    def __init__(self):
        self.inputFile = None


    @property
    def set_file_path(self):
        return getattr(self, 'inputFile', None)

    @set_file_path.setter
    def set_file_path(self, file_path):
        if os.path.exists(file_path):
            self.inputFile = file_path
            self.workbook = xlrd.open_workbook(self.inputFile)


    def getWorkSheet(self, sheet_index=None):
        self.sheet_index = int(sheet_index) if sheet_index != None else 1
        if self.inputFile == None:
            print('Excel data file is not exists')
            return False
        else:
            if not getattr(self, 'workbook', None):
                raise AttributeError('not find excel')
            self.worksheet = self.workbook.sheet_by_index(self.sheet_index)
            return self.worksheet

    def convert_excel_data_to_dict(self, start_row_num, column_map):
        if not isinstance(column_map, dict):
            return False
        self.xlsx_parse_dicts = []
        for row_num in range(start_row_num, self.worksheet.nrows):
            single_row_dict = {}
            for key, column in column_map.items():
                column_num = column_index_from_string(column) - 1
                cell_value = self.worksheet.row_values(row_num)[column_num]
                single_row_dict[key] = cell_value
            self.xlsx_parse_dicts.append(single_row_dict)

        return self.xlsx_parse_dicts


    def get_sheet_num(self):
        if not getattr(self, 'inputFile', None) or not getattr(self, 'workbook', None):
            raise AttributeError('not find excel')
        return len(self.workbook.sheets())


    # def get_sheet_name(self, index):
    #     if not getattr(self, 'inputFile', None) or not getattr(self, 'workbook', None):
    #         raise AttributeError('not find excel')
    #     return self.workbook.sheets()[index]