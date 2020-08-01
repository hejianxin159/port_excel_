# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 16:12

from models import db_session
import sys
# from utlis.ExcelResolver import ExcelResolver
from .ExcelResolver import ExcelResolver

def read_position(file_path: str) -> list:
    excel_path = ExcelResolver()
    excel_path.set_file_path = file_path
    sheet_num = excel_path.get_sheet_num()
    device_position_info = []
    data_dict = {
            "device_name": "L",
            "room": "G",
            "cabinet_letter": "H",
            "cabinet_num": "I",
            "u": "J"
    }
    # data_dict = {
    #     "device_name": "A",
    #     "room": "B",
    #     "cabinet": "C",
    #     "u": "D"
    # }
    #找到所有sheet的数据
    for sheet_index in range(sheet_num):
        excel_path.getWorkSheet(sheet_index)
        data_info_list = excel_path.convert_excel_data_to_dict(2, data_dict)
        for data_dict_info in data_info_list:
            cabinet_letter = data_dict_info['cabinet_letter']
            cabinet_num = data_dict_info['cabinet_num']
            if isinstance(cabinet_num, float):
                cabinet_num = str(int(cabinet_num))
            data_dict_info['cabinet'] = cabinet_letter + cabinet_num
            del data_dict_info['cabinet_letter']
            del data_dict_info['cabinet_num']


        device_position_info.extend(data_info_list)
    return device_position_info

# read_position('../zhenzhou/device_position.xlsx')