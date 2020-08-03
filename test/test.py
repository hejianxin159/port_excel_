# # -*- coding: utf-8 -*-
# # author: hejianxin
# # date: 2020/8/2 1:28
#
# from terminaltables import AsciiTable
# import os
# import xlrd
# from openpyxl.utils import column_index_from_string
#
#
#
#
#
#
#
# class ExcelResolver:
#     """docstring for InputResolver."""
#     is_change = None
#
#     def __new__(cls, *args, **kwargs):
#         if cls.is_change == None:
#             cls.is_change = super(ExcelResolver, cls).__new__(cls, *args, **kwargs)
#
#         return cls.is_change
#
#
#     def __init__(self):
#         self.inputFile = None
#
#
#     @property
#     def set_file_path(self):
#         return getattr(self, 'inputFile', None)
#
#     @set_file_path.setter
#     def set_file_path(self, file_path):
#         if os.path.exists(file_path):
#             self.inputFile = file_path
#             self.workbook = xlrd.open_workbook(self.inputFile)
#
#
#     def getWorkSheet(self, sheet_index=None):
#         self.sheet_index = int(sheet_index) if sheet_index != None else 1
#         if self.inputFile == None:
#             print('Excel data file is not exists')
#             return False
#         else:
#             if not getattr(self, 'workbook', None):
#                 raise AttributeError('not find excel')
#             self.worksheet = self.workbook.sheet_by_index(self.sheet_index)
#             return self.worksheet
#
#     def convert_excel_data_to_dict(self, start_row_num, column_map):
#         if not isinstance(column_map, dict):
#             return False
#         self.xlsx_parse_dicts = []
#         for row_num in range(start_row_num, self.worksheet.nrows):
#             single_row_dict = {}
#             for key, column in column_map.items():
#                 column_num = column_index_from_string(column) - 1
#                 cell_value = self.worksheet.row_values(row_num)[column_num]
#                 single_row_dict[key] = cell_value
#             self.xlsx_parse_dicts.append(single_row_dict)
#
#         return self.xlsx_parse_dicts
#
#
#     def get_sheet_num(self):
#         if not getattr(self, 'inputFile', None) or not getattr(self, 'workbook', None):
#             raise AttributeError('not find excel')
#         return len(self.workbook.sheets())
#
#
#
#
#
# def pretty_print(title, field_names, body_data):
#     if isinstance(field_names, list) and isinstance(body_data, list):
#         TABLE_DATA = []
#         TABLE_DATA.append(field_names)
#         if isinstance(body_data[0], list):
#             # 如果body_data是多个列表组合而成的，则要循环添加
#             for row in body_data:
#                 TABLE_DATA.append(row)
#         else:
#             TABLE_DATA.append(body_data)
#         print_table = AsciiTable(TABLE_DATA, title)
#         return print_table.table
#     else:
#         print('Print table input arguments is not list')
#         return False
#
# def change_data(data_info_list: list, position_dict: dict)->list:
#     pretty_data = []
#     for data_info in data_info_list:
#         local_interface = data_info['Local_interface']                          #物理接口
#         desc_remote_dev = data_info['desc_remote_dev']                          #应该下联的设备名
#
#         desc_remote_interface = data_info['desc_remote_interface']              #应该下联设备接口名称
#         lldp_remote_dev = data_info['lldp_remote_dev']                          #现在实际下联的设备名
#         lldp_remote_port = data_info['lldp_remote_port']                        #现在实际下联设备的接口名称
#         remote_position = position_dict.get(desc_remote_dev, {})
#         lldp_position = position_dict.get(lldp_remote_dev, {})
#         desc_remote_position = remote_position.get('position', '')                                               #应该下联设备位置
#
#         lldp_remote_position = lldp_position.get('position', '')                                               #现在实际下联设备位置
#
#         status = 'ok'                                                             #最终状态
#         if desc_remote_dev != lldp_remote_dev:
#             status = 'error'
#         else:
#             if desc_remote_interface != lldp_remote_port:
#                 status = 'error'
#
#
#         pretty_data.append({'local_interface' :local_interface,
#                             'desc_remote_dev': desc_remote_dev,
#                             'desc_remote_position': desc_remote_position,
#                             'desc_remote_interface': desc_remote_interface,
#                             'lldp_remote_dev': lldp_remote_dev,
#                             'lldp_remote_position': lldp_remote_position,
#                             'lldp_remote_port': lldp_remote_port,
#                             'status': status,
#                             'interface_a': data_info['interface_a'],
#                             'interface_b': data_info['interface_b'],
#                             'interface_c': data_info['interface_c']
#                             })
#     return pretty_data
#     #     print(data_info)
#
#
# def split_interface(data_info_list):
#     for data_info in data_info_list:
#         interface = data_info['Local_interface']
#         interface = interface.replace('gi', '')
#         interface_split = interface.split('/')
#         data_info['interface_a'] = int(interface_split[0])
#         data_info['interface_b'] = int(interface_split[1])
#         data_info['interface_c'] = int(interface_split[2])
#
#     return data_info_list
#
#
# if __name__ == '__main__':
#     #读取excel数据
#     excel_parser = ExcelResolver()
#     excel_parser.set_file_path = './devices_location.xlsx'
#     excel_parser.getWorkSheet(0)   #sheet index
#     position_data_list = excel_parser.convert_excel_data_to_dict(1, {'device_name': 'A', 'position': 'B'})
#     position_dict = {i['device_name']: i for i in position_data_list}
#
#
#
#     # 读取json文件
#     with open('./res.json') as f:
#         data_info_list = eval(f.read())
#     data_info_list = split_interface(data_info_list)
#
#     #生成数据 [[], [], []....]
#     pretty_data = change_data(data_info_list, position_dict)
#
#
#     new_s_2 = sorted(pretty_data,
#                      key=lambda e: (e['status'], e.__getitem__('interface_a'), e.__getitem__('interface_b'), e.__getitem__('interface_c')))
#     pretty_data = new_s_2
#     pretty_data_list = []
#     for pretty_data_info in pretty_data:
#         pretty_data_list.append([pretty_data_info['local_interface'],
#                             pretty_data_info['desc_remote_dev'],
#                             pretty_data_info['desc_remote_position'],
#                             pretty_data_info['desc_remote_interface'],
#                             pretty_data_info['lldp_remote_dev'],
#                             pretty_data_info['lldp_remote_position'],
#                             pretty_data_info['lldp_remote_port'],
#                             pretty_data_info['status'],
#                             ])
#
#     res = pretty_print('-', ['物理接口', '应该下联的设备名', '应该下联设备位置', '应该下联设备接口名称', '现在实际下联的设备名', '现在实际下联设备位置',
#                              '现在实际下联设备的接口名称', '最终状态'], pretty_data_list)
#     print(res)
#     # for index, data_status in enumerate(pretty_data):
#     #     status = data_status[-1]
#     #     list_index = None
#     #     # now_status = 'error' if index == 0 and status == 'error' else None
#     #     if status == 'error':
#     #         list_index = index
#
#
#
#
#
#
#
#
#     # print(res)
#
# -*- coding: utf-8 -*-

import xlwt

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('123')  # 生成sheet
worksheet.col(0).width = 256*20

worksheet.row(2).height_mismatch = 1
worksheet.row(2).height = 1000
workbook.save('test.xlsx')
