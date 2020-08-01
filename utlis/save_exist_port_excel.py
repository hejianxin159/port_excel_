# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/31 13:23
from models import db_session, Devices
from utlis.ExcelResolver import ExcelResolver

device_dict = {"A_home": 'A',
               "A_cabinet": 'B',
               "A_u_position": 'C',
               "A_device": 'D',
               "A_port": 'E',
               "Z_home": 'F',
               "Z_cabinet": 'G',
               "Z_u_position": 'H',
               "Z_device": 'I',
               "Z_port": 'J',
               }
for file_name in ['../shanxi/complete_pod1.xlsx', '../shanxi/complete_pod2.xlsx']:
    excel_parser = ExcelResolver()
    excel_parser.set_file_path = file_name
    index_sheet_list = excel_parser.get_sheet_num()
    for i in range(index_sheet_list):
        excel_parser.getWorkSheet(i)
        data_list = excel_parser.convert_excel_data_to_dict(1, device_dict)
        for data in data_list:
            exist_data = db_session.query(Devices).filter_by(a_device = data['A_device'], a_port = data['A_port'], z_device = data['Z_device']).first()
            if exist_data:
                print(data)
            else:
                db_session.add(Devices(a_device = data['A_device'], a_port = data['A_port'], z_device = data['Z_device']))

        db_session.commit()
