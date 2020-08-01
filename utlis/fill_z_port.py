# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/27 11:22


from utlis.ExcelResolver import ExcelResolver
from utlis.read_position import read_position
import xlwt
from models import db_session, Devices

#
class FullZPort():

    def __init__(self, position_dict: dict):
        self.position_dict = position_dict

    def full_z_port(self, file_name_list: list):
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
        for file_name in file_name_list:
            # 文件名称
            excel_name = file_name.split('/')[-1].split('.')[0]
            excel_parser = ExcelResolver()
            excel_parser.set_file_path = file_name

            workbook = xlwt.Workbook(encoding='utf-8')

            # 获取sheet名称
            sheet_name = ''
            # for sheet_index in range(excel_parser.get_sheet_num()):
            for sheet_index in range(excel_parser.get_sheet_num()):
                # 通过sheet下标获取sheet
                excel_parser.getWorkSheet(sheet_index)
                sheet_info_list = excel_parser.convert_excel_data_to_dict(1, device_dict)
                for data in sheet_info_list:
                    if not sheet_name:
                        sheet_name = data['A_device']
                        self.worksheet = workbook.add_sheet(sheet_name)
                    if sheet_index >= 35:
                        a_device = data['A_device']
                        z_device = data['Z_device']
                        device_data = db_session.query(Devices).filter_by(a_device = z_device, z_device = a_device)
                        for i in device_data:
                            if i.is_use:
                                continue
                            else:
                                data['Z_port'] = i.a_port
                                i.is_use = True
                                db_session.commit()
                                break
                        # a_position = self.position_dict.get(data['A_device'])
                        z_position = self.position_dict.get(data['Z_device'])
                        # if a_position:
                        #     data['A_home'] = a_position['room']
                        #     data['A_cabinet'] = a_position['cabinet']
                        #     data['A_u_position'] = a_position['u']
                        if z_position:
                            data['Z_home'] = z_position['room']
                            data['Z_cabinet'] = z_position['cabinet']
                            data['Z_u_position'] = z_position['u']
                self.update_file(sheet_info_list)
                sheet_name = ''
            workbook.save('../file/' + excel_name + 'bak_2' + '.xlsx')

    def update_file(self, sheet_data: list) -> None:
        self.worksheet.write(0, 0, "A端设备所在机房")
        self.worksheet.write(0, 1, "A端设备所在机柜")
        self.worksheet.write(0, 2, "A端设备所在U位", )
        self.worksheet.write(0, 3, "A端设备", )
        self.worksheet.write(0, 4, "A端物理端口", )
        self.worksheet.write(0, 5, "Z端设备所在机房")
        self.worksheet.write(0, 6, "Z端设备所在机柜", )
        self.worksheet.write(0, 7, "Z端设备所在U位", )
        self.worksheet.write(0, 8, "Z端设备", )
        self.worksheet.write(0, 9, "Z端物理端口", )
        for index, data in enumerate(sheet_data):
            index += 1
            self.worksheet.write(index, 0, data['A_home'])
            self.worksheet.write(index, 1, data['A_cabinet'])
            self.worksheet.write(index, 2, data['A_u_position'])
            self.worksheet.write(index, 3, data['A_device'])
            self.worksheet.write(index, 4, data['A_port'])
            self.worksheet.write(index, 5, data['Z_home'])
            self.worksheet.write(index, 6, data['Z_cabinet'])
            self.worksheet.write(index, 7, data['Z_u_position'])
            self.worksheet.write(index, 8, data['Z_device'])
            self.worksheet.write(index, 9, data['Z_port'])


if __name__ == '__main__':
    data = read_position('../shanxi/device_position.xlsx')
    data = {i['device_name']: i for i in data}
    db_session.query(Devices).update({Devices.is_use: 0})
    db_session.commit()
    full = FullZPort(data)
    full.full_z_port(['../file/admin2bak.xlsx'])
    # full_position(['../POD1report.xlsx'], data)
    # print(data)
    # print(len(data))



