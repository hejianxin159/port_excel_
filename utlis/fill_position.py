# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/27 11:22


from utlis.ExcelResolver import ExcelResolver
import xlwt

def read_position(file_path: str) -> list:
    excel_path = ExcelResolver()
    excel_path.set_file_path = file_path
    sheet_num = excel_path.get_sheet_num()
    device_position_info = []
    data_dict = {
            "en_name": "L",
            "room": "G",
            "cabinet_letter": "H",
            "cabinet_num": "I",
            "u": "J",
            "cn_name": "M"
    }

    #找到所有sheet的数据
    # for sheet_index in range(sheet_num):
    excel_path.getWorkSheet(4)
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


class FullPosition():

    def __init__(self, position_dict: dict):
        self.position_dict = position_dict

    def full_position(self, file_name_list: list):
        device_dict = {"A_home": 'A',
                       "A_cabinet": 'B',
                       "A_u_position": 'C',
                       "A_device": 'D',
                       "A_port": 'E',
                       "port_type": "F",
                       "Z_home": 'G',
                       "Z_cabinet": 'H',
                       "Z_u_position": 'I',
                       "Z_device": 'J',
                       "Z_port": 'K',
                       }
        for file_name in file_name_list:
            #文件名称
            excel_name = file_name.split('/')[-1].split('.')[0]
            excel_parser = ExcelResolver()
            excel_parser.set_file_path = file_name

            workbook = xlwt.Workbook(encoding='utf-8')

            #获取sheet名称
            sheet_name = ''
            for sheet_index in range(excel_parser.get_sheet_num()):
                #通过sheet下标获取sheet
                excel_parser.getWorkSheet(sheet_index)
                sheet_info_list = excel_parser.convert_excel_data_to_dict(1, device_dict)

                for data in sheet_info_list:
                    if not sheet_name:
                        sheet_name = data['A_device']
                        self.worksheet = workbook.add_sheet(sheet_name)
                    a_position = self.position_dict.get(data['A_device'].strip())
                    z_position = self.position_dict.get(data['Z_device'].strip())
                    if a_position:
                        data['A_device'] = a_position['en_name']
                        data['A_home'] = a_position['room']
                        data['A_cabinet'] = a_position['cabinet']
                        data['A_u_position'] = a_position['u']
                    if z_position and len(data['Z_home']) == 0 and len(data['Z_cabinet']) == 0:
                        data['Z_device'] = z_position.get('en_name') if z_position.get('en_name') else data['Z_device']
                        data['Z_home'] = z_position['room']
                        data['Z_cabinet'] = z_position['cabinet']
                        data['Z_u_position'] = z_position['u']
                    #
                    
                self.update_file(sheet_info_list)
                sheet_name = ''
            workbook.save('../file/' + excel_name + 'bak' + '.xlsx')

    def update_file(self, sheet_data: list) -> None:
        self.worksheet.write(0, 0, "A端设备所在机房")
        self.worksheet.write(0, 1, "A端设备所在机柜")
        self.worksheet.write(0, 2, "A端设备所在U位", )
        self.worksheet.write(0, 3, "A端设备")
        self.worksheet.write(0, 4, "A端物理端口")
        self.worksheet.write(0, 5, "A端端口类型")
        self.worksheet.write(0, 6, "Z端设备所在机房")
        self.worksheet.write(0, 7, "Z端设备所在机柜", )
        self.worksheet.write(0, 8, "Z端设备所在U位", )
        self.worksheet.write(0, 9, "Z端设备", )
        self.worksheet.write(0, 10, "Z端物理端口", )
        for index, data in enumerate(sheet_data):
            index += 1
            self.worksheet.col(index).width = 256 * 25
            self.worksheet.write(index, 0, data['A_home'])
            self.worksheet.write(index, 1, data['A_cabinet'])
            self.worksheet.write(index, 2, data['A_u_position'])
            self.worksheet.write(index, 3, data['A_device'])
            self.worksheet.write(index, 4, data['A_port'])
            self.worksheet.write(index, 5, data['port_type'])
            self.worksheet.write(index, 6, data['Z_home'])
            self.worksheet.write(index, 7, data['Z_cabinet'])
            self.worksheet.write(index, 8, data['Z_u_position'])
            self.worksheet.write(index, 9, data['Z_device'])
            self.worksheet.write(index, 10, data['Z_port'])











if __name__ == '__main__':

    data = read_position('../zhenzhou/device_position_s.xlsx')

    data = {i['cn_name']: i for i in data}

    full = FullPosition(data)
    full.full_position(['../PODbakreport.xlsx'])
    # full_position(['../POD1report.xlsx'], data)
    # print(data)
    # print(len(data))



