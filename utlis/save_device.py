# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 17:14

# from utlis.ExcelResolver import ExcelResolver
from .ExcelResolver import ExcelResolver
from models import db_session, Devices
import copy
import re


z_device_port_dict = {}

def change_device_data(device_name: str, device_data: list) -> list:
    data = []
    # re_find_name_num = re.findall(r'.*?(\d+) [锐,华,贝].*', device_name)
    # if len(re_find_name_num) == 0:
    #     re_find_name_num = re.findall('(\d+).*', device_name.split('-')[0])  #POD2 1G管理接入交换机-7  (4个10GE光+48个GE电） 华为S5335
    # assert len(re_find_name_num) == 1
    # device_name_num = int(re_find_name_num[0])
    device_name_split_list = device_name.split('-')
    if len(device_name_split_list) > 1 and len(device_name_split_list) == 2 and '网络设备管' not in device_name:
        device_name_num = int(device_name_split_list[-1])
    elif 'IPMI接入交换机' in device_name or '九期' in device_name or '网络设备管' in device_name or '管理域管理汇聚交换机' in device_name:
        device_name_num = int(re.findall(r'.*?(\d+)', device_name)[0])
    else:
        print(device_name)
    for port_info in device_data:
        if port_info["is_exist"] == "原有":
            continue
        cabinet_num = port_info['cabinet_num']
        a_port = port_info["a_port"]
        if isinstance(a_port, float):
            a_port = str(int(a_port))

        # print(port_info)
        is_reserve = port_info["is_reserve"]
        z_device = port_info["z_device"].strip()
        port_type = port_info['port_type']
        port_info.update({'a_device': device_name})
        if is_reserve == "预留" or z_device == "预留":
            del port_info["is_reserve"]
            data.append(port_info)
        elif is_reserve == '' and z_device == '' and a_port == '':
            continue
        elif is_reserve == "空" or z_device == "空":
            del port_info["is_reserve"]
            port_info['z_device'] = "预留"
            data.append(port_info)
        else:
            del port_info["is_reserve"]
            a_port_list = a_port.split('-')              #   1-5   1/1/1-1/1/5
            if len(a_port_list) == 2:
                range_left_port_name = a_port_list[0].split('/')   # [1, 5], [1/1/1, 1/1/5]
                range_right_port_name = a_port_list[1].split('/')

                if len(range_left_port_name) > 1 and len(range_right_port_name) > 1:
                    #[1/1/1, 1/1/5]
                    a_port_first = int(range_left_port_name[-1])
                    a_port_last = int(range_right_port_name[-1])
                else:

                    a_port_first = int(range_left_port_name[0])
                    a_port_last = int(range_right_port_name[0])
                a_port_num = a_port_last - a_port_first + 1
                z_device_list = z_device.split('~')
                if len(z_device_list) > 1:
                    z_device_name_prefix = z_device_list[0].split('-')[:-1]
                    z_device_first = int(z_device_list[0].split('-')[-1])
                    z_device_last = int(z_device_list[1])
                    z_device_port_num = z_device_last - z_device_first + 1
                    z_device_num_list = [str(i) for i in list(range(z_device_first, z_device_last + 1))]
                    #
                    # is_exist_z_port = z_device_port_dict.get(z_device_list[0])
                    # if is_exist_z_port:
                    #     if is_exist_z_port == '1#业务口'
                    #


                    if device_name_num % 2 == 0:
                        if '管理' in device_name:
                            z_port_first = '2#管理口'
                            z_port_last = '1#管理口'
                        elif 'IPMI' in device_name:
                            z_port_first = 'IPMI口'
                            z_port_last = 'IPMI口'
                        else:
                            z_port_first = '3#业务口'
                            z_port_last = '4#业务口'
                    else:
                        if '管理' in device_name:
                            z_port_first = '1#管理口'
                            z_port_last = '2#管理口'
                        elif 'IPMI' in device_name:
                            z_port_first = 'IPMI口'
                            z_port_last = 'IPMI口'
                        else:
                            z_port_first = '1#业务口'
                            z_port_last = '2#业务口'
                    # z_port_first = ''
                    # z_port_last = ''
                    #
                    # if '管理' in device_name:
                    #     port_first = db_session.query(Devices.z_port).filter_by(z_device = z_device_list[0]).first()
                    #     if port_first:
                    #         if port_first[0] == '1#管理口':
                    #             z_port_first = '3#管理口'
                    #             z_port_last = '4#管理口'
                    #         else:
                    #             z_port_first = '1#管理口'
                    #             z_port_last = '2#管理口'
                    #     else:
                    #         z_port_first = '1#管理口'
                    #         z_port_last = '2#管理口'
                    # elif 'IPMI' in device_name:
                    #     z_port_first = 'IPMI口'
                    #     z_port_last = 'IPMI口'
                    # else:
                    #     port_first = db_session.query(Devices.z_port).filter_by(z_device=z_device_list[0]).first()
                    #     if z_device_list[0] == 'P10F2-POD3-ST-osd-S3-NF5466M5-224':
                    #         print(port_first)
                    #     if port_first:
                    #         if port_first[0] == '1#业务口':
                    #             z_port_first = '2#业务口'
                    #             z_port_last = '4#业务口'
                    #         else:
                    #             z_port_first = '1#业务口'
                    #             z_port_last = '3#业务口'
                    #     else:
                    #         z_port_first = '1#业务口'
                    #         z_port_last = '3#业务口'

                    if z_device_port_num * 2 == a_port_num:         #前面是后面2倍的情况
                        z_device_num_list = [i for i in z_device_num_list for j in range(2)]
                        # a_device_num_list = [str(i+1) for i in list(range(a_port_num))]
                        #123456  123456
                        for index, num in enumerate(range(a_port_first, a_port_last + 1)):
                            if index % 2 == 0:
                                z_port_name = z_port_first
                            else:
                                z_port_name = z_port_last
                            data.append({"a_device": device_name,
                                         "a_port": '/'.join(range_left_port_name[:-1]) + '/' + str(num) if len(range_left_port_name) > 1 and len(range_right_port_name) > 1 else str(num) ,
                                         "z_device": '-'.join(z_device_name_prefix) + '-' + z_device_num_list[index],
                                         "z_port": z_port_name,
                                         "cabinet_num": cabinet_num,
                                         "port_type": port_type})
                    elif z_device_port_num * 3 == a_port_num:
                        z_device_num_list = [i for i in z_device_num_list for j in range(3)]

                        for index, num in enumerate(range(a_port_first, a_port_last + 1)):
                            data.append({"a_device": device_name,
                                         "a_port": '/'.join(range_left_port_name[:-1]) + '/' + str(num) if len(range_left_port_name) > 1 and len(range_right_port_name) > 1 else str(num) ,
                                         "z_device": '-'.join(z_device_name_prefix) + '-' + z_device_num_list[index],
                                         "cabinet_num": cabinet_num,
                                         "port_type": port_type
                                         })

                    else:                                           #前面和后面相等的情况
                        if device_name_num % 2 == 0:
                            if 'POD' in device_name and '业务' in device_name:
                                z_port_first = '3#业务口'
                            elif 'POD' in device_name and '存储' in device_name:
                                z_port_first = '4#业务口'
                        else:
                            if 'POD' in device_name and '业务' in device_name:
                                z_port_first = '1#业务口'
                            elif 'POD' in device_name and '存储' in device_name:
                                z_port_first = '2#业务口'
                        for index, num in enumerate(range(a_port_first, a_port_last + 1)):

                            data.append({"a_device": device_name,
                                         "a_port": '/'.join(range_left_port_name[:-1]) + '/' + str(num) if len(range_left_port_name) > 1 and len(range_right_port_name) > 1 else str(num) ,
                                         "z_device": '-'.join(z_device_name_prefix) + '-' + z_device_num_list[index],
                                         "z_port": z_port_first,
                                         "cabinet_num": cabinet_num,
                                         "port_type": port_type})

                else:                                                #前面是两个，后面是一个的情况
                    for index, num in enumerate(range(a_port_first, a_port_last + 1)):
                        data.append({"a_device": device_name,
                                     "a_port":'/'.join(range_left_port_name[:-1]) + '/' + str(num) if len(range_left_port_name) > 1 and len(range_right_port_name) > 1 else str(num) ,
                                     "z_device": z_device_list[-1],
                                     "cabinet_num": cabinet_num,
                                     "port_type": port_type})
            else:
                if "&" in z_device: #{'a_device': 'POD1-业务核心交换机-锐捷N18010-1', 'a_port': '', 'z_device': '网络设备管理接入交换机-华为S5335-5&6'}
                    z_device_split_list = z_device.split('&')
                    z_device_split_first = z_device_split_list[0]
                    if len(a_port) == 0 and '网络设备管理接入交换机' in port_info['z_device']:
                        port_info['a_port'] = '管理口'
                    # z_device_prefix = '-'.join(z_device_split_first.split('-'))
                    z_device_prefix = z_device_split_first
                    port_info = copy.deepcopy(port_info)

                    port_info['cabinet_num'] = cabinet_num
                    port_info['z_device'] = z_device_split_first
                    data.append(port_info)
                    port_info_copy = copy.deepcopy(port_info)
                    port_info_copy['z_device'] = z_device_prefix[:-1] + z_device_split_list[-1]

                    data.append(port_info_copy)

                else:
                    if len(a_port) == 0 and '网络设备管理接入交换机' in port_info['z_device']:
                        port_info['a_port'] = '管理口'


                        data.append(port_info)
                    else:
                        data.append(port_info)
    return data





def save_device(file_path_list: list) -> list:
    excel_parser = ExcelResolver()
    message = []
    for file_path in file_path_list:    #遍历excel文件
        excel_name = file_path.split('/')[-1]  #找出当前这个excel的名字
        excel_parser.set_file_path = file_path
        sheet_num = excel_parser.get_sheet_num()
        data_dict = {
            "a_device": "B",           #判断是否是设备
            "a_port": "D",            #板卡编号
            "is_reserve": "E",       #判断是否预留
            "z_device": "F",           #z端设备
            "is_exist": "I",
            "port_type": "E"
        }
        # for sheet_index in range(sheet_num):
        for sheet_index in range(sheet_num):  #遍历sheet
            excel_parser.getWorkSheet(sheet_index)
            data_list = excel_parser.convert_excel_data_to_dict(0, data_dict) #164
            device_name = ''        #当前设备名字
            device_info = []
            for index, data_info in enumerate(data_list):
                a_device = data_info['a_device']
                # if a_device != "板卡编号" and len(a_device) > 0 and device_name != '':
                #     a_device_num = str(int(a_device))
                if isinstance(a_device, float):
                    cabinet_num = str(int(a_device))
                if a_device == "板卡编号":
                    if len(device_info) > 0:
                        message.append({excel_name: change_device_data(device_name, device_info[:-1])}) #取到-1是因为-1是下一台设备的名字
                        device_info = []

                    device_name = data_list[index-1]['a_device']
                # if '实配' in data_info['a_port']:
                #     continue
                if device_name != '' and a_device != "板卡编号":
                    #找到一台设备
                    data_info['cabinet_num'] = cabinet_num
                    device_info.append(data_info)
            if len(device_info) != 0:
                message.append({excel_name: change_device_data(device_name, device_info)})


    return message




# data =  save_device(['../POD1.xlsx', '../POD2.xlsx'])
if __name__ == '__main__':

    data =  save_device(['../zhenzhou/POD.xlsx'])








