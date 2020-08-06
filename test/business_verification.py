# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 15:27
from ExcelResolver import ExcelResolver
import os
import xlwt





# tall_style = xlwt.easyxf('font:name SimSun,height 720;align: wrap on, vert centre, horiz center;')
style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
al = xlwt.Alignment()
al.horz = 0x02  # 设置水平居中
al.vert = 0x01  # 设置垂直居中

borders = xlwt.Borders()  # Create borders
borders.left = 1  # 添加边框-虚线边框
borders.right = 1  # 添加边框-虚线边框
borders.top = 1  # 添加边框-虚线边框
borders.bottom = 1  # 添加边框-虚线边框

#橘色
orange_pattern = xlwt.Pattern()
orange_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
orange_pattern.pattern_fore_colour = 52
orange_style = xlwt.XFStyle()
orange_style.pattern = orange_pattern
orange_style.alignment = al
orange_style.borders = borders
#黄色
yellow_pattern = xlwt.Pattern()
yellow_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
yellow_pattern.pattern_fore_colour = 34
yellow_style = xlwt.XFStyle()
yellow_style.pattern = yellow_pattern
yellow_style.alignment = al
yellow_style.borders = borders
#蓝色
blue_pattern = xlwt.Pattern()
blue_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
blue_pattern.pattern_fore_colour = 49
blue_style = xlwt.XFStyle()
blue_style.pattern = blue_pattern
blue_style.alignment = al
blue_style.borders = borders

#浅蓝
light_blue_pattern = xlwt.Pattern()
light_blue_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
light_blue_pattern.pattern_fore_colour = 41
light_blue_style = xlwt.XFStyle()
light_blue_style.pattern = light_blue_pattern
light_blue_style.alignment = al
light_blue_style.borders = borders

#绿色
green_pattern = xlwt.Pattern()
green_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
green_pattern.pattern_fore_colour = 17
green_style = xlwt.XFStyle()
green_style.pattern = green_pattern
green_style.alignment = al
green_style.borders = borders

excel_parser = ExcelResolver()
excel_parser.set_file_path = './read_file.xls'
excel_parser.getWorkSheet(0)
file_list = excel_parser.convert_excel_data_to_dict(1, {"date_time": "B", "sheet_name": "I", "file_name": "S"})
file_dict = {}
for file in file_list:
    is_exist_date_time = file_dict.get(file["date_time"])
    if is_exist_date_time:
        # print(is_exist_date_time)
        is_exist_file = is_exist_date_time.get(file['file_name'])
        if is_exist_file:
            # print(is_exist_file)
            is_exist_file.append(file['sheet_name'])

            is_exist_date_time[file['file_name']] = is_exist_file
            # file_dict[file['date_time']] = is_exist_date_time
        else:
            is_exist_date_time[file['file_name']] = [file['sheet_name']]
    else:
        file_dict[file['date_time']] = {file['file_name']: [file['sheet_name']]}

# print(file_dict)
# exit()

for key_i, val_i in file_dict.items():
    for key, val in val_i.items():
        val_i[key] = list(set(val))
    file_dict[key_i] = val_i
    # val[list(val.keys())[0]] = list(set(list(val.values())[0]))
    # file_dict[key] = val

# print(file_dict)
# #
# for key, val in file_dict.items():
#     print(key, val )
# workbook.save('2020业务核查.xlsx')

# file_dict = {"20203276广东(_PS-IMS_NGN_SGi)调整单": ['ChinaMobile_IUPS_Media']}

for date_time, data_dict in file_dict.items():

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')

    head_name_list = ['调度单编号', '设备名称', '详细地址', '设备类型', '互联端口', '', 'PE侧接口地址', '端口状态', 'VPN名称',
                      '业务类型', 'OSPF/VPRN进程/ebgp', 'OSPF邻居状态', 'PE对外发布的VPN路由', '路由状态', 'RD值', 'RR路由查看结果',
                      '亚信联系人', '核查人', '备注', 'AR局地、编号']
    tall_style = xlwt.easyxf('font:height 720;')
    worksheet.row(0).set_style(tall_style)

    for index, head_name in enumerate(head_name_list):
        worksheet.col(index).width = 255 * 12
        if index in [0, 9, 16]:
            worksheet.write(0, index, head_name, orange_style)
        elif index in [1, 2, 3, 4, 5, 6, 8, 12, 14]:
            worksheet.write(0, index, head_name, yellow_style)
        elif index in [7, 11, 13, 15]:
            worksheet.write(0, index, head_name, blue_style)
        elif index in [12, 10, 17]:
            worksheet.write(0, index, head_name, green_style)
        elif index == 19:
            worksheet.write(0, index, head_name, light_blue_style)

    file_index = 1
    for file_name, file_sheet_list in data_dict.items():
        excel_parser.set_file_path = './Desktop/' + file_name + '.xls'
        # print(file_path)
        #生成文件名
        # print('./Desktop/' + file_name + '.xls')
        worksheet.write(file_index, 0, file_name, light_blue_style)
        worksheet.row(file_index).set_style(tall_style)

        for sheet_name in file_sheet_list:
            sheet_index = excel_parser.get_sheet_name(sheet_name)
            excel_parser.getWorkSheet(sheet_index)
            res_e = excel_parser.convert_excel_data_to_dict(0, {"is_device_name": "E"})
            data_info_list = []
            sheet_type = None  # 1 为调成前后  2为ce设备名称
            for get_device_name in res_e:
                if get_device_name['is_device_name'] == "调整前后":
                    #判断E列是调整前后还是CE设备名称
                    get_col_dict = {
                        "after": "E",
                        "device_name": "R",
                        "place": "S",
                        "device_type": "T",
                        "z_port": "U",
                        "port_type": "V",
                        "PE_ip": "W",
                        "VPN": "P",
                        "RD": "R",
                        "RD_val": "S"
                    }
                    data_info_list = excel_parser.convert_excel_data_to_dict(0, get_col_dict)
                    sheet_type = 1
                    break
                elif get_device_name['is_device_name'] == "CE设备名称":
                    get_col_dict = {
                        "after": "E",
                        "device_name": "Q",
                        "place": "R",
                        "device_type": "S",
                        "z_port": "T",
                        "port_type": "U",
                        "PE_ip": "V",
                        "VPN": "O",
                        "RD": "Q",
                        "RD_val": "R"

                    }
                    data_info_list = excel_parser.convert_excel_data_to_dict(0, get_col_dict)
                    sheet_type = 2
                    break

            if sheet_type ==1:
                rd_val = ''
                vpn = ''
                for index_after, data_info in enumerate(data_info_list):
                    if not isinstance(data_info["RD"], float):
                        if data_info["RD"].strip() == "RD":
                            rd_val = data_info["RD_val"]

                    if data_info['after'] == "调整后":
                        worksheet.row(file_index).set_style(tall_style)
                        if len(data_info["VPN"]) == 0:
                            vpn = data_info_list[index_after -1]["VPN"]
                        worksheet.write(file_index, 1, data_info['device_name'], orange_style)
                        worksheet.write(file_index, 2, data_info['place'], orange_style)
                        worksheet.write(file_index, 3, data_info['device_type'], orange_style)
                        worksheet.write(file_index, 4, data_info['z_port'], orange_style)
                        worksheet.write(file_index, 5, data_info['port_type'], orange_style)
                        worksheet.write(file_index, 6, data_info['PE_ip'], orange_style)
                        worksheet.write(file_index, 8, sheet_name, light_blue_style)
                        worksheet.write(file_index, 12, vpn, yellow_style)
                        worksheet.write(file_index, 14, rd_val, light_blue_style)
                        file_index += 1

            elif sheet_type == 2:
                rd_val = ''
                vpn = ''
                is_write = False
                for data_info in data_info_list:
                    if not isinstance(data_info["RD"], float):
                        if data_info["RD"].strip() == "RD":
                            rd_val = data_info["RD_val"]
                    if data_info["after"] == "CE设备名称":
                        is_write = True
                        continue
                    if is_write and len(data_info["after"]) != 0:
                        if len(data_info['VPN']) > 0:
                            vpn = data_info['VPN']
                        worksheet.row(file_index).set_style(tall_style)

                        worksheet.write(file_index, 1, data_info['device_name'], orange_style)
                        worksheet.write(file_index, 2, data_info['place'], orange_style)
                        worksheet.write(file_index, 3, data_info['device_type'], orange_style)
                        worksheet.write(file_index, 4, data_info['z_port'], orange_style)
                        worksheet.write(file_index, 5, data_info['port_type'], orange_style)
                        worksheet.write(file_index, 6, data_info['PE_ip'], orange_style)
                        worksheet.write(file_index, 8, sheet_name, light_blue_style)
                        worksheet.write(file_index, 12, vpn, yellow_style)
                        worksheet.write(file_index, 14, rd_val, light_blue_style)
                        file_index += 1
        # file_index += 1
            # exit()
    print(date_time)
    workbook.save('./file/' + date_time + '业务核查.xls')
    # exit()

    # exit()