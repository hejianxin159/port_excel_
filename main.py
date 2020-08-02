# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 15:26

from utlis.read_position import read_position
from models import db_session, Devices, DevicePosition, engine, Base
from utlis.save_device import save_device
import asyncio
import xlwt
import re


def splicing_name(cabinet: str, port_name: str)-> str:
    #  1  1/1/1   1-5  1/1/1-1/1/5
    if port_name in ('1#业务口', '2#业务口', '3#业务口', '4#业务口'):
        return port_name
    if port_name == None:
        return ''
    if len(port_name) == 0:
        return ''
    if len(re.findall('[\u4e00-\u9fa5]', port_name)) > 0:
        return port_name
    port_name_split = port_name.split('/')
    splicing_name_res = ''
    if len(port_name_split) >= 3:
        splicing_name_res = port_name
    elif len(port_name_split) == 2:
        return cabinet + '/' + port_name
    elif len(port_name_split) == 1:
        port_name_is_two = port_name_split[0].split('-')
        if len(port_name_is_two) == 1:
            splicing_name_res = cabinet + '/0/' + port_name_is_two[0]
        elif len(port_name_is_two) == 2:
            splicing_name_res = cabinet + '/0/' + port_name_is_two[0] + '~' + cabinet + '/0/' + port_name_is_two[1]

    else:
        return port_name
    return splicing_name_res



async def save_device_data(device_data: dict):
    for key, value in device_data.items():  #{"POD1.xlsx" :{'POD1-业务核心交换机-锐捷N18010-1': [{'a_device': 'POD1-业务核心交换机-锐捷N18010-1', 'a_port': '1/1'}]}
        for data_info in value:
            #查找设备位置
            a_device_position_id = db_session.query(DevicePosition).filter_by(device_name = data_info['a_device']).first()
            z_device_position_id = db_session.query(DevicePosition).filter_by(device_name = data_info['z_device']).first()
            # device_position_id = db_session.query(DevicePosition).filter_by(device_name = 'as').first()
            if a_device_position_id:
                data_info['a_cabinet_id'] = a_device_position_id.id
            if z_device_position_id:
                data_info['z_cabinet_id'] = z_device_position_id.id
            data_info['sheet_name'] = key
    for i in value:
        if 'is_exist' in i:
            del i['is_exist']

        # print(i)
        db_session.add(Devices(**i))
    db_session.commit()
    # exit()
    # db_session.execute(Devices.__table__.insert(), value)


async def change_port_excel(worksheet: object, device_name: str):
    device_detail_data_list = db_session.query(Devices).filter_by(a_device = device_name)
    for index, device_detail_data in enumerate(device_detail_data_list):
        a_device_name = device_detail_data.a_device
        a_port = device_detail_data.a_port
        a_room = ''         #机房
        a_cabinet_name = '' #机柜
        a_u = ''              #u位
        z_port = device_detail_data.z_port
        z_device_name = device_detail_data.z_device
        z_room = ''         #机房
        z_cabinet_name = '' #机柜
        z_u = ''              #u位
        a_cabinet_num = device_detail_data.cabinet_num
        z_cabinet_num = ''
        a_cabinet = device_detail_data.a_device_position_id
        if a_cabinet:
            a_room = a_cabinet.room
            a_u = a_cabinet.u
            a_cabinet_name = a_cabinet.cabinet
        z_cabinet = device_detail_data.z_device_position_id
        if z_cabinet:
            z_room = z_cabinet.room
            z_u = z_cabinet.u
            z_cabinet_name = z_cabinet.cabinet

        z_port_name = device_detail_data.z_port

        if not z_port_name:
            is_z_port_list = db_session.query(Devices).filter_by(a_device = z_device_name, z_device = a_device_name)
            for is_z_port in is_z_port_list:
                if is_z_port.is_use:
                    continue
                else:
                    z_port = is_z_port.a_port
                    is_z_port.is_use = True
                    z_cabinet_num = is_z_port.cabinet_num
                    # db_session.commit()
                    db_session.add(is_z_port)
                    break

        a_port_s = splicing_name(a_cabinet_num, a_port)
        z_port_s = splicing_name(z_cabinet_num, z_port)

        # print('12312', a_port_s, z_port_s)
        index += 1
        worksheet.write(index, 0, a_room)
        worksheet.write(index, 1, a_cabinet_name)
        worksheet.write(index, 2, a_u)
        worksheet.write(index, 3, a_device_name)
        worksheet.write(index, 4, a_port_s)
        worksheet.write(index, 5, device_detail_data.port_type)
        worksheet.write(index, 6, z_room)
        worksheet.write(index, 7, z_cabinet_name)
        worksheet.write(index, 8, z_u)
        worksheet.write(index, 9, z_device_name)
        worksheet.write(index, 10, z_port_s)
    db_session.commit()

    # workbook.save('report.xlsx')
    # exit()


async def main():
    data =  save_device(['./zhenzhou/PODbak.xlsx', './zhenzhou/nine.xlsx', './zhenzhou/huiju.xlsx', './zhenzhou/admin.xlsx'])
    # # data =  save_device(['./zhenzhou/PODbak.xlsx'])
    # # data =  save_device(['./zhenzhou/nine.xlsx'])
    # # data =  save_device(['./zhenzhou/admin.xlsx'])
    for data_info in data:
        # 数据入库
        future = asyncio.ensure_future(save_device_data(data_info))
        await asyncio.sleep(0)

    #生成报表
    file_name_list = db_session.query(Devices.sheet_name).group_by(Devices.sheet_name)
    for file_name in file_name_list:
        if file_name[0] == 'admin.xlsx' or file_name[0] == 'huiju.xlsx':
            continue
        device_name_list = db_session.query(Devices.a_device).filter_by(sheet_name = file_name[0]).group_by(Devices.a_device)
        workbook = xlwt.Workbook(encoding='utf-8')
        # device_name_list = [i[0] for i in device_name_list]
        # device_name_dict = {}       #POD1-25G接入交换机-锐捷S6510 [1, 10, 11, 12, 13, 14, 15, 16, 17, 18, 2, 3, 4, 5, 6, 7, 8, 9]
        # for device_name in device_name_list:
        #     device_name_split_list = device_name.split('-')
        #     device_name_prefix = '-'.join(device_name_split_list[:-1])
        #     exist_device_name = device_name_dict.get(device_name_prefix)
        #     if exist_device_name:
        #         exist_device_name.append(int(device_name_split_list[-1]))
        #         device_name_dict[device_name_prefix] = exist_device_name
        #     else:
        #         device_name_dict[device_name_prefix] = [int(device_name_split_list[-1])]
        #
        #
        # for device_name_prefix, device_num_list in device_name_dict.items():
        #     device_num_list = sorted(device_num_list) #排序

        for select_device_name in device_name_list:
            # select_device_name = device_name_prefix + '-' + str(device_num)
            select_device_name = select_device_name[0]
            worksheet = workbook.add_sheet(select_device_name)  #生成sheet
            worksheet.write(0, 0, "A端设备所在机房")
            worksheet.write(0, 1, "A端设备所在机柜")
            worksheet.write(0, 2, "A端设备所在U位", )
            worksheet.write(0, 3, "A端设备")
            worksheet.write(0, 4, "A端物理端口")
            worksheet.write(0, 5, "A端端口类型")
            worksheet.write(0, 6, "Z端设备所在机房")
            worksheet.write(0, 7, "Z端设备所在机柜", )
            worksheet.write(0, 8, "Z端设备所在U位", )
            worksheet.write(0, 9, "Z端设备", )
            worksheet.write(0, 10, "Z端物理端口", )
            print(select_device_name)
            asyncio.ensure_future(change_port_excel(worksheet, select_device_name))
            await asyncio.sleep(0)
        workbook.save(file_name[0].split('.')[0] + 'report.xlsx')


if __name__ == '__main__':
    # 插入设备位置
    # Base.metadata.drop_all(engine)
    # Base.metadata.create_all(engine)
    # device_position = read_position('./zhenzhou/device_position.xlsx')
    ## device_position = read_position('./shanxi/device_position.xlsx')
    # db_session.execute(DevicePosition.__table__.insert(), device_position)
    # db_session.commit()

    Devices.__table__.drop(engine)
    Devices.__table__.create(engine)
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())

