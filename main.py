# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 15:26

from utlis.read_position import read_position
from models import db_session, Devices, DevicePosition, engine, Base
from utlis.save_device import save_device
import asyncio
import xlwt

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
                    db_session.commit()
                    break


        index += 1
        worksheet.write(index, 0, a_room)
        worksheet.write(index, 1, a_cabinet_name)
        worksheet.write(index, 2, a_u)
        worksheet.write(index, 3, a_device_name)
        worksheet.write(index, 4, a_port)
        worksheet.write(index, 5, z_room)
        worksheet.write(index, 6, z_cabinet_name)
        worksheet.write(index, 7, z_u)
        worksheet.write(index, 8, z_device_name)
        worksheet.write(index, 9, z_port)


async def main():
    # data =  save_device(['./zhenzhou/PODbak.xlsx'])
    data =  save_device(['./zhenzhou/nine.xlsx'])
    # data =  save_device(['./shanxi/admin.xlsx'])
    for data_info in data:
        # 数据入库
        future = asyncio.ensure_future(save_device_data(data_info))
        await asyncio.sleep(0)

    #生成报表
    file_name_list = db_session.query(Devices.sheet_name).group_by(Devices.sheet_name)
    for file_name in file_name_list:
        if file_name[0] == 'admin.xlsx':
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
            print(select_device_name)
            worksheet = workbook.add_sheet(select_device_name)  #生成sheet
            worksheet.write(0, 0, "A端设备所在机房")
            worksheet.write(0, 1, "A端设备所在机柜")
            worksheet.write(0, 2, "A端设备所在U位", )
            worksheet.write(0, 3, "A端设备", )
            worksheet.write(0, 4, "A端物理端口", )
            worksheet.write(0, 5, "Z端设备所在机房")
            worksheet.write(0, 6, "Z端设备所在机柜", )
            worksheet.write(0, 7, "Z端设备所在U位", )
            worksheet.write(0, 8, "Z端设备", )
            worksheet.write(0, 9, "Z端物理端口", )
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

