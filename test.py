# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 15:27
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
#     # def get_sheet_name(self, index):
#     #     if not getattr(self, 'inputFile', None) or not getattr(self, 'workbook', None):
#     #         raise AttributeError('not find excel')
#     #     return self.workbook.sheets()[index]
# excel_path = ExcelResolver()
# excel_path.set_file_path = r'C:\Users\asus\Desktop\user.xlsx'
# excel_path.getWorkSheet(1)
# dict_data = {
# 	'user': 'B',
# 	'en_name': 'C',
# 	'role': 'D',
# 	'email': 'E'
# }
# data = excel_path.convert_excel_data_to_dict(0, dict_data)
data = [{'user': 'liws', 'en_name': '李为帅', 'role': 1.0, 'email': 'liws@asiainfo.com'}, {'user': 'wangxh11', 'en_name': '王晓辉', 'role': 1.0, 'email': 'wangxh11@asiainfo.com'}, {'user': 'wanglq7', 'en_name': '王丽倩', 'role': 2.0, 'email': 'wanglq7@asiainfo.com'}, {'user': 'liuyue7', 'en_name': '刘悦', 'role': 5.0, 'email': 'liuyue7@asiainfo.com'}, {'user': 'yangjg3', 'en_name': '杨建国', 'role': 1.0, 'email': 'yangjg3@asiainfo.com'}, {'user': 'leicw', 'en_name': '雷传伟', 'role': 2.0, 'email': 'leicw@asiainfo.com'}, {'user': 'wanghui16', 'en_name': '王辉', 'role': 2.0, 'email': 'wanghui16@asiainfo.com'}, {'user': 'quhui', 'en_name': '曲惠', 'role': 2.0, 'email': 'quhui@asiainfo.com'}, {'user': 'hanzw3', 'en_name': '韩志伟', 'role': 2.0, 'email': 'hanzw3@asiainfo.com'}, {'user': 'zhangwb', 'en_name': '张文波', 'role': 1.0, 'email': 'zhangwb@asiainfo.com'}, {'user': 'macl3', 'en_name': '马春磊', 'role': 2.0, 'email': 'macl3@asiainfo.com'}, {'user': 'zhouty', 'en_name': '周天怡', 'role': 2.0, 'email': 'zhouty@asiainfo.com'}, {'user': 'zhangyue3', 'en_name': '张玥', 'role': 5.0, 'email': 'zhangyue3@asiainfo.com'}, {'user': 'qiaolj', 'en_name': '乔丽娟', 'role': 5.0, 'email': 'qiaolj@asiainfo.com'}, {'user': 'zhangsj5', 'en_name': '张胜军', 'role': 1.0, 'email': 'zhangsj5@asiainfo.com'}, {'user': 'liyc5', 'en_name': '李一晨', 'role': 1.0, 'email': 'liyc5@asiainfo.com'}, {'user': 'wangaq3', 'en_name': '王安琪', 'role': 1.0, 'email': 'wangaq3@asiainfo.com'}, {'user': 'ligz5', 'en_name': '李国柱', 'role': 1.0, 'email': 'ligz5@asiainfo.com'}, {'user': 'wangjl13', 'en_name': '王嘉乐', 'role': 1.0, 'email': 'wangjl13@asiainfo.com'}, {'user': 'tanzj', 'en_name': '澹志军', 'role': 1.0, 'email': 'tanzj@asiainfo.com'}, {'user': 'liyh25', 'en_name': '李彦辉', 'role': 1.0, 'email': 'liyh25@asiainfo.com'}, {'user': 'haoshuang', 'en_name': '郝爽', 'role': 6.0, 'email': 'haoshuang@asiainfo.com'}, {'user': 'wangxt8', 'en_name': '王学婷', 'role': 1.0, 'email': 'wangxt8@asiainfo.com'}, {'user': 'sunyl3', 'en_name': '孙钰林', 'role': 1.0, 'email': 'sunyl3@asiainfo.com'}, {'user': 'nalige', 'en_name': '哪力格尔', 'role': 1.0, 'email': 'nalige@asiainfo.com'}, {'user': 'gaoqian', 'en_name': '高芊', 'role': 1.0, 'email': 'gaoqian@asiainfo.com'}, {'user': 'chenbo6', 'en_name': '陈波', 'role': 1.0, 'email': 'chenbo6@asiainfo.com'}, {'user': 'renxh3', 'en_name': '任笑涵', 'role': 5.0, 'email': 'renxh3@asiainfo.com'}, {'user': 'yangdb', 'en_name': '杨东波', 'role': 1.0, 'email': 'yangdb@asiainfo.com'}, {'user': 'zhangxn', 'en_name': '张晓楠', 'role': 3.0, 'email': 'zhangxn@asiainfo.com'}, {'user': 'liqian9', 'en_name': '李谦', 'role': 4.0, 'email': 'liqian9@asiainfo.com'}, {'user': 'limy13', 'en_name': '黎明阳', 'role': 4.0, 'email': 'limy13@asiainfo.com'}, {'user': 'ligh5', 'en_name': '李光辉', 'role': 4.0, 'email': 'ligh5@asiainfo.com'}, {'user': 'lish6', 'en_name': '李仕豪', 'role': 4.0, 'email': 'lish6@asiainfo.com'}, {'user': 'yinsl', 'en_name': '尹森林', 'role': 4.0, 'email': 'yinsl@asiainfo.com'}, {'user': 'wangyz5', 'en_name': '王宇哲', 'role': 3.0, 'email': 'wangyz5@asiainfo.com'}, {'user': 'liue', 'en_name': '刘娥', 'role': 4.0, 'email': 'liue@asiainfo.com'}, {'user': 'shixa', 'en_name': '石新澳', 'role': 4.0, 'email': 'shixa@asiainfo.com'}, {'user': 'dengmm', 'en_name': '邓蒙蒙', 'role': 6.0, 'email': 'dengmm@asiainfo.com'}, {'user': 'zhangze', 'en_name': '张泽', 'role': 3.0, 'email': 'zhangze@asiainfo.com'}, {'user': 'doush', 'en_name': '豆世红', 'role': 3.0, 'email': 'doush@asiainfo.com'}, {'user': 'mazhe3', 'en_name': '马喆', 'role': 3.0, 'email': 'mazhe3@asiainfo.com'}, {'user': 'lihui7', 'en_name': '李辉', 'role': 3.0, 'email': 'lihui7@asiainfo.com'}, {'user': 'caoym', 'en_name': '曹延敏', 'role': 3.0, 'email': 'caoym@asiainfo.com'}, {'user': 'liujing', 'en_name': '刘静', 'role': 3.0, 'email': 'liujing@asiainfo.com'}, {'user': 'xiening', 'en_name': '谢宁', 'role': 3.0, 'email': 'xiening@asiainfo.com'}, {'user': 'xingyd', 'en_name': '邢译丹', 'role': 3.0, 'email': 'xingyd@asiainfo.com'}, {'user': 'chenjy26', 'en_name': '陈佳宇', 'role': 1.0, 'email': 'chenjy26@asiainfo.com'}, {'user': 'hanye', 'en_name': '韩冶', 'role': 2.0, 'email': 'hanye@asiainfo.com'}, {'user': 'gulei3', 'en_name': '谷磊', 'role': 4.0, 'email': 'gulei3@asiainfo.com'}, {'user': 'changlc', 'en_name': '常岚超', 'role': 4.0, 'email': 'changlc@asiainfo.com'}, {'user': 'yixl', 'en_name': '伊晓璐', 'role': 1.0, 'email': 'yixl@asiainfo.com'}, {'user': 'maxy3', 'en_name': '马欣悦', 'role': 5.0, 'email': 'maxy3@asiainfo.com'}, {'user': 'libf', 'en_name': '李碧发', 'role': 1.0, 'email': 'libf@asiainfo.com'}, {'user': 'zhuanglk', 'en_name': '庄林凯', 'role': 1.0, 'email': 'zhuanglk@asiainfo.com'}, {'user': 'shish3', 'en_name': '石帅豪', 'role': 4.0, 'email': 'shish3@asiainfo.com'}, {'user': 'guojie', 'en_name': '郭杰', 'role': 4.0, 'email': 'guojie@asiainfo.com'}]

UserInfo(username = 'liujz2', email='liujz2@asiainfo.com', password='123', cn_name='刘剑铮').save()
for i in data:
    role = i['role']
    if role == 1 or role == 2:
        role = str(int(role))
        UserInfo(username=i['user'], role_id=role, email=i['email'], cn_name=i['en_name'], organization_id=2).save()
    else:
        UserInfo(username=i['user'], email=i['email'], cn_name=i['en_name'], organization_id=2).save()
