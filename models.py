# -*- coding: utf-8 -*-
# author: hejianxin
# date: 2020/7/22 15:28
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.dialects.postgresql import INET,ARRAY,JSON # ARRAY类型必须用postgresql下面的
from sqlalchemy import Column, Integer, String, Text, ForeignKey, DateTime, SmallInteger, Index, Boolean ,func#, ARRAY
from sqlalchemy.orm import sessionmaker, relationship
from config import Config

Base = declarative_base()

engine = create_engine('mysql+pymysql://%s:%s@%s:%s/%s' % (Config.DB_USER,Config.DB_PASSWORD,Config.DB_HOST,Config.DB_PORT,Config.DB_NAME))
db_session = sessionmaker(bind=engine)
db_session = db_session()



class Devices(Base):
    __tablename__ = 'devices'  # 数据库表名称

    id = Column(Integer, primary_key=True)
    a_device = Column(String(128), index=True)
    a_cabinet_id = Column(Integer, ForeignKey('deviceposition.id'), nullable=True)
    a_port = Column(String(30))

    z_device = Column(String(128))
    z_cabinet_id = Column(Integer, ForeignKey('deviceposition.id'), nullable=True)
    z_port = Column(String(30), nullable=True)
    sheet_name = Column(String(80))
    is_use = Column(Boolean, default=False)
    cabinet_num = Column(String(2))

    a_device_position_id = relationship('DevicePosition', single_parent=True, foreign_keys=[a_cabinet_id])
    z_device_position_id = relationship('DevicePosition', single_parent=True, foreign_keys=[z_cabinet_id])


    def to_dict(self):
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}


class DevicePosition(Base):
    __tablename__ = 'deviceposition'

    id = Column(Integer, primary_key=True)
    device_name = Column(String(60), unique=True)
    cabinet = Column(String(50))
    u = Column(String(30))
    room = Column(String(20))


if __name__ == '__main__':
    Base.metadata.drop_all(engine)
    Base.metadata.create_all(engine)
    # Devices.__table__.drop(engine)
    # Devices.__table__.create(engine)




# class UserShoppingCart(db.Model):
#     __tablename__ = 'user_shopping_carts'
#     phone_number_section_id = db.Column(db.Integer,
#                                         db.ForeignKey('private_number_property_value.id', ondelete='CASCADE'),
#                                         index=True)
#     phone_number_use_time_id = db.Column(db.Integer,
#                                          db.ForeignKey('private_number_property_value.id', ondelete='CASCADE'),
#                                          index=True)
#     phone_number_amount_id = db.Column(db.Integer,
#                                        db.ForeignKey('private_number_property_value.id', ondelete='CASCADE'),
#                                        index=True)
#     phone_number_section = db.relationship('PrivateNumberPropertyValue', foreign_keys=[phone_number_section_id],
#                                            cascade='all, delete-orphan', single_parent=True)
#     phone_number_use_time = db.relationship('PrivateNumberPropertyValue', foreign_keys=[phone_number_use_time_id],
#                                             cascade='all, delete-orphan', single_parent=True)
#     phone_number_amount = db.relationship('PrivateNumberPropertyValue', foreign_keys=[phone_number_amount_id],
#                                           cascade='all, delete-orphan', single_parent=True)














































