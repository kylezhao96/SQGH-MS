import os
from dotenv import load_dotenv


basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(basedir, '.env'))
OMS_PATH = r"C:\Users\Administrator\Desktop\1报表文件夹\日报表\2020年\2020年OMS日报.xlsx"
TY_PATH = r"C:\Users\Administrator\Desktop\1报表文件夹\每日00：30前石桥风电场每日风机电量、风速统计表报送诸城桃园风场公共邮箱\2020年\石桥风电场报送每日风机电量风速统计表 2020.xlsx"
EXCEL_PATH = r"C:\Users\Administrator\Desktop\1报表文件夹\日报表\2020年\2020年石桥风电场日报表.xlsx"
driverLoc = r"D:\submitTable\driver\IEDriverServer.exe"
DESK_PATH = r"C:\Users\Administrator\Desktop"

class Config(object):
    # 最好用环境变量方式设置密钥
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'hard to guess'
    DIALECT = 'mysql'
    DRIVER = 'pymysql'
    USERNAME = 'kylezhao'
    PASSWORD = '123456'
    HOST = '47.93.199.183'
    PORT = '3306'
    DATABASE = 'sqghdb'

    SQLALCHEMY_DATABASE_URI = '{}+{}://{}:{}@{}:{}/{}?charset=utf8'.format(
        DIALECT, DRIVER, USERNAME, PASSWORD, HOST, PORT, DATABASE
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = True
    SQLALCHEMY_COMMIT_ON_TEARDOWN = True

    SQLALCHEMY_POOL_SIZE = 10
    SQLALCHEMY_MAX_OVERFLOW = 5
