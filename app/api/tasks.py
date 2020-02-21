# from flask_cors import cross_origin
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging, time, datetime, openpyxl, os, sys, xlrd
from decimal import Decimal, ROUND_HALF_UP
import pyperclip
from flask import jsonify, request

from app import db
from app.api import bp
from app.api.errors import bad_request
from app.models import DailyTask, MonthlyTask
from app.tool.tool import realRound

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(message)s')
today = datetime.date.today()

@bp.route('/dailytasks', methods=['GET'])
def get_dailytasks():
    """
    获取每日定期工作
    """
    tasks = DailyTask.query.filter().order_by(DailyTask.index)
    # data = Task.to_collection_dict(tasks, 1, 20, '/task')
    data = DailyTask.to_col_dict(tasks)
    print(data)
    return jsonify(data)


@bp.route('/monthlytasks', methods=['GET'])
def get_monthlytasks():
    """
    获取今日月度工作
    """
    tasks = MonthlyTask.query.filter()
    data = MonthlyTask.to_col_dict(tasks)
    print(data)
    return jsonify(data)


@bp.route('/tasks', methods=['POST'])
def create_task():
    """
    创建定期工作——本函数未引用
    """
    data = request.get_json() or {}
    if 'name' not in data or 'day' not in data or 'time' not in data:
        return bad_request('缺少必要项！')
    data['hour'] = int(data['time'].split(':')[0])
    data['minute']  = int(data['time'].split(':')[1])
    print(data)
    if DailyTask.query.filter_by(name=data['name']).first():
        if DailyTask.query.filter_by(name=data['name']).first().hour==data['hour'] and DailyTask.query.filter_by(name=data['name']).first().minute==data['minute']:
            return bad_request('任务已存在！')
    task = DailyTask()
    task.from_dict(data)
    db.session.add(task)
    db.session.commit()
    response = jsonify(task.to_dict())
    response.status_code = 201
    # response.headers['Location'] = url_for('api.', id=task.id)
    return response


@bp.route('/dotask', methods=['POST'])
def do_task():
    """
    本函数用于定时发送出力
    :param data: a list contain hour,power and wind speed
    :returns: status_code
    """
    data = request.get_json()or {}
    hour = datetime.datetime.now().hour
    this_hour = 0
    if hour<=8:
        this_hour = '08'
    elif hour <=12:
        this_hour = '12'
    elif hour <= 17:
        this_hour = '17'
    else :
        this_hour = '21'
    print(data)
    sum = data['num1']+data['num2']+data['num3']
    pyperclip.copy(this_hour+':00：石桥风电场出力'+data['power']+'MW，风速'+data['windspeed']+'m/s，'+data['windir']+'风，风机停运共'+str(sum)+'台(维护'+str(data['num1'])+'台，故障'+str(data['num2'])+'台，无通讯'+str(data['num3'])+'台)，无输变电设备停电。')
    info = pyperclip.paste()
    print(info)
    response = jsonify(info)
    response.status_code = 201
    return response

@bp.route('/fixclip', methods = ['POST'])
def fix_clip():
    """
    本函数用于对粘贴板内出力信息进行二次编辑
    """
    data = request.get_json() or {}
    pyperclip.copy(data['info'])
    info = pyperclip.paste()
    response = jsonify(info)
    response.status_code = 201
    return response
# update


@bp.route('/submitjtrb', methods = ['GET'])
def submit_jtrb():
    """
    上报集团公司日报
    """
    mon = today.strftime('%m')
    day = today.strftime('%d')
    year = today.strftime('%y')
    val = readExcel()
    driverLoc = "D:\submitTable\driver\IEDriverServer.exe"  # 注意在此配置IE驱动位置
    broswer = webdriver.Ie(driverLoc)
    # broswer.get("http://10.82.1.60/")
    broswer.get("http://10.82.1.60:8082/Rum-web/login.jsp")
    try:
        element = WebDriverWait(broswer, 10).until(
            EC.title_contains('国家能源集团数据报送平台')
        )
    except:
        logging.debug('请检查是否登录VPN')
        broswer.quit()
    un = broswer.find_element_by_id('username')
    un.clear()
    un.send_keys('20027658')
    psw = broswer.find_element_by_id('password')
    psw.send_keys('Sqghwh1234@')
    psw.send_keys(Keys.ENTER)
    try:
        element = WebDriverWait(broswer, 10).until(
            EC.presence_of_element_located((By.ID, "ext-comp-1010"))
        )
    except:
        logging.debug('Error!')
        broswer.quit()
    broswer.get(
        "http://10.82.1.60:8082/Rum-web/rum/rqreport/inputReport.jsp?issue=620"+year + mon + day + "&username=JT_DOMAIN_20027658&resourcesid=T_AS_D_YXGL_SCYXRB_TB_FD&state=1&nodeId=&processinstanceId=&ticket=SlRfRE9NQUlOXzIwMDI3NjU4&w=1881&h=780")
    try:
        element = WebDriverWait(broswer, 10).until(
            EC.title_contains('国华石桥生产运行情况日报表')
        )
    except:
        logging.debug('Error!')
        broswer.quit()
    # 运行容量
    # 检修容量
    # 备用容量
    # 临检容量
    # 新增容量
    for i in range(5):
        tb = broswer.find_element_by_id('report1_N' + str(i + 12))
        tb.click()
        eb = broswer.find_element_by_id('report1_editBox')
        eb.send_keys(str(val[i]))
    for i in range(5, 7):
        tb = broswer.find_element_by_id('report1_N' + str(i + 13))
        tb.click()
        eb = broswer.find_element_by_id('report1_editBox')
        eb.send_keys(str(val[i]))
    # 场用电量
    tb = broswer.find_element_by_id('report1_N21')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[7]))
    # 厂用电量
    tb = broswer.find_element_by_id('report1_N23')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[8]))
    # tb = broswer.find_element_by_id('report1_O22')
    # tb.click()
    # tb.send_keys(Keys.DOWN)
    # 上网电量
    tb = broswer.find_element_by_id('report1_N25')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[9]))
    # 基数外
    tb = broswer.find_element_by_id('report1_G27')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys('0')
    # 被替代
    tb = broswer.find_element_by_id('report1_G28')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys('0')
    # 网购电量
    tb = broswer.find_element_by_id('report1_G29')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[10]))
    # 平均风速
    tb = broswer.find_element_by_id('report1_N30')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[11]))
    # #限电量
    tb = broswer.find_element_by_id('report1_N31')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[12]))
    # tb = broswer.find_element_by_id('report1_O30')
    # tb.click()
    # tb.send_keys(Keys.DOWN)
    # 可用小时
    tb = broswer.find_element_by_id('report1_N40')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[13]))
    # 计划停运小时
    tb = broswer.find_element_by_id('report1_N42')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[14]))
    # 非计划停运小时
    tb = broswer.find_element_by_id('report1_N43')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[15]))
    # 非计划停运损失电量
    tb = broswer.find_element_by_id('report1_N44')
    tb.click()
    eb = broswer.find_element_by_id('report1_editBox')
    eb.send_keys(str(val[16]))
    # 姓名
    tb = broswer.find_element_by_id('report1_G47')
    tb.click()
    response = jsonify()
    response.status_code = 200
    return response


def calRowNum():
    dayDif = (datetime.datetime.now() - datetime.datetime( datetime.datetime.now().year, 1, 1)).days
    return dayDif + 4


def readExcel():  # 读取日报表
    val = []  # 定义返回值
    # os.chdir('C:\\Users\\Kyle\\Desktop') #更改当前目录
    # dataLoc = 'C:\\Users\\kylez\\Desktop'
    dataLoc = 'C:\\Users\\Administrator\\Desktop\\1报表文件夹\\日报表\\2020年'
    os.chdir(dataLoc)
    wb = openpyxl.load_workbook("2020年石桥风电场日报表.xlsx", data_only=True, read_only=True)
    rowNum = calRowNum()
    rbjs = wb.get_sheet_by_name('日报计算表')
    jtgs = wb.get_sheet_by_name('集团公司报表')
    yxrl_col = 5
    jxrl_col = 6
    byrl_col = 7
    ljrl_col = 8
    cydl_col = 9
    kyxs_col = 10
    jhtyxs_col = 11
    fjhtyxs_col = 12
    rxdl_col  = 13
    for x in range(1,15):
        if jtgs.cell(row=3,column=x).value == '运行容量':
            yxrl_col = x
        if jtgs.cell(row=3,column=x).value == '检修容量':
            jxrl_col = x
        if jtgs.cell(row=3,column=x).value == '备用容量':
            byrl_col = x
        if jtgs.cell(row=3,column=x).value == '临检容量':
            ljrl_col = x
        if jtgs.cell(row=3,column=x).value == '厂用电量':
            cydl_col = x
        if jtgs.cell(row=3,column=x).value == '可用小时':
            kyxs_col = x
        if jtgs.cell(row=3,column=x).value == '计划停运小时':
            jhtyxs_col = x
        if jtgs.cell(row=3,column=x).value == '非计划停运小时':
            fjhtyxs_col = x
        if jtgs.cell(row=3,column=x).value == '日限电量':
            rxdl_col = x
    val.append('0')  # 运行容量 0
    val.append(realRound(jtgs.cell(row=rowNum,column=jxrl_col).value, 1))  # 检修容量 1
    val.append(realRound(jtgs.cell(row=rowNum,column=byrl_col).value, 1))  # 备用容量 2
    val.append(realRound(jtgs.cell(row=rowNum,column=ljrl_col).value, 1))  # 临检容量 3
    val[0] = Decimal(100000) - val[1] - val[2] - val[3]  # 运行容量 3
    val.append('0')  # 新增容量 4
    val.append(realRound(rbjs['AF' + str(rowNum)].value / 10000, 4))  # 发电量 5
    val.append('0')  # 试运行发电量 6
    val.append(realRound(rbjs['AI' + str(rowNum)].value / 10000, 4))  # 场用电量 7
    val.append(realRound(jtgs.cell(row=rowNum,column=cydl_col).value, 4))  # 厂用电量 8
    val.append(realRound(rbjs['AG' + str(rowNum)].value / 10000, 4))  # 上网电量 9
    val.append(realRound(rbjs['AH' + str(rowNum)].value / 10000, 4))  # 下网电量 10
    val.append(readExcel2())  # 平均风速 11
    if jtgs.cell(row = rowNum,column=rxdl_col).value in [None,'']:
        val.append(0)
    else:
        val.append(jtgs.cell(row = rowNum,column=rxdl_col).value)  # 限电量   12
    val.append('24')  # 可用小时13
    val.append(realRound(jtgs.cell(rowNum,jhtyxs_col).value, 2))  # 计划停运小时14
    val.append(realRound(jtgs.cell(row=rowNum,column=fjhtyxs_col).value, 2))  # 非计划停运小时15
    gztj = wb.get_sheet_by_name('风机故障统计')
    val[13] = Decimal(24) - val[14] - val[15]
    # print(val)  # 测试用，时常使用时请注释掉
    ssdl = 0
    monthlist = [1, 19, 40, 66, 97, 122, 155, 184, 208, 233, 256, 278]
    month = datetime.datetime.now().month
    day = datetime.datetime.now().day
    sNum = monthlist[month - 1]  # 损失电量表中的行标
    found = False
    while sNum < 400:  # 寻找一期故障
        try:
            cellVal = gztj['F' + str(sNum)].value
        except BaseException as e:
            logging.debug(e)
        else:
            if type(cellVal) == datetime.datetime:  # 如果该行数据为字符串形式
                # dateSepa=dateRegex.findall(cellVal)   #正则表达式匹配
                # if dateSepa !=[]:                                   #匹配得到的值不为空
                if cellVal.month == month and cellVal.day == day - 1:
                    ssdl = ssdl + gztj['H' + str(sNum)].value
                    found = True
                else:
                    if found == True:
                        break
            else:
                if found == True:
                    break
        sNum = sNum + 1
    sNum = monthlist[month - 1]  # 初始化
    found = False  # 初始化
    while sNum < 400:  # 寻找二期故障
        try:
            cellVal = gztj['P' + str(sNum)].value
        except BaseException as e:
            logging.debug(e)
        else:
            if type(cellVal) == datetime.datetime:  # 如果该行数据为字符串形式
                if cellVal.month == month and cellVal.day == day - 1:
                    ssdl = ssdl + gztj['R' + str(sNum)].value
                    found = True
                else:
                    if found == True:
                        break
            else:
                if found == True:
                    break
        sNum = sNum + 1
    val.append(ssdl)  # 损失电量 16

    return val


def readExcel2():
    # 读取风速风量统计表
    dataLoc = 'C:\\Users\\Administrator\\Desktop\\1报表文件夹\\每日00：30前石桥风电场每日风机电量、风速统计表报送诸城桃园风场公共邮箱\\2020年'
    # dataLoc = 'C:\\Users\\kylez\\Desktop'
    os.chdir(dataLoc)
    wb = xlrd.open_workbook('石桥风电场报送每日风机电量风速统计表 2020.xlsx')
    # wb = openpyxl.load_workbook(filename = ,data_only=True)
    # fstj = wb.get_sheet_by_name('风速统计')
    fstj = wb.sheet_by_name(u'风速统计')
    date = fstj.col_values(0)
    windVel1 = fstj.col_values(1)
    windVel2 = fstj.col_values(2)
    rowNum = 2
    while True:
        cellVal = xlrd.xldate_as_tuple(date[rowNum] + 1, 0)
        if cellVal[1] == datetime.datetime.now().month and cellVal[2] == datetime.datetime.now().day:
            return realRound((windVel1[rowNum] + windVel2[rowNum]) / 2, 2)
        rowNum = rowNum + 1