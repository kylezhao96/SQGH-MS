import datetime
import os
import re
from copy import copy

import numpy as np
import openpyxl

import pandas as pd
from flask import jsonify, request
from sqlalchemy import or_

from app import db
from app.api import bp
from app.models import WT, WTMaintain, User, Gzp
from app.api.users import get_user_id, get_user
from app.tool.tool import realRound
from app.api.dailyform import EXCEL_PATH


@bp.route('/createwtm', methods=['POST'])
def create_wtm():
    """
    使用此函数创建维护单
    """
    data = request.get_json() or {}
    print(data)
    # wtm = WTMaintain()
    gzp = Gzp()
    wtm = WTMaintain()
    # 暂时设定一张维护单对应一张工作票
    if not User.query.filter_by(name=data['manager']).first():  # 工作负责人不存在
        manager = User()
        manager.name = data['manager']
        db.session.add(manager)
        db.session.commit()
        gzp.manage_person_id = manager.id
    else:
        gzp.manage_person_id = User.query.filter_by(name=data['manager']).first().id
    wt_regex = re.compile(r'A(\d){,2}')
    gzp.wt_id = WT.query.filter_by(id=wt_regex.search(data['wt'][1]).group(1)).first().id
    gzp.task = data['task']
    wtm.type = data['type']
    # wtm.allow_time = datetime.datetime.fromtimestamp(data['allow_time']/1000)
    members = []
    for item in data['members']:
        if not User.query.filter_by(name=item).first():  # 工作班成员不存在
            member = User()
            member.name = item
            db.session.add(member)
            db.session.commit()
            members.append(members)
        else:
            members.append(User.query.filter_by(name=item).first())
    gzp.members = members
    db.session.add(wtm)
    db.session.commit()
    gzp.wtm_id = wtm.id
    db.session.add(gzp)
    db.session.commit()

    return jsonify("ok")


@bp.route('/getgzps', methods=['GET'])
def get_gzps():
    """
    获取所有今日工作票单及未终结工作票
    """
    # today_gzps = Gzp.query.filter(or_(Gzp.pstart_time > datetime.date.today(), Gzp.is_end == False))
    today_gzps = Gzp.query.all()
    data = []
    for item in today_gzps:
        members = ''
        for member in item.members:
            members = members + member.name + '，'
        members = members[:-1]
        wts = ''
        wtms = list()
        index = 0
        for wt in item.wts:
            wtms.append({})
            if item.wtms.all():
                for wtm in item.wtms:
                    if wt.id == wtm.wt_id:
                        wtms[index] = {
                            'wt_id': wt.id,
                            'stop_time': datetime.datetime.strftime(wtm.stop_time, '%Y-%m-%d %H:%M'),
                            'start_time': '' if not wtm.start_time else datetime.datetime.strftime(wtm.start_time,
                                                                                                   '%Y-%m-%d %H:%M'),
                            'lost_power': wtm.lost_power,
                            'time': wtm.time
                        }
            else:
                wtms[index] = {
                    'wt_id': wt.id,
                }
            wts = wts + 'A' + str(wt.id) + '，'
            index = index + 1
        wts = wts[:-1]
        x = {
            'id': item.gzp_id,
            'wt_id': wts,
            'manager': User.query.filter_by(id=item.manage_person_id).first().name,
            'task': item.task,
            'members': members,
            'pstart_time': datetime.datetime.strftime(item.pstart_time, '%Y-%m-%d %H:%M'),
            'pstop_time': datetime.datetime.strftime(item.pstop_time, '%Y-%m-%d %H:%M'),
            'wtms': wtms,
            'is_end': item.is_end,
            'error_code': item.error_code
        }
        data.append(x)
        index = index + 1
    print(data)
    return jsonify(data)


@bp.route('/getwttasks', methods=['GET'])
def get_wt_tasks():
    """
    获取所有创建过的任务
    """
    tasks = Gzp.query.with_entities(Gzp.task).distinct().all()
    re_tasks = []
    for item in tasks:
        x = {}
        re_tasks.append(
            {
                "value": item[0]
            }
        )
    print(re_tasks)
    return jsonify(re_tasks)


@bp.route('/delgzpdb', methods=['POST'])
def del_gzp_db():
    """
    删除工作票
    """
    data = request.get_json() or {}
    print(data)
    gzp = Gzp.query.filter_by(gzp_id=data['id']).first()
    db.session.delete(gzp)
    db.session.commit()
    response = jsonify()
    response.status_code = 200
    return response


@bp.route('/delgzpcdf', methods=['POST'])
def del_gzp_cdf():
    """
    删除日报表中维护记录
    """
    data = request.get_json() or {}
    wtms = data['wtms']
    workbook = openpyxl.load_workbook(EXCEL_PATH)
    response = jsonify()
    response.status_code = 202  # 202代表未找到
    if 'error_code' in data:  # 故障
        wtm_type = '故障'
    else:
        wtm_type = '维护'
    worksheet = workbook['风机' + wtm_type + '统计']
    for wtm in wtms:
        if wtm['wt_id'] <= 20:  # 一号线
            #  在此写入
            start_col = 1
        elif wtm_type == '故障':  # 二号线 故障
            start_col = 11
        else:
            start_col = 9
        for row_num in range(1, worksheet.max_row):
            if type(worksheet.cell(row_num, start_col + 4).value) == datetime.datetime:  # 判断为时间类型才进入循环
                if wtm_type == '故障':
                    if worksheet.cell(row_num, start_col + 4).value.strftime('%Y-%m-%d %H:%M') == wtm['stop_time'] \
                            and worksheet.cell(row_num, start_col + 5).value.strftime('%Y-%m-%d %H:%M') == wtm[
                        'start_time']:
                        worksheet.cell(row_num, start_col, '')
                        worksheet.cell(row_num, start_col + 1, '')
                        worksheet.cell(row_num, start_col + 2, '')
                        worksheet.cell(row_num, start_col + 3, '')
                        worksheet.cell(row_num, start_col + 4, '')
                        worksheet.cell(row_num, start_col + 5, '')
                        worksheet.cell(row_num, start_col + 6, '')
                        worksheet.cell(row_num, start_col + 7, '')
                        worksheet.cell(row_num, start_col + 8, '')
                        response.status_code = 200  # 代表找到
                        break
                else:
                    if worksheet.cell(row_num, start_col + 2).value.strftime('%Y-%m-%d %H:%M') == wtm['stop_time'] \
                            and worksheet.cell(row_num, start_col + 3).value.strftime('%Y-%m-%d %H:%M') == wtm[
                        'start_time']:
                        worksheet.cell(row_num, start_col, '')
                        worksheet.cell(row_num, start_col + 1, '')
                        worksheet.cell(row_num, start_col + 2, '')
                        worksheet.cell(row_num, start_col + 3, '')
                        worksheet.cell(row_num, start_col + 4, '')
                        worksheet.cell(row_num, start_col + 5, '')
                        worksheet.cell(row_num, start_col + 6, '')
                        response.status_code = 200  # 代表找到
                        break
    workbook.save(EXCEL_PATH)
    return response


@bp.route('/postgzp', methods=['POST'])
def post_gzp():
    data = request.files
    data.get('file').save('temp/gzp.xls')  # 将前端发来的文件暂存
    data_gzp = pd.read_excel('temp/gzp.xls')  # 读取
    if not Gzp.query.filter_by(gzp_id=data_gzp.loc[1].values[13]).first():
        gzp = Gzp()
    else:
        gzp = Gzp.query.filter_by(gzp_id=data_gzp.loc[1].values[13]).first()
    # 以下开始对各项数据进行读取
    gzp.firm = data_gzp.loc[1].values[1]
    gzp.gzp_id = data_gzp.loc[1].values[13]
    gzp.manage_person = get_user(data_gzp.loc[3].values[4])
    members = re.split("\W+", data_gzp.loc[6].values[0])
    members_temp = []
    for member in members:
        members_temp.append(get_user(member))
    gzp.members = members_temp
    if not pd.isnull(data_gzp.loc[9].values[5]):
        gzp.error_code = data_gzp.loc[9].values[5]
        gzp.error_content = re.match(r'(处理)?(\w+)', data_gzp.loc[11].values[10]).group(2)
    gzp_wts_id = list(
        map(lambda x: re.match(r'^(A)(\d+)$', x).group(2), re.findall(re.compile(r'A\d+'), data_gzp.loc[11].values[0])))
    gzp.wts = list(map(lambda x: WT.query.filter_by(id=int(x)).first(), gzp_wts_id))  # wt放在最后
    gzp.postion = data_gzp.loc[11].values[5]
    gzp.task = data_gzp.loc[11].values[10]
    gzp.pstart_time = datetime.datetime(data_gzp.loc[15].values[2], data_gzp.loc[15].values[4],
                                        data_gzp.loc[15].values[6], data_gzp.loc[15].values[8],
                                        data_gzp.loc[15].values[10])
    gzp.pstop_time = datetime.datetime(data_gzp.loc[16].values[2], data_gzp.loc[16].values[4],
                                       data_gzp.loc[16].values[6], data_gzp.loc[16].values[8],
                                       data_gzp.loc[16].values[10])
    start_flag = 20
    while True:  # flag 指向签发行
        if data_gzp.loc[start_flag].values[0] == '8、签发人签名':
            break
        start_flag = start_flag + 1
    gzp.sign_person = get_user(data_gzp.loc[start_flag].values[2])
    sign_time_year = data_gzp.loc[start_flag].values[7]
    sign_time_month = data_gzp.loc[start_flag].values[9]
    sign_time_day = data_gzp.loc[start_flag].values[11]
    sign_time_hour = data_gzp.loc[start_flag].values[13]
    sign_time_minutes = data_gzp.loc[start_flag].values[15]
    gzp.sign_time = datetime.datetime(sign_time_year, sign_time_month, sign_time_day, sign_time_hour, sign_time_minutes)
    db.session.add(gzp)
    db.session.commit()

    os.remove('temp/gzp.xls')  # 删除暂存文件
    return jsonify('ok')


@bp.route('/wtmstodb', methods=['POST'])
def wtms2db():
    """
    将风机维护数据写入数据库
    """
    data = request.get_json() or {}
    response = jsonify()
    for wtm in data['wtms']:
        if not WTMaintain.query.filter(WTMaintain.gzp_id == data['id'], WTMaintain.wt_id == wtm['wt_id']).first():
            wtm_db = WTMaintain()
        else:
            wtm_db = WTMaintain.query.filter(WTMaintain.gzp_id == data['id'], WTMaintain.wt_id == wtm['wt_id']).first()
        wtm_db.wt_id = wtm['wt_id']
        gzp = Gzp.query.filter_by(gzp_id=data['id']).first()
        gzp.wtms.append(wtm_db)
        gzp.is_end = True
        if gzp.error_code:  # 故障
            wtm_db.error_code = gzp.error_code
        else:
            wtm_db.task = gzp.task
        stop_time = datetime.datetime.strptime(wtm['stop_time'], '%Y-%m-%d %H:%M')
        wtm_db.stop_time = datetime.datetime(stop_time.year, stop_time.month, stop_time.day, stop_time.hour,
                                             stop_time.minute)

        if 'start_time' in wtm.keys():
            if wtm['start_time']:
                start_time = datetime.datetime.strptime(wtm['start_time'], '%Y-%m-%d %H:%M')
                wtm_db.start_time = datetime.datetime(start_time.year, start_time.month, start_time.day, start_time.hour,
                                                      start_time.minute)
                wtm_db.time = realRound((wtm_db.start_time - wtm_db.stop_time).seconds / 3600, 2)
                wtm_db.lost_power = float(wtm['lost_power'])
            else:
                gzp.is_end = False
                response.status_code = 201
        else:
            gzp.is_end = False
            response.status_code = 201
        db.session.add(wtm_db)
        db.session.commit()
        db.session.add(gzp)
        db.session.commit()
    return response


@bp.route('/wtmstocdf', methods=['POST'])
def wtms2cdf():
    data = request.get_json() or {}
    gzp_id = data['gzp_id']
    workbook = openpyxl.load_workbook(EXCEL_PATH)
    gzp = Gzp.query.filter_by(gzp_id=gzp_id).first()
    this_month = False
    flag = 0
    if gzp.error_code:  # 故障
        wtm_type = '故障'
    else:
        wtm_type = '维护'
    worksheet = workbook['风机' + wtm_type + '统计']
    for wtm in gzp.wtms:
        if wtm.wt_id <= 20:  # 一号线
            #  在此写入
            start_col = 1
        elif wtm_type == '故障':  # 二号线 故障
            start_col = 11
        else:
            start_col = 9
        for row_num in range(1, worksheet.max_row):
            month = wtm.stop_time.month
            if row_num == flag:
                continue
            if re.findall(r'石桥一期(\d)月风机' + wtm_type + '统计', str(worksheet.cell(row_num, 1).value)):
                if str(month) == re.findall(r'石桥一期(\d)月风机' + wtm_type + '统计', worksheet.cell(row_num, 1).value)[0]:
                    #  定位到当月标题
                    this_month = True
                    flag = row_num + 1
            if this_month and worksheet.cell(row_num, 1).value == '合计':
                worksheet.insert_rows(row_num, amount=1)  # 若没有空行了，插入一行
                for x in range(1, 20):  # 复制样式
                    worksheet.cell(row_num, x)._style = copy(worksheet.cell(row_num - 1, x)._style)
                    worksheet.cell(row_num, x).font = copy(worksheet.cell(row_num - 1, x).font)
                    worksheet.cell(row_num, x).border = copy(worksheet.cell(row_num - 1, x).border)
                    worksheet.cell(row_num, x).fill = copy(worksheet.cell(row_num - 1, x).fill)
                    worksheet.cell(row_num, x).number_format = copy(worksheet.cell(row_num - 1, x).number_format)
                    worksheet.cell(row_num, x).protection = copy(worksheet.cell(row_num - 1, x).protection)
                    worksheet.cell(row_num, x).alignment = copy(worksheet.cell(row_num - 1, x).alignment)
            if this_month and worksheet.cell(row_num, start_col).value in [None, '']:
                if wtm_type == '故障':
                    worksheet.cell(row_num, start_col,
                                   'A' + str(wtm.wt_id) + ' ' + str(WT.query.filter_by(id=wtm.wt_id).first().dcode))
                    worksheet.cell(row_num, start_col + 1, gzp.error_code)
                    worksheet.cell(row_num, start_col + 2, gzp.task)
                    worksheet.cell(row_num, start_col + 4, wtm.stop_time)
                    worksheet.cell(row_num, start_col + 5, wtm.start_time)
                    worksheet.cell(row_num, start_col + 6, wtm.time)
                    worksheet.cell(row_num, start_col + 7, wtm.lost_power)
                    # worksheet.cell(row_num, start_col+2, gzp.error_code)
                else:
                    worksheet.cell(row_num, start_col, 'A' + str(wtm.wt_id) + ' ' + str(
                        WT.query.filter_by(id=wtm.wt_id).first().dcode))
                    worksheet.cell(row_num, start_col + 1, '其他')
                    worksheet.cell(row_num, start_col + 2, gzp.task)
                    worksheet.cell(row_num, start_col + 3, wtm.stop_time)
                    worksheet.cell(row_num, start_col + 4, wtm.start_time)
                    worksheet.cell(row_num, start_col + 5, wtm.time)
                    worksheet.cell(row_num, start_col + 6, wtm.lost_power)
                break
    workbook.save(EXCEL_PATH)
    response = jsonify()
    response.status_code = 200
    # response.headers['Location'] = url_for('api.', id=task.id)
    return response


@bp.route('/changecdf', methods=['POST'])
def change_cdf():
    data = request.get_json() or {}
    wtms = data['new']['wtms']
    wtms_pre = data['old']['wtms']
    workbook = openpyxl.load_workbook(EXCEL_PATH)
    flag = False
    if 'error_code' in data['old']:  # 故障
        wtm_type = '故障'
    else:
        wtm_type = '维护'
    worksheet = workbook['风机' + wtm_type + '统计']
    for index, wtm in enumerate(wtms):
        stop_time = datetime.datetime.strptime(wtm['stop_time'], '%Y-%m-%d %H:%M')
        start_time = datetime.datetime.strptime(wtm['start_time'], '%Y-%m-%d %H:%M')
        if wtm['wt_id'] <= 20:  # 一号线
            #  在此写入
            start_col = 1
        elif wtm_type == '故障':  # 二号线 故障
            start_col = 11
        else:
            start_col = 9
        for row_num in range(1, worksheet.max_row):
            if type(worksheet.cell(row_num, start_col + 4).value) == datetime.datetime:  # 判断为时间类型才进入循环
                if wtm_type == '故障':
                    if worksheet.cell(row_num, start_col + 4).value.strftime('%Y-%m-%d %H:%M') == wtms_pre[index][
                        'stop_time'] \
                            and worksheet.cell(row_num, start_col + 5).value.strftime('%Y-%m-%d %H:%M') == \
                            wtms_pre[index][
                                'start_time']:
                        worksheet.cell(row_num, start_col,
                                       'A' + str(wtm['wt_id']) + ' ' + str(
                                           WT.query.filter_by(id=wtm['wt_id']).first().dcode))
                        worksheet.cell(row_num, start_col + 1, data['new']['error_code'])
                        worksheet.cell(row_num, start_col + 2, data['new']['task'])
                        worksheet.cell(row_num, start_col + 4, stop_time)
                        worksheet.cell(row_num, start_col + 5, start_time)
                        worksheet.cell(row_num, start_col + 6, realRound((start_time - stop_time).seconds / 3600, 2))
                        worksheet.cell(row_num, start_col + 7, wtm['lost_power'])
                        # worksheet.cell(row_num, start_col+2, gzp.error_code)
                        flag = True
                        break
                else:
                    if worksheet.cell(row_num, start_col + 4).value.strftime('%Y-%m-%d %H:%M') == wtms_pre[index][
                        'stop_time'] \
                            and worksheet.cell(row_num, start_col + 5).value.strftime('%Y-%m-%d %H:%M') == \
                            wtms_pre[index][
                                'start_time']:
                        worksheet.cell(row_num, start_col, 'A' + str(wtm['wt_id']) + ' ' + str(
                            WT.query.filter_by(id=wtm['wt_id']).first().dcode))
                        worksheet.cell(row_num, start_col + 1, '其他')
                        worksheet.cell(row_num, start_col + 2, data['new']['task'])
                        worksheet.cell(row_num, start_col + 3, stop_time)
                        worksheet.cell(row_num, start_col + 4, start_time)
                        worksheet.cell(row_num, start_col + 5, realRound((start_time - stop_time).seconds / 3600, 2))
                        worksheet.cell(row_num, start_col + 6, wtm['lost_power'])
                        flag = True
                        break
    workbook.save(EXCEL_PATH)
    return jsonify(flag)


@bp.route('/wtmsyn', methods=['GET'])
def gzp_syn():
    whtj = pd.read_excel(EXCEL_PATH, sheet_name='风机维护统计', usecols=range(16), header=None).fillna('')
    gztj = pd.read_excel(EXCEL_PATH, sheet_name='风机故障统计', usecols=range(20), header=None).fillna('')
    # 维护
    for x in range(len(whtj)):
        if re.findall(r'^(A)(\d+)(\s*)(\d{5})$', whtj.loc[x].values[0]):
            if WTMaintain.query.filter(WTMaintain.stop_time == whtj.loc[x].values[3],
                                       WTMaintain.start_time == whtj.loc[x].values[4]).first():
                wtm = WTMaintain.query.filter(WTMaintain.stop_time == whtj.loc[x].values[3],
                                              WTMaintain.start_time == whtj.loc[x].values[4]).first()
            else:
                wtm = WTMaintain()
            wtm.wt_id = int(re.match(r'^(A)(\d+)(\s*)(\d{5})$', whtj.loc[3].values[0]).group(2))
            wtm.type = whtj.loc[x].values[1]
            wtm.task = whtj.loc[x].values[2]
            wtm.stop_time = whtj.loc[x].values[3]
            wtm.start_time = whtj.loc[x].values[4]
            wtm.time = realRound((wtm.start_time - wtm.stop_time).seconds / 3600, 2)
            wtm.lost_power = realRound(float(whtj.loc[x].values[6]), 4)
            db.session.add(wtm)
            db.session.commit()
        if re.findall(r'^(A)(\d+)(\s*)(\d{5})$', whtj.loc[x].values[8]):
            if WTMaintain.query.filter(WTMaintain.stop_time == whtj.loc[x].values[11],
                                       WTMaintain.start_time == whtj.loc[x].values[12]).first():
                wtm = WTMaintain.query.filter(WTMaintain.stop_time == whtj.loc[x].values[11],
                                              WTMaintain.start_time == whtj.loc[x].values[12]).first()
            else:
                wtm = WTMaintain()
            wtm.wt_id = int(re.match(r'^(A)(\d+)(\s*)(\d{5})$', whtj.loc[x].values[8]).group(2))
            wtm.type = whtj.loc[x].values[9]
            wtm.task = whtj.loc[x].values[10]
            wtm.stop_time = whtj.loc[x].values[11]
            wtm.start_time = whtj.loc[x].values[12]
            wtm.time = realRound((wtm.start_time - wtm.stop_time).seconds / 3600, 2)
            wtm.lost_power = realRound(float(whtj.loc[x].values[13]), 4)
            db.session.add(wtm)
            db.session.commit()
    # 故障
    for x in range(len(gztj)):
        if re.findall(r'^(A)(\d+)(\s*)(\d{5})$', gztj.loc[x].values[0]):
            if WTMaintain.query.filter(WTMaintain.stop_time == gztj.loc[x].values[4],
                                       WTMaintain.start_time == gztj.loc[x].values[5]).first():
                wtm = WTMaintain.query.filter(WTMaintain.stop_time == gztj.loc[x].values[4],
                                              WTMaintain.start_time == gztj.loc[x].values[5]).first()
            else:
                wtm = WTMaintain()
            wtm.wt_id = int(re.match(r'^(A)(\d+)(\s*)(\d{5})$', gztj.loc[x].values[0]).group(2))
            wtm.error_code = gztj.loc[x].values[1]
            wtm.error_content = gztj.loc[x].values[2]
            wtm.type = gztj.loc[x].values[3]
            wtm.stop_time = gztj.loc[x].values[4]
            wtm.start_time = gztj.loc[x].values[5]
            wtm.time = realRound((wtm.start_time - wtm.stop_time).seconds / 3600, 2)
            wtm.lost_power = realRound(float(gztj.loc[x].values[7]), 4)
            wtm.error_approach = gztj.loc[x].values[8]
            db.session.add(wtm)
            db.session.commit()
        if re.findall(r'^(A)(\d+)(\s*)(\d{5})$', gztj.loc[x].values[10]):
            if WTMaintain.query.filter(WTMaintain.stop_time == gztj.loc[x].values[14],
                                       WTMaintain.start_time == gztj.loc[x].values[15]).first():
                wtm = WTMaintain.query.filter(WTMaintain.stop_time == gztj.loc[x].values[14],
                                              WTMaintain.start_time == gztj.loc[x].values[15]).first()
            else:
                wtm = WTMaintain()
            wtm.wt_id = int(re.match(r'^(A)(\d+)(\s*)(\d{5})$', gztj.loc[x].values[10]).group(2))
            wtm.error_code = gztj.loc[x].values[11]
            wtm.error_content = gztj.loc[x].values[12]
            wtm.type = gztj.loc[x].values[13]
            wtm.stop_time = gztj.loc[x].values[14]
            wtm.start_time = gztj.loc[x].values[15]
            wtm.time = realRound((wtm.start_time - wtm.stop_time).seconds / 3600, 2)
            wtm.lost_power = realRound(float(gztj.loc[x].values[17]), 4)
            wtm.error_approach = gztj.loc[x].values[18]
            db.session.add(wtm)
            db.session.commit()
    return jsonify({})
