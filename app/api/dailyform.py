from app.api import bp
from app.models import CalDailyForm
from flask import jsonify, request, url_for
from app.api.errors import bad_request
from app import db
# from flask_cors import cross_origin
from sqlalchemy import and_,or_
import pandas as pd
import xlutils, xlrd, openpyxl
import datetime
import xlwings as xl
import pythoncom

from app.tool.tool import realRound
EXCEL_PATH = "C:\\Users\\Kyle\\Desktop\\2019年石桥风电场日报表.xlsx"


def save_excel(path):
    pythoncom.CoInitialize()
    app = xl.App(visible=False)
    book = app.books.open(path)
    book.save()
    app.kill()


@bp.route('/updatefullcdf', methods=['GET'])
def import_cdf():
    cdf = pd.read_excel(EXCEL_PATH, sheet_name='日报计算表', usecols=range(76), skiprows=range(3), header=None)
    # cdf = df_from_excel(EXCEL_PATH, '日报计算表', 76, 3, None)
    response = []
    for x in range(366):
        data = {}
        if x == 0:
            data['date'] = datetime.datetime(2018, 12, 31, 0, 0)
            data['fka312'] = float(cdf.loc[x].values[1] )
            data['bka312'] = float(cdf.loc[x].values[2] )
            data['fka313'] = float(cdf.loc[x].values[3] )
            data['bka313'] = float(cdf.loc[x].values[4] )
            data['fka322'] = float(cdf.loc[x].values[5] )
            data['bka322'] = float(cdf.loc[x].values[6] )
            data['fka323'] = float(cdf.loc[x].values[7] )
            data['bka323'] = float(cdf.loc[x].values[8] )
            data['fka31b'] = float(cdf.loc[x].values[9] )
            data['fka21b'] = float(cdf.loc[x].values[10])
            data['fka311'] = float(cdf.loc[x].values[11])
            data['bka311'] = float(cdf.loc[x].values[12])
            data['fkr311'] = float(cdf.loc[x].values[13])
            data['bkr311'] = float(cdf.loc[x].values[14])
            data['fka321'] = float(cdf.loc[x].values[15])
            data['bka321'] = float(cdf.loc[x].values[16])
            data['fkr321'] = float(cdf.loc[x].values[17])
            data['bkr321'] = float(cdf.loc[x].values[18])
            data['bka111'] = float(cdf.loc[x].values[19])
            data['fka111'] = float(cdf.loc[x].values[20])
        else:
            if cdf.loc[x].values[0] > datetime.datetime(2019, 9, 30, 0, 0):
                # if cdf.loc[x].values[0] >= datetime.now():
                break
            data['date'] = cdf.loc[x].values[0]
            data['fka312'] = float(cdf.loc[x].values[1])
            data['bka312'] = float(cdf.loc[x].values[2])
            data['fka313'] = float(cdf.loc[x].values[3])
            data['bka313'] = float(cdf.loc[x].values[4])
            data['fka322'] = float(cdf.loc[x].values[5])
            data['bka322'] = float(cdf.loc[x].values[6])
            data['fka323'] = float(cdf.loc[x].values[7])
            data['bka323'] = float(cdf.loc[x].values[8])
            data['fka31b'] = float(cdf.loc[x].values[9])
            data['fka21b'] = float(cdf.loc[x].values[10])
            data['fka311'] = float(cdf.loc[x].values[11])
            data['bka311'] = float(cdf.loc[x].values[12])
            data['fkr311'] = float(cdf.loc[x].values[13])
            data['bkr311'] = float(cdf.loc[x].values[14])
            data['fka321'] = float(cdf.loc[x].values[15])
            data['bka321'] = float(cdf.loc[x].values[16])
            data['fkr321'] = float(cdf.loc[x].values[17])
            data['bkr321'] = float(cdf.loc[x].values[18])
            data['bka111'] = float(cdf.loc[x].values[19])
            data['fka111'] = float(cdf.loc[x].values[20])
            data['dgp1'] = cdf.loc[x].values[21]
            data['donp1'] = cdf.loc[x].values[22]
            data['doffp1'] = cdf.loc[x].values[23]
            data['dcp1'] = cdf.loc[x].values[24]
            data['dcl1'] = cdf.loc[x].values[25]
            data['dgp2'] = cdf.loc[x].values[26]
            data['donp2'] = cdf.loc[x].values[27]
            data['doffp2'] = cdf.loc[x].values[28]
            data['dcp2'] = cdf.loc[x].values[29]
            data['dcl2'] = cdf.loc[x].values[30]
            data['dgp'] = cdf.loc[x].values[31]
            data['donp'] = cdf.loc[x].values[32]
            data['doffp'] = cdf.loc[x].values[33]
            data['dcp'] = cdf.loc[x].values[34]
            data['dcl'] = cdf.loc[x].values[35]
            data['doffp31b'] = cdf.loc[x].values[36]
            data['doffp21b'] = cdf.loc[x].values[37]
            data['agp1'] = cdf.loc[x].values[38]
            data['aonp1'] = cdf.loc[x].values[39]
            data['aoffp1'] = cdf.loc[x].values[40]
            data['acp1'] = cdf.loc[x].values[41]
            data['acl1'] = cdf.loc[x].values[42]
            data['agp2'] = cdf.loc[x].values[43]
            data['aonp2'] = cdf.loc[x].values[44]
            data['aoffp2'] = cdf.loc[x].values[45]
            data['acp2'] = cdf.loc[x].values[46]
            data['acl2'] = cdf.loc[x].values[47]
            data['agp'] = cdf.loc[x].values[48]
            data['aonp'] = cdf.loc[x].values[49]
            data['aoffp'] = cdf.loc[x].values[50]
            data['acp'] = cdf.loc[x].values[51]
            data['acl'] = cdf.loc[x].values[52]
            data['mgp1'] = cdf.loc[x].values[53]
            data['monp1'] = cdf.loc[x].values[54]
            data['moffp1'] = cdf.loc[x].values[55]
            data['mcp1'] = cdf.loc[x].values[56]
            data['mcl1'] = cdf.loc[x].values[57]
            data['mgp2'] = cdf.loc[x].values[58]
            data['monp2'] = cdf.loc[x].values[59]
            data['moffp2'] = cdf.loc[x].values[60]
            data['mcp2'] = cdf.loc[x].values[61]
            data['mcl2'] = cdf.loc[x].values[62]
            data['mgp'] = cdf.loc[x].values[63]
            data['monp'] = cdf.loc[x].values[64]
            data['moffp'] = cdf.loc[x].values[65]
            data['mcp'] = cdf.loc[x].values[66]
            data['mcl'] = cdf.loc[x].values[67]
            data['offja311'] = cdf.loc[x].values[69]
            data['offjr311'] = cdf.loc[x].values[71]
            data['offja321'] = cdf.loc[x].values[73]
            data['offjr321'] = cdf.loc[x].values[75]
        response.append(data)
        cdf2 = CalDailyForm()
        cdf2.from_dict(data)
        db.session.add(cdf2)
        db.session.commit()
    return jsonify(response)


@bp.route('/addtodaycdf', methods=["POST"])
def add_cdf():
    cdf = pd.read_excel(EXCEL_PATH, sheet_name='日报计算表', usecols=range(76), skiprows=range(3), header=None)
    data = {}
    x = (datetime.datetime(2019, 10, 3) - datetime.datetime(datetime.datetime.now().year, 1, 1)).days+1
    data['date'] = cdf.loc[x].values[0]
    data['fka312'] = float(cdf.loc[x].values[1])
    data['bka312'] = float(cdf.loc[x].values[2])
    data['fka313'] = float(cdf.loc[x].values[3])
    data['bka313'] = float(cdf.loc[x].values[4])
    data['fka322'] = float(cdf.loc[x].values[5])
    data['bka322'] = float(cdf.loc[x].values[6])
    data['fka323'] = float(cdf.loc[x].values[7])
    data['bka323'] = float(cdf.loc[x].values[8])
    data['fka31b'] = float(cdf.loc[x].values[9])
    data['fka21b'] = float(cdf.loc[x].values[10])
    data['fka311'] = float(cdf.loc[x].values[11])
    data['bka311'] = float(cdf.loc[x].values[12])
    data['fkr311'] = float(cdf.loc[x].values[13])
    data['bkr311'] = float(cdf.loc[x].values[14])
    data['fka321'] = float(cdf.loc[x].values[15])
    data['bka321'] = float(cdf.loc[x].values[16])
    data['fkr321'] = float(cdf.loc[x].values[17])
    data['bkr321'] = float(cdf.loc[x].values[18])
    data['bka111'] = float(cdf.loc[x].values[19])
    data['fka111'] = float(cdf.loc[x].values[20])
    data['dgp1'] = cdf.loc[x].values[21]
    data['donp1'] = cdf.loc[x].values[22]
    data['doffp1'] = cdf.loc[x].values[23]
    data['dcp1'] = cdf.loc[x].values[24]
    data['dcl1'] = cdf.loc[x].values[25]
    data['dgp2'] = cdf.loc[x].values[26]
    data['donp2'] = cdf.loc[x].values[27]
    data['doffp2'] = cdf.loc[x].values[28]
    data['dcp2'] = cdf.loc[x].values[29]
    data['dcl2'] = cdf.loc[x].values[30]
    data['dgp'] = cdf.loc[x].values[31]
    data['donp'] = cdf.loc[x].values[32]
    data['doffp'] = cdf.loc[x].values[33]
    data['dcp'] = cdf.loc[x].values[34]
    data['dcl'] = cdf.loc[x].values[35]
    data['doffp31b'] = cdf.loc[x].values[36]
    data['doffp21b'] = cdf.loc[x].values[37]
    data['agp1'] = cdf.loc[x].values[38]
    data['aonp1'] = cdf.loc[x].values[39]
    data['aoffp1'] = cdf.loc[x].values[40]
    data['acp1'] = cdf.loc[x].values[41]
    data['acl1'] = cdf.loc[x].values[42]
    data['agp2'] = cdf.loc[x].values[43]
    data['aonp2'] = cdf.loc[x].values[44]
    data['aoffp2'] = cdf.loc[x].values[45]
    data['acp2'] = cdf.loc[x].values[46]
    data['acl2'] = cdf.loc[x].values[47]
    data['agp'] = cdf.loc[x].values[48]
    data['aonp'] = cdf.loc[x].values[49]
    data['aoffp'] = cdf.loc[x].values[50]
    data['acp'] = cdf.loc[x].values[51]
    data['acl'] = cdf.loc[x].values[52]
    data['mgp1'] = cdf.loc[x].values[53]
    data['monp1'] = cdf.loc[x].values[54]
    data['moffp1'] = cdf.loc[x].values[55]
    data['mcp1'] = cdf.loc[x].values[56]
    data['mcl1'] = cdf.loc[x].values[57]
    data['mgp2'] = cdf.loc[x].values[58]
    data['monp2'] = cdf.loc[x].values[59]
    data['moffp2'] = cdf.loc[x].values[60]
    data['mcp2'] = cdf.loc[x].values[61]
    data['mcl2'] = cdf.loc[x].values[62]
    data['mgp'] = cdf.loc[x].values[63]
    data['monp'] = cdf.loc[x].values[64]
    data['moffp'] = cdf.loc[x].values[65]
    data['mcp'] = cdf.loc[x].values[66]
    data['mcl'] = cdf.loc[x].values[67]
    data['offja311'] = cdf.loc[x].values[69]
    data['offjr311'] = cdf.loc[x].values[71]
    data['offja321'] = cdf.loc[x].values[73]
    data['offjr321'] = cdf.loc[x].values[75]
    cdf2 = CalDailyForm()
    cdf2.from_dict(data)
    db.session.add(cdf2)
    db.session.commit()


@bp.route('/fillcdf', methods=["POST"])
def fill_caldailyform():
    data = request.get_json() or {}
    # data['data'] = datetime.datetime.combine(datetime.date.today(), datetime.time(0, 0, 0))
    print(data)
    workbook = openpyxl.load_workbook(EXCEL_PATH)
    worksheet = workbook['日报计算表']
    # rowNum = (datetime.datetime.now()- datetime.datetime(datetime.datetime.now().year, 1, 1)).days+5
    rowNum = (datetime.datetime(2019, 10, 3) - datetime.datetime(datetime.datetime.now().year, 1, 1)).days+5
    worksheet.cell(rowNum, 2, data['fka312'])  #312正向
    worksheet.cell(rowNum, 3, data['bka312'])  # 312反向
    worksheet.cell(rowNum, 4, data['fka313'])  # 313正向
    worksheet.cell(rowNum, 5, data['bka313'])  # 313反向
    worksheet.cell(rowNum, 6, data['fka322'])  # 322正向
    worksheet.cell(rowNum, 7, data['bka322'])  # 322反向
    worksheet.cell(rowNum, 8, data['fka323'])  # 323正向
    worksheet.cell(rowNum, 9, data['bka323'])  # 323正向
    worksheet.cell(rowNum, 10, data['fka31b'])  # 31b正向
    worksheet.cell(rowNum, 11, 836.27)  # 21b正向
    worksheet.cell(rowNum, 12, 0)  # 311正向有
    worksheet.cell(rowNum, 13, data['bka311'])  # 311反向有
    worksheet.cell(rowNum, 14, 0)  # 311正向无
    worksheet.cell(rowNum, 15, data['bkr311'])  # 311反向无
    worksheet.cell(rowNum, 16, 0)  # 321正向有
    worksheet.cell(rowNum, 17, data['bka321'])  # 321反向有
    worksheet.cell(rowNum, 18, 0)  # 321正向无
    worksheet.cell(rowNum, 19, data['bkr321'])  # 321反向无
    worksheet.cell(rowNum, 20, data['bka111'])  # 111反向
    worksheet.cell(rowNum, 21, data['fka111'])  # 111正向
    workbook.save(EXCEL_PATH)
    save_excel(EXCEL_PATH)
    add_cdf()   # 保存至数据库
    response = jsonify(data)
    response.status_code = 201
    return response









