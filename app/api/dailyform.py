from app.api import bp
from app.models import CalDailyForm
from flask import jsonify, request, url_for
from app.api.errors import bad_request
from app import db
# from flask_cors import cross_origin
from sqlalchemy import and_,or_
import pandas as pd
import xlrd
import datetime


@bp.route('/importcdf', methods=['GET'])
def import_cdf():
    excel_path='D:\\MyRepositories\\files\\2019年石桥风电场日报表.xlsx'
    cdf = pd.read_excel('D:\\MyRepositories\\files\\2019年石桥风电场日报表.xlsx', sheet_name='日报计算表',usecols=range(76),skiprows=range(3),header=None)
    data = {}
    for x in range(366):
        if cdf.loc[x].values[1] == 0.0:
            break
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
            data['fka32b'] = float(cdf.loc[x].values[10])
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
            cdf2 = CalDailyForm()
            cdf2.from_dict(data)
            db.session.add(cdf2)
            db.session.commit()
            return jsonify(data)







