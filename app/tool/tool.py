import calendar
import datetime
import json
import re
from decimal import Decimal, ROUND_HALF_UP
from app.models import WTMaintain, Gzp
from sqlalchemy import or_

OMS_PATH = "C:\\Users\\Kyle\\Desktop\\2020年OMS日报.xlsx"
TY_PATH = "C:\\Users\\Kyle\\Desktop\\石桥风电场报送每日风机电量风速统计表 2020.xlsx"


def realRound(d, n=0):
    format = '0.'
    while (n):
        format = format + '0'
        n = n - 1
    return Decimal(str(d)).quantize(Decimal(format), rounding=ROUND_HALF_UP)


# 获取第一天和最后一天
def getFirstAndLastDay(today):
    # 获取当前年份
    year = today.year
    # 获取当前月份
    month = today.month
    # 获取当前月的第一天的星期和当月总天数
    weekDay, monthCountDay = calendar.monthrange(year, month)
    # 获取当前月份第一天
    firstDay = datetime.date(year, month, day=1)
    # 获取当前月份最后一天
    lastDay = datetime.date(year, month, day=monthCountDay)
    # 返回第一天和最后一天
    return firstDay, lastDay


# DecimalEncoder解码器
class DecimalEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Decimal):
            return float(obj)
        return super(DecimalEncoder, self).default(obj)


# 用于字符串分割
# def div_string():
#     pattern = re.compile("")


# 获取今日维护、故障停机时间
def get_stop_time(date):
    wtms = WTMaintain.query.filter(
        or_(WTMaintain.start_time >= date, WTMaintain.stop_time >= date, WTMaintain.start_time.is_(None))).all()
    sum_time = {
        'gz': 0,
        'wh': 0,
    }

    for wtm in wtms:
        if Gzp.query.filter_by(gzp_id=wtm.gzp_id).first().error_code:
            key = 'gz'
        else:
            key = 'wh'
        print(wtm.stop_time.date())
        if wtm.stop_time.date() == date.date():  # 今日停机
            if wtm.start_time.date() == date.date():  # 今日启机
                sum_time[key] = sum_time[key] + wtm.time
            else:  # 尚未启机
                sum_time[key] = sum_time[key] + realRound(
                    (date + datetime.timedelta(1) - wtm.stop_time).seconds / 3600, 2)
        else:  # 非今日停机
            if wtm.start_time.date() == date.date():  # 今日启机
                sum_time[key] = sum_time[key] + realRound((wtm.start_time - date).seconds / 3600, 2)
            else:
                sum_time[key] = sum_time[key] + 24
    sum_time['sum'] = sum_time['gz'] + sum_time['wh']
    return sum_time


# 获取今日维护、故障停机损失电量
def get_lost_power(date):
    wtms = WTMaintain.query.filter(WTMaintain.start_time >= date,
                                   WTMaintain.start_time <= (date + datetime.timedelta(1))).all()
    res = {
        'gz': 0,
        'wh': 0,
    }

    for wtm in wtms:
        if Gzp.query.filter_by(gzp_id=wtm.gzp_id).first().error_code:
            key = 'gz'
        else:
            key = 'wh'
        if wtm.start_time.date() == date.date():  # 今日启机
            res[key] = res[key] + wtm.lost_power

    res['sum'] = res['gz'] + res['wh']
    return res


# 获取今日维护、故障停机次数
def get_num(date):
    wtms = WTMaintain.query.filter(
        or_(WTMaintain.start_time >= date, WTMaintain.stop_time >= date, WTMaintain.start_time.is_(None))).all()
    num= {
        'gz': [],
        'wh': [],
    }
    for x in range(40):
        num['gz'].append(0)
        num['wh'].append(0)
    for wtm in wtms:
        if Gzp.query.filter_by(gzp_id=wtm.gzp_id).first().error_code:
            key = 'gz'
        else:
            key = 'wh'
        num[key][wtm.wt_id-1] = num[key][wtm.wt_id-1] +1
    return num
