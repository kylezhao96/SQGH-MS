# from flask_cors import cross_origin
import pyperclip
from flask import jsonify, request

from app import db
from app.api import bp
from app.api.errors import bad_request
from app.models import DailyTask, MonthlyTask


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
    print(data)
    sum = data['num1']+data['num2']+data['num3']
    pyperclip.copy(str(data['hour'])+':00：石桥风电场出力'+data['power']+'MW，风速'+data['windspeed']+'m/s，'+data['windir']+'风，风机停运共'+str(sum)+'台(维护'+str(data['num1'])+'台，故障'+str(data['num2'])+'台，无通讯'+str(data['num3'])+'台)，无输变电设备停电。')
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