from app.api import bp
from app.models import User, DailyTask, MonthlyTask
from flask import jsonify, request, url_for
from app.api.errors import bad_request
from app import db
# from flask_cors import cross_origin
from sqlalchemy import and_,or_
import pyperclip


@bp.route('/dailytasks', methods=['GET'])
def get_dailytasks():
    tasks = DailyTask.query.filter().order_by(DailyTask.index)
    # data = Task.to_collection_dict(tasks, 1, 20, '/task')
    data = DailyTask.to_col_dict(tasks)
    print(data)
    return jsonify(data)


@bp.route('/monthlytasks', methods=['GET'])
def get_monthlytasks():
    tasks = MonthlyTask.query.filter()
    # data = Task.to_collection_dict(tasks, 1, 20, '/task')
    data = MonthlyTask.to_col_dict(tasks)
    print(data)
    return jsonify(data)


@bp.route('/tasks', methods=['POST'])
def create_task():
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
    data = request.get_json() or {}
    pyperclip.copy(data['info'])
    info = pyperclip.paste()
    response = jsonify(info)
    response.status_code = 201
    return response
# update