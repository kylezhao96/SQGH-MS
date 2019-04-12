from app.api import bp
from app.models import User, Task
from flask import jsonify, request, url_for
from app.api.errors import bad_request
from app import db
from flask_cors import cross_origin
from sqlalchemy import and_,or_


@bp.route('/tasks/day=<int:day>', methods=['GET'])
@cross_origin()
def get_tasksofday(day):
    tasks = Task.query.filter(or_(Task.day == day, Task.day == 0)).order_by(Task.time)
    # data = Task.to_collection_dict(tasks, 1, 20, '/task')
    data = Task.to_col_dict(tasks)
    return jsonify(data)


@bp.route('/tasks', methods=['POST'])
@cross_origin()
def create_task():
    data = request.get_json() or {}
    if 'name' not in data or 'day' not in data or 'time' not in data:
        return bad_request('缺少必要项！')
    data['time'] = int(data['time'].split(':')[0])
    print(data)
    if Task.query.filter_by(name=data['name']).first():
        if Task.query.filter_by(name=data['name']).first().time==data['time']:
            return bad_request('任务已存在！')
    task = Task()
    task.from_dict(data)
    db.session.add(task)
    db.session.commit()
    response = jsonify(task.to_dict())
    response.status_code = 201
    # response.headers['Location'] = url_for('api.', id=task.id)
    return response
