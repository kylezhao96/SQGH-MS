from app.api import bp
from app.models import User, Task
from flask import jsonify, request, url_for
from app.api.errors import bad_request
from app import db


@bp.route('/tasks/<int:day>', methods=['GET'])
def get_tasks_byday(day):
    tasks = Task.query.filter_by(day=day)
    page = request.args.get('page', 1, type=int)
    per_page = min(request.args.get('per_page', 10, type=int), 100)
    data = Task.to_collection_dict(tasks, page, per_page, 'api.get_tasks_byday')
    return jsonify(data)


@bp.route('/tasks', methods=['POST'])
def create_task():
    data = request.get_json() or {}
    print(data)
    if 'name' not in data or 'day' not in data or 'time' not in data:
        return bad_request('缺少必要项！')
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

