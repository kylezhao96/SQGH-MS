from app.api import bp
from app.models import WT, WTMaintain, User
from flask import jsonify, request, url_for
from app.api.errors import bad_request
from app import db
from sqlalchemy import and_,or_
import re


@bp.route('/getwts', methods=['GET'])
def get_wts():
    options = []
    wts = []
    for n in range(1, 5):
        wts.append(WT.query.filter(WT.line == n))
        options.append({
            'value': 'line' + str(n),
            'label': '集电线' + str(n),
            'children': []
        })
        for i in wts[n-1]:
            x = {
                'value': 'A' + str(i.id),
                'label': 'A' + str(i.id) + '风机'
            }
            options[n-1]['children'].append(x)
    print(options)
    return jsonify(options)


@bp.route('/createwtm', methods=['POST'])
def create_wtm():
    data = request.get_json() or {}
    print(data)
    wtm = WTMaintain()
    wtm.manager_id = User.query.filter_by(name=data['manager']).first().id
    wt_regex = re.compile(r'A(\d){,2}')
    wtm.wt_id = WT.query.filter_by(id=wt_regex.search(data['wt'][1]).group(1)).first().id
    wtm.task = data['task']
    wtm.type = data['type']
    db.session.add(wtm)
    db.session.commit()
    return jsonify("ok")


@bp.route('/getwtms', methods=['GET'])
def get_wtms():
    unstoped_wtm = WTMaintain.query.filter_by(is_end=0)
    data = []
    for item in unstoped_wtm:
        x = {
            'id': item.wt_id,
            'manager': User.query.filter_by(id=item.manager_id).first().name,
            'task': item.task,
            'members': item.members,
            'allow_time': item.allow_time
        }
        data.append(x)
    print(data)
    return jsonify(data)


@bp.route('/getwttasks', methods=['GET'])
def get_wt_tasks():
    tasks = WTMaintain.query.with_entities(WTMaintain.task).distinct().all()
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
