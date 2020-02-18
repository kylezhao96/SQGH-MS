from flask import jsonify

from app import db
from app.api import bp
from app.models import User


@bp.route('/getusers', methods=['GET'])
def get_users():
    res = []
    users = User.query.all()
    for i in users:
        res.append({
            'value': i.name,
            'label': i.name
        })
    print(res)
    return jsonify(res)


def get_user_id(name):
    if not User.query.filter_by(name=name).first():
        manager = User()
        manager.name = name
        db.session.add(manager)
        db.session.commit()
        return manager.id
    else:
        return User.query.filter_by(name=name).first().id


def get_user(name):
    if not User.query.filter_by(name=name).first():
        manager = User()
        manager.name = name
        db.session.add(manager)
        db.session.commit()
        return manager
    else:
        return User.query.filter_by(name=name).first()