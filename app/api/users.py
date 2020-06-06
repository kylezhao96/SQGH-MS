from flask import jsonify
from sqlalchemy import func

from app import db
from app.api import bp
from app.models import User


@bp.route('/getusers', methods=['GET'])
def get_users():
    res = []
    companys1 = db.session.query(User.company).filter(User.company=='石桥子风电场').distinct().all()
    companys2 = db.session.query(User.company).filter(User.company!='石桥子风电场').distinct().all()
    companys = companys1+companys2
    for item in companys:
        company = item[0]
        users_by_company = db.session.query(User).filter(User.company == company).all()
        users = []
        for user in users_by_company:
            users.append({
                'value':user.id,
                'label': user.name
            })
        res.append({
            'label':company,
            'options':users
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