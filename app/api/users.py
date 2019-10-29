from app.api import bp
from app.models import User
from flask import jsonify, request, url_for
from app.api.errors import bad_request
from app import db
from sqlalchemy import and_,or_
import pyperclip


@bp.route('/getusers', methods=['GET'])
def get_users():
    users = []
    res = []
    users = User.query.all()
    for i in users:
        res.append({
            'value': i.name,
            'label': i.name
        })
    print(res)
    return jsonify(res)

