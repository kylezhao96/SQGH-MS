from app import db
from flask import url_for


class PaginatedAPIMixin(object):
    @staticmethod
    def to_collection_dict(query, page, per_page, endpoint, **kwargs):
        resources = query.paginate(page, per_page, False)
        data = {
            'items': [item.to_dict() for item in resources.items],
            '_meta': {
                'page': page,
                'per_page': per_page,
                'total_pages': resources.pages,
                'total_items': resources.total
            },
            '_links': {
                'self': url_for(endpoint, page=page, per_page=per_page,
                                **kwargs),
                'next': url_for(endpoint, page=page + 1, per_page=per_page,
                                **kwargs) if resources.has_next else None,
                'prev': url_for(endpoint, page=page - 1, per_page=per_page,
                                **kwargs) if resources.has_prev else None
            }
        }
        return data


class User(PaginatedAPIMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(20), unique=True, nullable=False)
    oa_account = db.Column(db.Integer, unique=True, nullable=False)
    oa_password = db.Column(db.String(100), nullable=False)

    def __repr__(self):
        return '<User {}>'.format(self.name)

    def to_dict(self):
        data = {
            'id': self.id,
            'name': self.name,
            'oa_account': self.oa_account,
            'oa_password': self.oa_password
        }
        return data

    def form_dict(self, data):
        for field in ['name', 'oa_account', 'oa_password']:
            if field in data:
                setattr(self, field, data[field])


class Task(PaginatedAPIMixin, db.Model):
    __tablename__ = 'tasks'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    day = db.Column(db.Integer)
    time = db.Column(db.Integer)

    def __repr__(self):
        return '<Task {}'.format(self.name)

    def to_dict(self):
        data = {
            'id': self.id,
            'name': self.name,
            'day': self.day,
            'time': self.time
        }
        return data

    def from_dict(self, data):
        for field in ['name', 'day', 'time']:
            if field in data:
                setattr(self, field, data[field])

