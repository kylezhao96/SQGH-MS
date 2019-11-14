from app import db
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship  # 创建关系


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
            }
            # '_links': {
            #     'self': url_for(endpoint, page=page, per_page=per_page,
            #                     **kwargs),
            #     'next': url_for(endpoint, page=page + 1, per_page=per_page,
            #                     **kwargs) if resources.has_next else None,
            #     'prev': url_for(endpoint, page=page - 1, per_page=per_page,
            #                     **kwargs) if resources.has_prev else None
            # }
        }
        return data

    @staticmethod
    def to_col_dict(query):
        data = {
            'items': [item.to_dict() for item in query]
        }
        return data


class User(PaginatedAPIMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(20), unique=True, nullable=False)
    oa_account = db.Column(db.String(20), unique=True)
    oa_password = db.Column(db.String(100))
    wtm = db.relationship('WTMaintain', backref='users', lazy='dynamic')

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


class DailyTask(PaginatedAPIMixin, db.Model):
    __tablename__ = 'dailytasks'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    hour = db.Column(db.Integer)
    minute = db.Column(db.Integer)
    index = db.Column(db.Integer, nullable=False)
    type = db.Column(db.String(100))

    def __repr__(self):
        return '<Task {}'.format(self.name)

    def to_dict(self):
        data = {
            'id': self.id,
            'name': self.name,
            'hour': self.hour,
            'minute': self.minute,
            'type': self.type
        }
        return data

    def from_dict(self, data):
        for field in ['name', 'hour', 'minute', 'type']:
            if field in data:
                setattr(self, field, data[field])


class MonthlyTask(PaginatedAPIMixin, db.Model):
    __tablename__ = 'monthlytasks'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    day = db.Column(db.Integer)

    def __repr__(self):
        return '<Task {}'.format(self.name)

    def to_dict(self):
        data = {
            'id': self.id,
            'name': self.name,
            'day': self.day,
        }
        return data

    def from_dict(self, data):
        for field in ['name', 'day']:
            if field in data:
                setattr(self, field, data[field])

# 风机表
class WT(PaginatedAPIMixin, db.Model):
    __tablename__ = 'wts'
    id = db.Column(db.Integer, primary_key=True)
    line = db.Column(db.Integer)
    dcode = db.Column(db.Integer)
    type = db.Column(db.String(100), default='En121-2.5')
    wtm = db.relationship('WTMaintain', backref='wts', lazy='dynamic')

# 日报计算表
class CalDailyForm(PaginatedAPIMixin, db.Model):
    __tablename__ = 'caldailyform'
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, unique=True)
    # ka有功 kr有功 f正向 b反向
    fka312 = db.Column(db.Float)
    bka312 = db.Column(db.Float)
    fka313 = db.Column(db.Float)
    bka313 = db.Column(db.Float)
    fka322 = db.Column(db.Float)
    bka322 = db.Column(db.Float)
    fka323 = db.Column(db.Float)
    bka323 = db.Column(db.Float)
    fka31b = db.Column(db.Float)
    fka21b = db.Column(db.Float)
    fka311 = db.Column(db.Float)
    bka311 = db.Column(db.Float)
    fkr311 = db.Column(db.Float)
    bkr311 = db.Column(db.Float)
    fka321 = db.Column(db.Float)
    bka321 = db.Column(db.Float)
    fkr321 = db.Column(db.Float)
    bkr321 = db.Column(db.Float)
    bka111 = db.Column(db.Float)
    fka111 = db.Column(db.Float)
    # p 发电量 d 每日的 g总的 on上网的 off下网的 c场用的 l率
    dgp1 = db.Column(db.Integer)
    donp1 = db.Column(db.Integer)
    doffp1 = db.Column(db.Integer)
    dcp1 = db.Column(db.Integer)
    dcl1 = db.Column(db.Float)
    dgp2= db.Column(db.Integer)
    donp2 = db.Column(db.Integer)
    doffp2 = db.Column(db.Integer)
    dcp2 = db.Column(db.Integer)
    dcl2 = db.Column(db.Float)
    dgp = db.Column(db.Integer)
    donp = db.Column(db.Integer)
    doffp = db.Column(db.Integer)
    dcp = db.Column(db.Integer)
    dcl = db.Column(db.Float)
    doffp31b = db.Column(db.Integer)
    doffp21b = db.Column(db.Integer)
    #年的
    agp1 = db.Column(db.Integer)
    aonp1 = db.Column(db.Integer)
    aoffp1 = db.Column(db.Integer)
    acp1 = db.Column(db.Integer)
    acl1 = db.Column(db.Float)
    agp2 = db.Column(db.Integer)
    aonp2 = db.Column(db.Integer)
    aoffp2 = db.Column(db.Integer)
    acp2 = db.Column(db.Integer)
    acl2 = db.Column(db.Float)
    agp = db.Column(db.Integer)
    aonp = db.Column(db.Integer)
    aoffp = db.Column(db.Integer)
    acp = db.Column(db.Integer)
    acl = db.Column(db.Float)
    #月的
    mgp1 = db.Column(db.Integer)
    monp1 = db.Column(db.Integer)
    moffp1 = db.Column(db.Integer)
    mcp1 = db.Column(db.Integer)
    mcl1 = db.Column(db.Float)
    mgp2 = db.Column(db.Integer)
    monp2 = db.Column(db.Integer)
    moffp2 = db.Column(db.Integer)
    mcp2 = db.Column(db.Integer)
    mcl2 = db.Column(db.Float)
    mgp = db.Column(db.Integer)
    monp = db.Column(db.Integer)
    moffp = db.Column(db.Integer)
    mcp = db.Column(db.Integer)
    mcl = db.Column(db.Float)
    #svg ja有功功率 jr无功功率
    offja311 = db.Column(db.Integer)
    offjr311 = db.Column(db.Integer)
    offja321 = db.Column(db.Integer)
    offjr321 = db.Column(db.Integer)

    def __repr__(self):
        return '<CalDailyForm {}'.format(self.name)

    def to_dict(self):
        data = {
            # 'id': self.id,
            # 'name': self.name,
            # 'day': self.day,
        }
        return data

    def from_dict(self, data):
        for field in ['date', 'fka312', 'bka312', 'fka313', 'bka313', 'fka322', 'bka322', 'fka323', 'bka323', 'fka31b', 'fka21b', 'fka311', 'fkr311', 'bka311', 'bkr311', 'fka321', 'fkr321', 'bka321', 'bkr321', 'bka111', 'fka111', 'dgp1', 'donp1', 'doffp1', 'dcp1', 'dcl1', 'dgp2', 'donp2', 'doffp2', 'dcp2', 'dcl2', 'dgp', 'donp', 'doffp', 'dcp', 'dcl', 'doffp31b', 'doffp21b', 'agp1', 'aonp1', 'aoffp1', 'acp1', 'acl1', 'agp2', 'aonp2', 'aoffp2', 'acp2', 'acl2', 'agp', 'aonp', 'aoffp', 'acp', 'acl', 'mgp1', 'monp1', 'moffp1', 'mcp1', 'mcl1', 'mgp2', 'monp2', 'moffp2', 'mcp2', 'mcl2', 'mgp', 'monp', 'moffp', 'mcp', 'mcl', 'offja311', 'offjr311', 'offja321', 'offjr321']:
            if field in data:
                setattr(self, field, data[field])


# 风机维护记录
class WTMaintain(PaginatedAPIMixin, db.Model):
    __tablename__ = 'wtm'
    id = db.Column(db.Integer, primary_key=True)
    allow_time = db.Column(db.DateTime)  # 许可时间
    stop_time = db.Column(db.DateTime)  # 停机时间
    start_time = db.Column(db.DateTime)  # 启机时间
    end_time = db.Column(db.DateTime)  # 终结时间
    manager_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    members = db.Column(db.String(100))
    type = db.Column(db.String(100))  # 维护类型
    task = db.Column(db.String(100))  # 任务
    wt_id = db.Column(db.Integer, db.ForeignKey('wts.id'))
    lost_power = db.Column(db.Float)
    time = db.Column(db.Float)
    is_end = db.Column(db.Integer, default=0) # 是否终结，默认为0/false

    def from_dict(self, data):  # unfinished
        for field in ['manager_id', 'allow_time', 'stop_time', 'start_time', 'end_time', 'members', 'lost_power']:
            if field in data:
                setattr(self, field, data[field])
