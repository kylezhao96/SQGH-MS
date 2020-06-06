from app import db
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship  # 创建关系
from sqlalchemy import Column

# 中间表 工作班成员与工作票号 多对多关系
member_gzp = db.Table("member_gzp",
                      db.Column("member_id", db.Integer, db.ForeignKey("user.id")),
                      db.Column("gzp_id", db.Integer, db.ForeignKey("gzp.id"))
                      )

# 中间表 一台风机对应多张工作票，一张工作票也可以有多台风机
wt_gzp = db.Table("wt_gzp",
                  db.Column("wt_id", db.Integer, db.ForeignKey("wt.id")),
                  db.Column("gzp_id", db.Integer, db.ForeignKey("gzp.id"))
                  )


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
    __tablename__ = 'user'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(20), unique=True, nullable=False)
    oa_account = db.Column(db.String(20), unique=True)
    oa_password = db.Column(db.String(100))
    company = db.Column(db.String(100),default='其他')
    # 关联属性，多对多的情况，可以写在任意一个模型类中
    # gzps = db.relationship('Gzp', secondary=member_gzp,
    #                        backref='relate_gzp',
    #                        lazy='dynamic')

    # signed_gzp_id = db.Column(db.Integer, db.ForeignKey("gzp.id"))  # 签发的工作票 外键
    # signed_gzp = db.relationship('Gzp', foreign_keys=[signed_gzp_id])  # 签发的工作票
    #
    # managed_gzp_id = db.Column(db.Integer, db.ForeignKey("gzp.id"))  # 担任工作负责人的工作票 外键
    # managed_gzp = db.relationship('Gzp', foreign_keys=[managed_gzp_id])  # 担任负责人的工作票
    #
    # allowed1_gzp_id = db.Column(db.Integer, db.ForeignKey("gzp.id"))  # 值班许可的工作票
    # allowed1_gzp = db.relationship('Gzp', foreign_keys=[allowed1_gzp_id])
    #
    # allowed2_gzp_id = db.Column(db.Integer, db.ForeignKey("gzp.id"))  # 现场许可的工作票
    # allowed2_gzp = db.relationship('Gzp', foreign_keys=[allowed2_gzp_id])  # 现场许可的工作票

    # exec_gzps = db.relationship("Gzp", secondary=member_gzp, backref=db.backref('members', lazy='dynamic'), lazy="dynamic")  # 作为工作班成员执行的工作票

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
    __tablename__ = 'dailytask'
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
    __tablename__ = 'monthlytask'
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
    __tablename__ = 'wt'
    id = db.Column(db.Integer, primary_key=True)
    line = db.Column(db.Integer)
    dcode = db.Column(db.Integer)
    type = db.Column(db.String(100), default='En121-2.5')
    wtm = db.relationship('WTMaintain', backref="wt", lazy='dynamic')  # 一


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
    fka311 = db.Column(db.Float, default=0)
    bka311 = db.Column(db.Float)
    fkr311 = db.Column(db.Float, default=0)
    bkr311 = db.Column(db.Float)
    fka321 = db.Column(db.Float, default=0)
    bka321 = db.Column(db.Float)
    fkr321 = db.Column(db.Float, default=0)
    bkr321 = db.Column(db.Float)
    bka111 = db.Column(db.Float)
    fka111 = db.Column(db.Float)
    # p 发电量 d 每日的 g总的 on上网的 off下网的 c场用的 l率
    dgp1 = db.Column(db.Integer)
    donp1 = db.Column(db.Integer)
    doffp1 = db.Column(db.Integer)
    dcp1 = db.Column(db.Integer)
    dcl1 = db.Column(db.Float)
    dgp2 = db.Column(db.Integer)
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
    # 年的
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
    # 月的
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
    # svg ja有功功率 jr无功功率
    offja311 = db.Column(db.Integer)
    offjr311 = db.Column(db.Integer)
    offja321 = db.Column(db.Integer)
    offjr321 = db.Column(db.Integer)
    # s 风速
    dmaxs = db.Column(db.Float)
    dmins = db.Column(db.Float)
    davgs = db.Column(db.Float)
    dmaxs1 = db.Column(db.Float)
    dmins1 = db.Column(db.Float)
    davgs1 = db.Column(db.Float)
    dmaxs2 = db.Column(db.Float)
    dmins2 = db.Column(db.Float)
    davgs2 = db.Column(db.Float)
    # lp 限电
    dlp1 = db.Column(db.Float, default=0)  # 一期日限电量
    dlp2 = db.Column(db.Float, default=0)  # 二期日限电量
    dlp = db.Column(db.Float, default=0)  # 总日限电量
    mlp1 = db.Column(db.Float, default=0)  # 一期月限电量
    mlp2 = db.Column(db.Float, default=0)  # 二期月限电量
    mlp = db.Column(db.Float, default=0)  # 总月限电量
    alp1 = db.Column(db.Float, default=0)  # 一期年限电量
    alp2 = db.Column(db.Float, default=0)  # 二期年限电量
    alp = db.Column(db.Float, default=0)  # 总年限电量
    # l 负荷
    dmaxl = db.Column(db.Float)  # 日最大负荷
    dminl = db.Column(db.Float)  # 日最小负荷

    def __repr__(self):
        return '<CalDailyForm {}'.format(self.name)

    def to_dict(self):
        data = {}
        for name in dir(self):
            value = getattr(self, name)
            if not name.startswith('_') and (type(value) == int or type(value) == float):
                data[name] = value
        return data

    def from_dict(self, data):
        for field in dir(self):
            if field in data:
                setattr(self, field, data[field])


# 风机维护记录
class WTMaintain(PaginatedAPIMixin, db.Model):
    __tablename__ = 'wtm'
    id = db.Column(db.Integer, primary_key=True)
    wt_id = db.Column(db.Integer, db.ForeignKey('wt.id'))
    task = db.Column(db.String(100))  # 工作内容
    stop_time = db.Column(db.DateTime)  # 停机时间
    start_time = db.Column(db.DateTime)  # 启机时间
    lost_power = db.Column(db.Float)  # 损失电量
    time = db.Column(db.Float)  # 停机时间
    error_code = db.Column(db.String(100))
    error_content = db.Column(db.String(100))
    type = db.Column(db.String(100))
    error_approach = db.Column(db.String(100))
    gzp_id = db.Column(db.String(50), db.ForeignKey('gzp.gzp_id', ondelete="CASCADE"))  # 定义关系工作票与维护单


# 工作票记录
class Gzp(PaginatedAPIMixin, db.Model):
    __tablename__ = 'gzp'
    id = db.Column(db.Integer, primary_key=True)
    # 定义维护单与工作票 一对多关系 外键
    wtms = db.relationship("WTMaintain", backref="gzp", lazy='dynamic', cascade='all, delete-orphan',
                           passive_deletes=True)
    gzp_id = db.Column(db.String(50), unique=True)  # 工作票票号
    firm = db.Column(db.String(100))  # 单位
    pstart_time = db.Column(db.DateTime)  # 计划开始时间
    pstop_time = db.Column(db.DateTime)  # 计划结束时间
    error_code = db.Column(db.String(100))  # 故障代码
    error_content = db.Column(db.String(100))  # 故障内容
    # members = db.relationship("User", secondary=member_gzp, backref="exec_gzps", lazy="dynamic")  # 工作班成员
    task = db.Column(db.String(500))  # 任务
    postion = db.Column(db.String(100), server_default='整机')  # 工作位置
    sign_time = db.Column(db.DateTime)  # 签发时间
    allow1_time = db.Column(db.DateTime)  # 值班许可时间
    end1_time = db.Column(db.DateTime)  # 值班许可终结时间
    allow2_time = db.Column(db.DateTime)  # 现场许可时间
    end2_time = db.Column(db.DateTime)  # 现场许可终结时间

    sign_person_id = db.Column(db.Integer, db.ForeignKey("user.id"))  # 签发的工作票 外键
    sign_person = db.relationship('User', foreign_keys=[sign_person_id])  # 签发的工作票

    manage_person_id = db.Column(db.Integer, db.ForeignKey("user.id"))  # 担任工作负责人的工作票 外键
    manage_person = db.relationship('User', foreign_keys=[manage_person_id])  # 担任负责人的工作票

    allow1_person_id = db.Column(db.Integer, db.ForeignKey("user.id"))  # 值班许可的工作票
    allow1_person = db.relationship('User', foreign_keys=[allow1_person_id])

    allow2_person_id = db.Column(db.Integer, db.ForeignKey("user.id"))  # 现场许可的工作票
    allow2_person = db.relationship('User', foreign_keys=[allow2_person_id])  # 现场许可的工作票

    members = db.relationship("User", secondary=member_gzp, backref=db.backref('gzps', lazy='joined'),
                              lazy="dynamic")  # 工作班成员
    wts = db.relationship("WT", secondary=wt_gzp, backref=db.backref('gzps', lazy='joined'),
                          lazy="dynamic")  # 风机号
    is_end = db.Column(db.Boolean, default=False)  # 是否终结

    @staticmethod
    def to_col_dict(query):
        data = {
            'items': [item.to_dict() for item in query]
        }
        return data


# 限电记录
class PowerCut(PaginatedAPIMixin, db.Model):
    __tablename__ = 'powercut'
    id = db.Column(db.Integer, primary_key=True)
    start_time = db.Column(db.DateTime)  # 限电开始时间
    stop_time = db.Column(db.DateTime)  # 限电结束时间
    time = db.Column(db.Float)  # 限电时间
    lost_power1 = db.Column(db.Float)  # 1期损失电量
    lost_power2 = db.Column(db.Float)  # 2期损失电量
