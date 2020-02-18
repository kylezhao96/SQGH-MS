from flask import Blueprint


bp = Blueprint('api', __name__)

from app.api import errors, tasks, dailyform, wts, users, gzp, pc
