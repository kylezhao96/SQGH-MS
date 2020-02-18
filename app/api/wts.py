from flask import jsonify

from app.api import bp
from app.models import WT


@bp.route('/getwts', methods=['GET'])
def get_wts():
    """
    return a list stand for wind turbine code
    """
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


