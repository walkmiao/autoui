from collections import namedtuple

BaoYangQuery = namedtuple('BaoYangQuery', ['date', 'order'])
BaoYangQuery.__new__.__defaults__ = ([], '')
BaoYangEdit = namedtuple('BaoYangEdit',
                         ['mark', 'chepai', 'bian_su_xiang', 'baoxiu_order', 'baoyang_date', 'kehu_name', 'kehu_addr',
                          'youbian', 'phone', 'mobile', 'songbao_name', 'songbao_phone'])
BaoYangEdit.__new__.__defaults__ = ('', '', '', '', '', '', '', '', '', '', '', '')
