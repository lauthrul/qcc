# by Ganxiaozhe (hi@gxzv.com)
# 2022-11-29
# https://zhuanlan.zhihu.com/p/587896657
# https://gxzv.com/blog/qcc_headers_hash/
#
# win_tid获取：
# https://qcc-static.qichacha.com/qcc/pc-web/prod-23.02.70/common-88d3322f.80048e3e.js
# s (common-88d3322f.80048e3e.js:formatted:11139)
# u (common-88d3322f.80048e3e.js:formatted:10589) 控制台：(0,s.default)()

import json
import hashlib
import hmac


# 在这里填写请求数据
# req_url = '/api/search/searchcount'
# req_data = {
#     'count': True,
#     'filter': "{\"i\":[\"A\"],\"r\":[{\"pr\":\"GD\"},{\"pr\":\"CQ\"}]}"
# }
# win_tid = '8b4ebc1e4a1b8c21235f34bf9db8f1a8'


def seeds_generator(s):
    seeds = {
        "0": "W",
        "1": "l",
        "2": "k",
        "3": "B",
        "4": "Q",
        "5": "g",
        "6": "f",
        "7": "i",
        "8": "i",
        "9": "r",
        "10": "v",
        "11": "6",
        "12": "A",
        "13": "K",
        "14": "N",
        "15": "k",
        "16": "4",
        "17": "L",
        "18": "1",
        "19": "8"
    }
    seeds_n = 20

    if not s:
        s = "/"
    s = s.lower()
    s = s + s

    res = ''
    for i in s:
        res += seeds[str(ord(i) % seeds_n)]
    return res


def a_default(url: str = '/', data: object = {}):
    url = url.lower()
    dataJson = json.dumps(data, ensure_ascii=False, separators=(',', ':')).lower()

    hash = hmac.new(
        bytes(seeds_generator(url), encoding='utf-8'),
        bytes(url + dataJson, encoding='utf-8'),
        hashlib.sha512
    ).hexdigest()
    return hash.lower()[8:28]


def r_default(url: str = '/', data: object = {}, tid: str = ''):
    url = url.lower()
    dataJson = json.dumps(data, ensure_ascii=False, separators=(',', ':')).lower()

    payload = url + 'pathString' + dataJson + tid
    key = seeds_generator(url)

    hash = hmac.new(
        bytes(key, encoding='utf-8'),
        bytes(payload, encoding='utf-8'),
        hashlib.sha512
    ).hexdigest()
    return hash.lower()


if __name__ == '__main__':
    url = '/api/datalist/tenderlist?companyId=5706dde2154629887c658d8c9687973e&pageIndex=2&type=100'
    data = {}
    win_tid = '8c2ee8f227b83e1fe4a450b4b6c63dd1'
    print(url)
    print('%s: %s' % (a_default(url, data), r_default(url, data, win_tid)))
