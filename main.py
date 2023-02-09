import json
import random
import sys
import requests
from excel import append_excel, open_excel, remove_duplicates
from qcc import a_default, r_default

# 通用参数
company = '北京嘉和美康信息技术有限公司'
path = company + "-投标信息.xlsx"
sheet_name = 'Sheet1'


# 开始爬取数据
# https://www.qcc.com/crun/5706dde2154629887c658d8c9687973e.html
def run(from_page: int, to_page: int):
    # 打开excel
    try:
        wb = open_excel(path, sheet_name)
    except:
        print('文件"{}"已在其他程序中打开，请关闭后再重试！'.format(path))
        exit(0)

    # 写入标题
    if wb.active.max_row <= 1:
        titles = [['页码', '时间', '内容', '类型', '中标金额', '省份', '城市', '招标主体-招标人', '采购方式', '备注', '竞争对手', '中标单位', '是否流标']]
        append_excel(wb, sheet_name, titles)
        wb.save(path)

    # 请求头
    headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'zh-CN,zh;q=0.9',
        'cookie': 'QCCSESSID=6e384441a07dbe8ffa6fd74b71; qcc_did=a6f9f54a-b898-4a21-b2a8-819a35a84e04; '
                  'UM_distinctid=18635f0d2986da-0e004d75d04a-26021051-186a00-18635f0d299d21; '
                  'CNZZDATA1254842228=1387962885-1675940614-%7C1675940614; '
                  'acw_tc=77939c9e16759440568586977e7e204fb2d6c1a156104bf31c86de5a6a',
        'referer': 'https://www.qcc.com/crun/5706dde2154629887c658d8c9687973e.html',
        'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/109.0.0.0 Safari/537.36',
        'x-pid': '259aa12cdd22c15033b54a1f6c0ec28f',
        'x-requested-with': 'XMLHttpRequest'
    }
    # 通过web调试获取得到
    win_tid = '8c2ee8f227b83e1fe4a450b4b6c63dd1'
    # 请求地址
    host = 'https://www.qcc.com'

    # 批量获取页面数据
    row = 0
    for page in range(from_page, to_page + 1):
        # 防止服务器防爬虫，随机睡眠一定时间
        t = random.random()
        # time.sleep(t)

        url = '/api/datalist/tenderlist?companyId=5706dde2154629887c658d8c9687973e&pageIndex={}' \
              '&pageSize=50&type=100'.format(page)
        data = {}
        # 请求头中的{hash_key: hash_value}为关系信息，如果没有或者不对，服务器会返回错误
        hash_key = a_default(url, data)
        hash_value = r_default(url, data, win_tid)
        headers[hash_key] = hash_value
        # 获取数据
        resp = requests.get(host + url, headers=headers)
        print('[%.2fs] [#%d] -> %d' % (t, page, resp.status_code), resp.text if resp.status_code != 200 else "")
        if resp.status_code == 200:
            data = json.loads(resp.text)
            if data is None:
                print('no data')
                continue

            # 内容
            rows = []
            for item in data['data']:
                row += 1
                cols = {
                    '页码': page,
                    '时间': item.get('publishdate'),
                    '内容': item.get('title'),
                    '类型': '',
                    '中标金额': 0,
                    '省份': '',
                    '城市': '',
                    '招标主体-招标人': item.get('ifbunit'),
                    '采购方式': '',
                    '备注': '{0}/tenderDetail/{1}.html'.format(host, item.get('id')),
                    '竞争对手': '',
                    '中标单位': item.get('wtbunit'),
                    '是否流标': '',
                }

                v = item.get('wtbamttotales')
                if v is not None:
                    cols['中标金额'] = float('0' + v)

                v = item.get('arealabels')
                if v is not None:
                    if len(v) >= 1:
                        cols['省份'] = v[0]
                    if len(v) >= 2:
                        cols['城市'] = v[1]

                if str(item.get('title')).find('病历') >= 0:
                    v = '电子病历'
                elif str(item.get('title')).find('改造') >= 0:
                    v = '系统改造'
                else:
                    v = '系统采购'
                cols['类型'] = v

                if str(item.get('wtbunit')).find(company) >= 0:
                    v = '中标'
                else:
                    v = '流标'
                cols['是否流标'] = v

                rows.append(list(cols.values()))

            append_excel(wb, sheet_name, rows)
            wb.save(path)


def usge():
    print('使用方法:')
    print('python {} [run|clean] [<from_page> <to_page>]'.format(sys.argv[0]))
    print('     run <from_page> <to_page>   -   开启爬取数据')
    print('         <from_page>             -   开启的页码，如：1')
    print('         <to_page>               -   结束的页码，如：20')
    print('     clean                       -   数据去重')


if __name__ == '__main__':
    argc = len(sys.argv)
    if argc < 2:
        usge()
        exit(0)

    cmd = sys.argv[1]
    print('开始执行动作: ' + cmd)
    if cmd == 'run':
        if argc < 4:
            usge()
        else:
            from_page = int(sys.argv[2])
            to_page = int(sys.argv[3])
            run(from_page, to_page)
    elif cmd == 'clean':
        remove_duplicates(path, sheet_name, [4, 5, 6, 7])
        print('执行完成!')
