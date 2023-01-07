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
        'Host': 'www.qcc.com',
        'Connection': 'keep-alive',
        'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Microsoft Edge";v="108"',
        'x-pid': '4d8acbf5774ccb2938b764721b040c71',
        'sec-ch-ua-mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.54',
        'Accept': 'application/json, text/plain, */*',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua-platform': '"Windows"',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://www.qcc.com/crun/5706dde2154629887c658d8c9687973e.html',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Cookie': 'qcc_did=98db4971-f6ef-4f21-bc84-d0459fbfd1b7; '
                  'UM_distinctid=1857cedaa299f8-0285a851cf387a-26021151-fa000-1857cedaa2a10f4; '
                  'CNZZDATA1254842228=1798699866-1672835787-%7C1672835787; QCCSESSID=dc3dd8c794799708a01a4b8e44; '
                  'acw_tc=0e77411216729196579005931e1797a5b160044c77b01921c7c56efe76 '
    }
    # 通过web调试获取得到
    win_tid = 'b46d21917b04bdbf9ea01dbc2bb7cb79'
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
        print('[%.2fs] [#%d] -> %d' % (t, page, resp.status_code))
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
