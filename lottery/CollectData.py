import requests
import json

from lottery.db.MySqlUtil import MySqlUtil


class Crawl():
    def __init__(self):
        super().__init__()

    def getData(self, page):
        url = 'https://webapi.sporttery.cn/gateway/lottery/getHistoryPageListV1.qry?gameNo=85&provinceId=0&pageSize=1&isVerify=1&pageNo={}'.format(
            page)
        headers = {
            'Host': 'webapi.sporttery.cn',
            'Connection': 'keep-alive',
            'Pragma': 'no-cache',
            'Cache-Control': 'no-cache',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36',
            'Origin': 'https://static.sporttery.cn',
            'Sec-Fetch-Site': 'same-site',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Dest': 'empty',
            'Referer': 'https://static.sporttery.cn/',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7'

        }
        print(url)
        r = requests.get(url, headers=headers)
        print(r.status_code)
        r.encoding = 'utf-8'
        if r.status_code == 200:
            return r.text
        else:
            return -1


if __name__ == '__main__':
    crawl = Crawl()
    util = MySqlUtil()
    for i in range(1):
        data = crawl.getData(i + 1)
        if data != -1:
            dataJson = json.loads(data)
            dataValue = dataJson['value']['list']
            i = 0
            while (i < len(dataValue)):
                value = dataValue[i]
                record = {}
                record['draw_num'] = value['lotteryDrawNum']
                record['draw_time'] = value['lotteryDrawTime']
                record['draw_result'] = value['lotteryDrawResult']
                numStr = value['lotteryDrawResult']
                numArray = numStr.split()
                record['pre_num1'] = numArray[0]
                record['pre_num2'] = numArray[1]
                record['pre_num3'] = numArray[2]
                record['pre_num4'] = numArray[3]
                record['pre_num5'] = numArray[4]
                record['aft_num1'] = numArray[5]
                record['aft_num2'] = numArray[6]
                record['unsort_draw_result'] = value['lotteryUnsortDrawresult']
                record['prize_info'] = json.dumps(value['prizeLevelList'], ensure_ascii=False)
                record['created_time'] = util.current_time()
                util.insert(table='lottery', data=record)
                i += 1

    # util.close_db()
