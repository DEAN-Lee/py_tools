import json
import os
import configparser
import sys
import pymysql
import time


def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def readDBConf():
    config = configparser.ConfigParser()
    config.read(resource_path(os.path.join('conf', 'conf.ini')), encoding="utf8")
    return config.get("database", "host"), config.get("database", "user"), config.get(
        "database", "passwd"), config.get(
        "database", "charset"), config.get("database", "db"), config.getint("database", "port")


class MySqlUtil:
    def __init__(self):
        try:
            config = readDBConf()
            self._conn = pymysql.connect(host=config[0],
                                         user=config[1],
                                         password=config[2],
                                         charset=config[3],
                                         database=config[4],
                                         port=config[5],
                                         cursorclass=pymysql.cursors.DictCursor)

            self.__cursor = None
            print("连接数据库")
            # set charset charset = ('latin1','latin1_general_ci')
        except Exception as err:
            print('mysql连接错误：' + err.msg)

    def close_db(self):
        self.__cursor.close()
        self._conn.close()

    def insert(self, **kwargs):
        """新增一条记录
          table: 表名
          data: dict 插入的数据
        """
        fields = ','.join('`' + k + '`' for k in kwargs["data"].keys())
        values = ','.join(("%s",) * len(kwargs["data"]))
        sql = 'INSERT INTO `%s` (%s) VALUES (%s)' % (kwargs["table"], fields, values)
        cursor = self.__getCursor()
        cursor.execute(sql, tuple(kwargs["data"].values()))
        insert_id = cursor.lastrowid
        self._conn.commit()
        return insert_id

    def exist(self, **kwargs):
        """判断是否存在"""
        return self.count(**kwargs) > 0

    def close(self):
        """关闭游标和数据库连接"""
        if self.__cursor is not None:
            self.__cursor.close()
        self._conn.close()

    def __getCursor(self):
        """获取游标"""
        if self.__cursor is None:
            self.__cursor = self._conn.cursor()
        return self.__cursor


def current_time():
    # 毫秒级时间戳
    t = time.time()
    return int(round(t * 1000))


if __name__ == '__main__':
    util = MySqlUtil()
    with open(os.path.join('conf', 'data.json'), encoding='utf-8') as f:
        data = json.load(f)
        i = 0
        while i < len(data):
            print("level1 JSON 对象：", data[i])
            level1 = data[i]
            opData = level1.copy()
            if 'sublist' in level1:
                del opData['sublist']
            opData['created_time'] = current_time()
            opData['updated_time'] = current_time()
            opData['is_leaf'] = 0
            level1Id = util.insert(table='pro_category', data=opData)
            if 'sublist' in level1:
                level2list = level1['sublist']
                j = 0
                while j < len(level2list):
                    level2 = level2list[j]
                    print("level2 对象：", level2)
                    opData2 = level2.copy()
                    if 'sublist' in level2:
                        del opData2['sublist']
                    opData2['created_time'] = current_time()
                    opData2['updated_time'] = current_time()
                    opData2['is_leaf'] = 0
                    opData2['parent_id'] = level1Id
                    opData2['path'] = level1Id
                    level2Id = util.insert(table='pro_category', data=opData2)
                    if 'sublist' in level2:
                        level3List = level2['sublist']
                        m = 0
                        while m < len(level3List):
                            level3 = level3List[m]
                            print("level3 对象：", level3)
                            opData3 = level3.copy()
                            if 'sublist' in level3:
                                del opData3['sublist']
                            opData3['created_time'] = current_time()
                            opData3['updated_time'] = current_time()
                            opData3['is_leaf'] = 1
                            opData3['parent_id'] = level2Id
                            seq = (repr(level1Id), repr(level2Id))
                            opData3['path'] = ','.join(seq)
                            level3Id = util.insert(table='pro_category', data=opData3)

                            m += 1
                    j += 1

            i += 1

    print(repr(current_time()))
