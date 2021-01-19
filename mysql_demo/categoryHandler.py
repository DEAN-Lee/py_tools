import os
import configparser
import sys
import pymysql


def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def readDBConf():
    config = configparser.ConfigParser()
    config.read(resource_path(os.path.join('conf', 'conf.ini')), encoding="utf8")
    return config.get("database", "host"), config.get("database", "port"), config.get("database", "user"), config.get(
        "database", "passwd"), config.get("database", "db")


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

            self.__cursor = self._conn.cursor()
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


if __name__ == '__main__':
    util = MySqlUtil()
