import pymysql
import time

from lottery.conf import common_data


class MySqlUtil:
    def __init__(self):
        try:
            config = common_data.readDBConf()
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

    # 查询多条数据在数据表中
    def select_more(self, table, range_str, field='*'):
        sql = 'SELECT ' + field + ' FROM ' + table + ' WHERE ' + range_str
        try:
            with self.__getCursor() as cursor:
                cursor.execute(sql)
            self._conn.commit()
            return cursor.fetchall()
        except pymysql.Error as e:
            return False

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

    def current_time(self):
        # 毫秒级时间戳
        t = time.time()
        return int(round(t * 1000))
