import xlwt

from TestAnalysis.db.MySqlUtil import MySqlUtil

if __name__ == '__main__':
    dbUtil = MySqlUtil()
    placeIdData = [10101, 12701, 10201, 10801, 10816, 10301, 10401, 10803, 10804, 11501, 11601, 13101, 13106,
                11806,
                   11801, 10608,
                   12601, 11301, 12501, 13001, 11201, 10501, 12001, 12410, 10601, 10901, 11202, 11209, 11802, 11809,
                   12401, 10806,
                   10807, 10821, 10822, 11203, 11207, 11208, 11210, 11808, 11812, 12705, 12711, 12713, 12715, 13103,
                   11813, 10511,
                   10605, 10606, 10812, 10820, 10907, 10911, 11001, 11009, 11515, 11517, 11805, 11901, 12409, 12413,
                   12414, 12415,
                   12417, 13102, 13104, 13109, 13110, 10813, 10802]
    dict = {10101: '北京', 12701: '成都', 10201: '上海', 10801: '广州', 10816: '深圳', 10301: '天津', 10401: '重庆',
            10803: '东莞', 10804: '佛山', 11501: '武汉', 11601: '长沙', 13101: '杭州', 13106: '宁波',
            11806: '苏州', 11801: '南京', 10608: '厦门', 12601: '西安', 11301: '郑州', 12501: '太原', 13001: '昆明',
            11201: '石家庄',
            10501: '合肥', 12001: '沈阳', 12410: '青岛', 10601: '福州', 10901: '南宁', 11202: '保定', 11209: '唐山',
            11802: '常州', 11809: '无锡', 12401: '济南', 10806: '惠州', 10807: '江门', 10821: '中山', 10822: '珠海',
            11203: '沧州', 11207: '廊坊', 11208: '秦皇岛', 11210: '邢台', 11808: '泰州', 11812: '扬州', 12705: '德阳',
            12711: '泸州', 12713: '绵阳', 12715: '南充', 13103: '嘉兴', 11813: '镇江', 10511: '黄山', 10605: '莆田',
            10606: '泉州', 10812: '清远', 10820: '肇庆', 10907: '桂林', 10911: '柳州', 11001: '贵阳', 11009: '遵义',
            11515: '襄阳', 11517: '宜昌', 11805: '南通', 11901: '南昌', 12409: '临沂', 12413: '威海', 12414: '潍坊',
            12415: '烟台', 12417: '淄博', 13102: '湖州', 13104: '金华', 13109: '台州', 13110: '温州', 10813: '汕头',
            10802: '潮州'}

    workbook = xlwt.Workbook(encoding='utf-8')

    for value in placeIdData:
        print(dict[value])
        worksheet = workbook.add_sheet(dict[value])
        dataSql = "SELECT price.place_id as 'city',	MAX(CAST(brand_min_price AS SIGNED)) 'sale',t.*,ly.displacement as 'displacement' FROM `shop_car_model_min_price` as price LEFT JOIN ( SELECT  (SELECT `name` FROM lcb_car_db.car_brand_type n where n.id = m.brandId) as 'brand', m.brandId as '品牌ID', (SELECT `name` FROM lcb_car_db.car_brand_type n where n.id = m.factoryId) as 'fact', m.factoryId as '厂商ID', (SELECT `name` FROM lcb_car_db.car_brand_type n where n.id = m.seriaId) as 'seria', m.seriaId as '车系ID', (SELECT `name` FROM lcb_car_db.car_brand_type n where n.id = m.yearId) as 'year', m.yearId as '年款ID', name as 'name', m.id as brand_id1 FROM (SELECT id,name,path,SUBSTRING_INDEX(path, ',', 1) AS brandId, SUBSTRING_INDEX(SUBSTRING_INDEX(path, ',', 2),',' ,- 1) AS factoryId, SUBSTRING_INDEX(SUBSTRING_INDEX(path, ',', 3),',' ,- 1) AS seriaId, SUBSTRING_INDEX(SUBSTRING_INDEX(path, ',', 4),',' ,- 1) AS yearId FROM lcb_car_db.car_brand_type WHERE `level`= 5 and deleted=0) m ) AS t  on t.brand_id1=price.brand_type_id LEFT JOIN lcb_car_db.car_ly_brand_type as ly on ly.lcb_brand_type_id=price.brand_type_id WHERE 	price.item_id = 1 and price.place_id={} GROUP BY price.brand_type_id ;"
        sql = dataSql.format(value)
        print(sql)
        # dbUtil.get_db()
        dataVale = dbUtil.select_more_sql(sql)
        dbUtil.closeCursor()
        i = 0
        for data in dataVale:
            city = data['city']
            sale = data['sale']
            brand = data['brand']
            fact = data['fact']
            seria = data['seria']
            year = data['year']
            name = data['name']
            brand_id1 = data['brand_id1']
            displacement = data['displacement']
            worksheet.write(i, 0, label=city)
            worksheet.write(i, 1, label=sale)
            worksheet.write(i, 2, label=brand)
            worksheet.write(i, 3, label=fact)
            worksheet.write(i, 4, label=seria)
            worksheet.write(i, 5, label=year)
            worksheet.write(i, 6, label=name)
            worksheet.write(i, 7, label=brand_id1)
            worksheet.write(i, 8, label=displacement)
            i += 1
        # dbUtil.close_db()

        workbook.save('Excel_test.xls')
