from collections import Counter
from pyecharts.charts import Bar,Page
from pyecharts import options as opts

from lottery.db.MySqlUtil import MySqlUtil


class RenderChats:
    def __init__(self):
        super().__init__()

    def render(self, **kwargs):
        """新增一条记录
                  table: 表名
                  data: dict 插入的数据
        """
        fields = ','.join('`' + k + '`' for k in kwargs["data"].keys())
        values = ','.join(("%s",) * len(kwargs["data"]))


if __name__ == '__main__':
    dbUtil = MySqlUtil()
    data = dbUtil.select_more('lottery', 'draw_time>="2007-01-01"')
    red_balls = []
    blue_balls = []
    for value in data:
        red_balls.append(value['pre_num1'])
        red_balls.append(value['pre_num2'])
        red_balls.append(value['pre_num3'])
        red_balls.append(value['pre_num4'])
        red_balls.append(value['pre_num5'])
        blue_balls.append(value['aft_num1'])
        blue_balls.append(value['aft_num2'])

    red_counter = Counter(red_balls)
    blue_counter = Counter(blue_balls)

    red_dict = {}
    blue_dict = {}
    for i in red_counter.most_common():  # 使用collections模块的counter统计红球和蓝球各个号码的出现次数
        red_dict['{}'.format(i[0])] = i[1]
    for j in blue_counter.most_common():
        blue_dict['{}'.format(j[0])] = j[1]

    red_list = sorted(red_counter.most_common(), key=lambda number: number[0])  # 对红蓝球号码和次数重新进行排序
    blue_list = sorted(blue_counter.most_common(), key=lambda number: number[0])

    red_bar = Bar()
    red_x = ['{}'.format(str(x[0])) for x in red_list]
    red_y = ['{}'.format(str(x[1])) for x in red_list]
    red_bar.add_xaxis(red_x)
    red_bar.add_yaxis('红色球出现次数', red_y)
    red_bar.set_global_opts(title_opts=opts.TitleOpts(title='大乐透彩票', subtitle='近14年数据'), toolbox_opts=opts.ToolboxOpts()
                            , yaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(formatter="{value}/次")),
                            xaxis_opts=opts.AxisOpts(name='开奖号码'))
    red_bar.set_series_opts(markpoint_opts=opts.MarkPointOpts(
        data=[opts.MarkPointItem(type_='max', name='最大值'), opts.MarkPointItem(type_='min', name='最小值')]
    ))

    blue_bar = Bar()
    blue_x = ['{}'.format(str(x[0])) for x in blue_list]
    blue_y = ['{}'.format(str(x[1])) for x in blue_list]
    blue_bar.add_xaxis(blue_x)
    blue_bar.add_yaxis('蓝色球出现次数', blue_y, itemstyle_opts=opts.ItemStyleOpts(color='blue'))
    blue_bar.set_global_opts(title_opts=opts.TitleOpts(title='大乐透彩票', subtitle='近14年数据'),
                             toolbox_opts=opts.ToolboxOpts()
                             , yaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(formatter="{value}/次")),
                             xaxis_opts=opts.AxisOpts(name='开奖号码'))
    blue_bar.set_series_opts(markpoint_opts=opts.MarkPointOpts(
        data=[opts.MarkPointItem(type_='max', name='最大值'), opts.MarkPointItem(type_='min', name='最小值')]
    ))

    page = Page(page_title='大乐透历史开奖数据分析', interval=3)
    page.add(red_bar, blue_bar)
    page.render('大乐透历史开奖数据分析.html')
