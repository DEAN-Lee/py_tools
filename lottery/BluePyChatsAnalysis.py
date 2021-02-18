from collections import Counter
from pyecharts.charts import Bar, Page, Line
from pyecharts import options as opts
from pyecharts.options import MinorTickOpts

from lottery.db.MySqlUtil import MySqlUtil

"""篮球组合分析"""


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
    data = dbUtil.select_more('lottery', 'draw_time>="2021-01-01"')
    blue_balls = []
    blue_num = []
    for value in data:
        tag = int(str(value['aft_num1']) + str(value['aft_num2']))
        blue_balls.append(tag)
        blue_num.append(value['draw_num'])

    blue_counter = Counter(blue_balls)

    blue_dict = {}
    # 使用collections模块的counter统计红球和蓝球各个号码的出现次数
    for j in blue_counter.most_common():
        blue_dict['{}'.format(j[0])] = j[1]

    # 对红蓝球号码和次数重新进行排序
    blue_list = sorted(blue_counter.most_common(), key=lambda number: number[0])

    blue_bar = Bar(init_opts=opts.InitOpts(width="1800px", height="800px"))
    blue_x = ['{}'.format(str(x[0])) for x in blue_list]
    blue_y = ['{}'.format(str(x[1])) for x in blue_list]
    blue_bar.add_xaxis(blue_x)
    blue_bar.add_yaxis('蓝色球出现次数', blue_y, itemstyle_opts=opts.ItemStyleOpts(color='blue'), category_gap='30%')

    blue_bar.set_global_opts(title_opts=opts.TitleOpts(title='大乐透彩票', subtitle='开奖至今数据'),
                             # toolbox_opts=opts.ToolboxOpts(),
                             # datazoom_opts=opts.DataZoomOpts(),
                             yaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(formatter="{value}/次")),
                             xaxis_opts=opts.AxisOpts(name='开奖号码', axislabel_opts=opts.LabelOpts(rotate=-90)))
    blue_bar.set_series_opts(markpoint_opts=opts.MarkPointOpts(
        data=[opts.MarkPointItem(type_='max', name='最大值'), opts.MarkPointItem(type_='min', name='最小值')]
    ))

    blue_line = Line(init_opts=opts.InitOpts(width="1800px", height="800px"))
    blue_line.set_global_opts(title_opts=opts.TitleOpts(title="Line-基本示例"))
    blue_line.add_xaxis(blue_num)
    blue_line.add_yaxis('开奖结果', blue_balls)

    page = Page(page_title='大乐透历史开奖数据分析', interval=3)
    page.add(blue_bar, blue_line)
    page.render('大乐透-篮球——历史开奖数据分析.html')
