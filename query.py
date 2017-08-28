from spider import CricSpider
import pandas as pd
import xlwings as xw
from datetime import date
import os
from time import sleep


class ToExcel:
    """将df存入Excel"""

    def __init__(self):
        self.app = xw.App(visible=False, add_book=False)

    def creat_book(self, path):
        """创建一个新的excel文件"""
        wb = self.app.books.add()
        wb.save(path)
        wb.close()

    def df2sheet(self, book, sheet, df):
        """在指定excel文件下创建一个新的sheet,并写上df保存"""
        wb = self.app.books.open(book)
        sheets = wb.sheets
        sht = sheets.add(sheet, after=len(sheets))
        sht.range('A1').value = df
        wb.save()
        wb.close()
        print('>>> 保存成功！')


class CricQuery(CricSpider):
    """固化一些查询动作"""
    year_range = ('2013年', '2017年')
    month_range = ('2016年01月', '2017年07月')
    xmonth_range = ('2013年01月', '2017年07月')
    # 基准库存, 其他库存按照这个值,结合供销量,向前向后推算
    stock_base = 0

    def download2df(self):
        download_path = 'C:/Users/dell/Downloads/CRIC2016.xls'
        # 下载文件
        self.click('导出表格')
        sleep(1)
        # 读入df
        df = pd.read_excel(download_path, index_col=0, header=0)
        # 删除文件
        sleep(1)
        os.remove(download_path)

        return df

    #################################################
    # 土地
    #################################################

    def land_structure(self, area_tuple):
        """各属性用地成交结构
        行为每年,列为每种属性,值为成交幅数
        """
        df = pd.DataFrame()
        self.land(area_tuple, self.year_range)

        for usage in ['纯住宅', '商住', '商办', '综合用地']:
            print(f'>>> {usage}...')
            # 土地性质
            self.land_usg([usage])
            # 统计
            self.land_stat()
            # to Dataframe
            df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
            df[usage] = df_['幅数']

        # 将商住合并进综合用地
        df['综合用地'] = df['综合用地'] + df['商住']
        return df.drop('汇总').drop('商住', axis=1)

    def land_volume_price(self, area_tuple):
        """成交量(幅数)与楼面价特征
        行为每年, 列为量\价
        """
        self.land(area_tuple, self.year_range)
        # 土地性质
        self.land_usg(['纯住宅', '商住'])
        # 统计
        self.land_stat()
        # to Dataframe
        df = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
        return df[['幅数', '成交楼板价']].drop('汇总')

    #################################################
    # 市场走势
    #################################################

    def gxj(self, area):
        """供销价，返回一个年度走势df与一个月度走势df"""

        def ajust_df(df):
            # 面积换算成万方
            df.iloc[:, [0, 1]] = round(df.iloc[:, [0, 1]] / 1e4, 2)
            # 均价四舍五入为整数
            df.iloc[:, 2] = round(df.iloc[:, 2]).astype('int')
            return df.drop('汇总')

        # 年度
        print('>>> 年度...')
        df = self.monitor(self.year_range, ['供应面积', '成交面积', '成交均价'], area)
        df_year = ajust_df(df)

        # 月度 显示条件后,改变统计时间再统计一次
        sleep(1.5)
        print('>>> 月度...')
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.date_range(*self.month_range)
        self.monitor_stat()
        df_month = ajust_df(self.download2df())

        # 分片区
        print('>>> 月度...')
        df_plate = self.plate_gxj(area)

        return df_year, df_month, df_plate

    def stock(self, area):
        """库存及去化周期"""

        def calculate(df):
            df = df.drop('汇总')
            df['库存'] = 0

            # 计算滚动6个月月均成交面积
            df['去化速度'] = df['成交面积'].rolling(6).mean()

            # 计算滚动1年成交面积
            # df['去化速度'] = df['成交面积'].rolling(12).sum()

            # 向过去推算
            # df_ = df[self.xmonth_range[0]:'2016年12月']
            df_ = df
            stk = self.stock_base
            for index in reversed(df_.index):
                df_.at[index, '库存'] = stk
                # 上期库存(在下一次迭代时赋值给['库存']) = 本期库存 - (本期上市 - 本期成交)
                stk -= df_.at[index, '供应面积'] - df_.at[index, '成交面积']

            # 向未来推算
            # df_ = df['2017年01月':self.xmonth_range[1]]
            # stk = self.stock2017
            # for index in df_.index:
            #     # 本期库存 = 上期库存 + (本期上市 - 本期成交)
            #     stk += df_.at[index, '供应面积'] - df_.at[index, '成交面积']
            #     df_.at[index, '库存'] = stk

            # 去化周期 = 存量 / 去化速度速度
            df['去化周期'] = df['库存'] / df['去化速度']
            return df

        def ajust_df(df):
            # 除了17年保留月度数据,其他年份只保留年末数据
            for index in df.index:
                if '12月' not in index and '2017年' not in index:
                    df = df.drop(index)

            # 只保留库存和去化周期两列
            df = df[['库存', '去化周期']]

            # 面积换算成万方, 库存\去化周期保留两位小数
            df['库存'] = df['库存'] / 1e4
            return df.round(2)

        df = self.monitor(self.xmonth_range, ['供应面积', '成交面积'], area)
        df_ = calculate(df)
        return ajust_df(df_), df_

    def structure(self, area):
        """成交结构"""
        # 面积段
        print('>>> 面积段...')
        df1 = self.monitor(self.xmonth_range, ['成交套数'], area, column='面积段')[1:-1]

        # 单价段
        sleep(1.5)
        print('>>> 单价段...')
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.label('单价段', 1)
        self.monitor_stat()
        df2 = self.download2df()[1:-1]

        # 总价段
        sleep(1.5)
        print('>>> 总价段...')
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.label('总价段', 1)
        self.monitor_stat()
        df3 = self.download2df()[1:-1]

        return df1, df2, df3

    #################################################
    # 市场分片区
    #################################################

    def plate_gxj(self, area):
        return self.monitor(self.month_range, ['供应面积', '成交面积', '成交均价'], area2=area, index='地域').drop('汇总')


if __name__ == '__main__':
    excel = ToExcel()
    c = CricQuery()
    area_dict = {
        '庐阳区': ['四里河', '双岗板块', '长江中路板块', '城北板块'],
        '蜀山区': ['黄潜望板块', '蜀山东板块', '凤凰城板块', '小庙片区', '十里庙片区', '三里庵板块', '蜀山产业园'],
        '包河区': ['高铁南站板块', '马鞍山路板块', '葛大店板块', '南门换乘中心', '包河工业园板块', '南七板块', ],
        '瑶海区': ['东门换乘中心', '裕溪路板块', '长江东路板块', '龙岗片区'],
        '政务区': ['政务区板块'],
        '滨湖区': ['滨湖启动区', '金融后台中心板块', '文旅中心板块', '省府板块'],
        '经济区': ['翡翠湖板块', '明珠广场板块', '南艳湖板块', '政务南板块', '桃花工业园'],
        '高新区': ['王咀湖板块', '大铺头板块', '西门换乘中心', '柏堰湖板块', '空港产业园'],
        '新站区': ['新海公园板块', '职教城板块', '火车站板块', '新蚌埠路板块', '陶冲湖板块'],
        '肥东县': ['龙岗片区', '店埠片区', '撮镇片区'],
        '肥西县': ['上派片区', '柏堰科技园板块', '紫蓬山片区'],
        '长丰县': ['岗集片区', '北城片区', '双凤工业园片区']
    }

    stock = {
        '新站区': 386400,
        '高新区': 322000,
        '经济区': 269200,
        '包河区': 221000,
        '瑶海区': 209100,
        '政务区': 140000,
        '蜀山区': 117400,
        '庐阳区': 276600,
        '滨湖区': 311200,
        '长丰县': 2297526,
        '肥西县': 768596,
        '肥东县': 0,
    }
    today = date.today()

    # 遍历每个版块
    for area in area_dict:
        print('=' * 20, area, '=' * 20)
        # 库存基数
        c.stock_base = stock[area]
        # 区域tuple
        area_tuple = ('安徽省', '合肥', area)
        # 创建一个excel文件
        book_path = f'e:/city_report/合肥/{today}/{area}.xlsx'
        excel.creat_book(book_path)

        # 土地结构
        print(f'>>> 正在查询土地结构...')
        df = c.land_structure(area_tuple)
        excel.df2sheet(book_path, '土地成交结构', df)

        # 土地量价
        print(f'>>> 正在查询土地成交量价特征...')
        df = c.land_volume_price(area_tuple)
        excel.df2sheet(book_path, '土地量价特征', df)

        # 住宅供销价走势
        print(f'>>> 正在查询住宅供销价走势...')
        df_year, df_month, df_plate = c.gxj(area)
        excel.df2sheet(book_path, '年度供销价', df_year)
        excel.df2sheet(book_path, '月度供销价', df_month)
        excel.df2sheet(book_path, '片区供销价', df_plate)

        # 库存及去化周期
        print(f'>>> 正在计算库存及去化周期...')
        df, df_ = c.stock(area)
        excel.df2sheet(book_path, '库存及去化', df)
        # excel.df2sheet(book_path, '库存及去化备查', df_)

        # 成交结构
        print(f'>>> 正在查询成交结构数据...')
        df1, df2, df3 = c.structure(area)
        excel.df2sheet(book_path, '面积段', df1)
        excel.df2sheet(book_path, '单价段', df2)
        excel.df2sheet(book_path, '总价段', df3)
