import os
from datetime import date
from time import sleep

import pandas as pd
import xlwings as xw

from spider import CricSpider
from consts import areas


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
        print(f'>>> 【{sheet}】保存成功！')


class PianQu(CricSpider):
    this_year = ('2017年', '2017年')
    this_year_monthly = ('2017年01月', '2017年08月')
    short_yearly = ('2013年', '2017年')
    x_range = ('2007年', '2017年')
    # 基准库存, 其他库存按照这个值,结合供销量,向前向后推算
    # stock_dict = json_load('stock_dict')

    area_dict = areas

    def download2df(self):
        while True:
            try:
                download_path = 'C:/Users/dell/Downloads/CRIC2016.xls'
                # 下载文件
                self.click('导出表格')
                sleep(0.5)
                # 读入df
                df = pd.read_excel(download_path, index_col=0, header=0)
                # 删除文件
                sleep(0.5)
                os.remove(download_path)
                # 跳出循环
                break
            except:
                # 出错就重新执行
                print('>>> retrying...')
                continue

        return df

    def land_structure(self, area_tuple):
        df = pd.DataFrame()
        self.land(area_tuple, self.short_yearly)

        for usage in ['纯住宅', '商住', '商办', '综合用地', '其他']:
            print(f'>>> {usage}...')
            # 土地性质
            self.land_usg([usage])
            # 统计
            self.land_stat()
            # to Dataframe
            df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
            df[usage] = df_['幅数']

        # 合并商办、综合、其他
        df['商办及其他'] = df['综合用地'] + df['商办'] + df['其他']
        return df[['纯住宅', '商住', '商办及其他']].drop('汇总')

    def land_house(self, area_tuple):
        df = pd.DataFrame()

        # 土地可建
        self.land(area_tuple, self.short_yearly)
        self.land_stat()

        df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
        df['可建面'] = round(df_['建筑面积'] / 1e4, 2)
        df['楼面价'] = df_['成交楼板价']

        # 住宅成交
        plate = area_tuple[2]
        pianqu = self.area_dict[plate]
        df_ = self.monitor(self.short_yearly, ['成交面积', '成交均价'], area2=(plate, pianqu))
        df['住宅面积'] = round(df_['成交面积'] / 1e4, 2)
        df['住宅均价'] = df_['成交均价'].round().astype('int')

        df['地价房价比'] = round(df['楼面价'] / df['住宅均价'], 2)
        df = df.drop('汇总')

        df1 = df[['可建面', '住宅面积']]
        df2 = df[['楼面价', '地价房价比']]

        return df1, df2

    def gxj(self, area_tuple, date_range):
        def ajust_df(df):
            # 面积换算成万方
            df.iloc[:, [0, 1]] = round(df.iloc[:, [0, 1]] / 1e4, 2)
            # 均价四舍五入为整数
            df.iloc[:, 2] = round(df.iloc[:, 2]).astype('int')
            # 丢弃汇总行
            return df.drop('汇总')

        df = self.monitor(date_range, ['供应面积', '成交面积', '成交均价'], area2=area_tuple)

        return ajust_df(df)

    def stock(self, area_tuple):
        plate = area_tuple[0]
        pianqu = area_tuple[1]
        date_range = ('2006年08月', '2017年08月')
        df = self.monitor(date_range, ['供应面积', '成交面积'], area2=(plate, pianqu)).drop('汇总')
        df['库存'] = 0
        # 计算滚动6个月月均成交面积
        df['去化速度'] = df['成交面积'].rolling(6).mean()

        # 以2016年12月库存向过去推算
        df_ = df[date_range[0]:'2016年12月']
        stk = self.stock_dict[pianqu[0]]
        for index in reversed(df_.index):
            df_.at[index, '库存'] = stk
            # 上期库存(在下一次迭代时赋值给['库存']) = 本期库存 - (本期上市 - 本期成交)
            stk -= df_.at[index, '供应面积'] - df_.at[index, '成交面积']

        # 向未来推算
        df_ = df['2017年01月':date_range[1]]
        stk = self.stock_dict[pianqu[0]]
        for index in df_.index:
            # 本期库存 = 上期库存 + (本期上市 - 本期成交)
            stk += df_.at[index, '供应面积'] - df_.at[index, '成交面积']
            df_.at[index, '库存'] = stk

        # 去化周期 = 存量 / 去化速度速度
        df['去化周期'] = df['库存'] / df['去化速度']
        # 面积换算成万方
        df['库存'] = df['库存'] / 1e4
        # 留两位小数
        df = df[['库存', '去化周期']].round(2)

        # 年度
        df_year = df['2007年01月':'2016年12月']
        for each in df_year.index:
            if '12月' not in each:
                df_year = df_year.drop(each)

        # 月度
        df_month = df['2017年01月':'2017年08月']

        return df_year, df_month


if __name__ == '__main__':
    today = date.today()
    excel = ToExcel()
    c = PianQu()
    d = dict()
    date_range = ('2006年08月', '2016年12月')

    for plate in c.area_dict:
        print('=' * 20, plate, '=' * 20)
        area_tuple = ('安徽省', '合肥', plate)

        # 创建exced文件
        book_path = f'e:/city_report/合肥/{today}/{plate}土地.xlsx'
        excel.creat_book(book_path)

        # 土地结构
        print(f'>>> {plate}土地结构')
        df = c.land_structure(area_tuple)
        excel.df2sheet(book_path, '土地成交结构', df)

        # 土地和住宅
        print(f'>>> {plate}land&house')
        df1, df2 = c.land_house(area_tuple)
        excel.df2sheet(book_path, '土地住宅面积', df1)
        excel.df2sheet(book_path, '土地价格', df2)

        for pianqu in c.area_dict[plate]:
            area_tuple = (plate, [pianqu])
            # 创建exced文件
            book_path = f'e:/city_report/合肥/{today}/{plate}_{pianqu}.xlsx'
            excel.creat_book(book_path)

            # 年度供销价
            print(f'>>> {plate}-{pianqu}年度供销价')
            df = c.gxj(area_tuple, c.x_range)
            excel.df2sheet(book_path, '年度供销价', df)

            # 月度供销价
            print(f'>>> {plate}-{pianqu}月度供销价')
            df = c.gxj(area_tuple, c.this_year_monthly)
            excel.df2sheet(book_path, '月度供销价', df)

            # 库存
            print(f'>>> {plate}-{pianqu}库存')
            df_year, df_month = c.stock(area_tuple)
            excel.df2sheet(book_path, '年度库存', df_year)
            excel.df2sheet(book_path, '月度库存', df_month)
