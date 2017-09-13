import os
from datetime import date
from time import sleep

import pandas as pd
import xlwings as xw

from spider import CricSpider
from luhe_2 import ToExcel, PianQu
from consts import stocks


class PianQu2(PianQu):
    stock_dict = stocks

    def stock(self, plate, pianqu):
        date_range = ('2006年08月', '2017年08月')
        df = self.monitor(date_range, ['供应面积', '成交面积'], area2=(plate, [pianqu])).drop('汇总')
        df['库存'] = 0
        # 计算滚动6个月月均成交面积
        df['去化速度'] = df['成交面积'].rolling(12).mean()

        # 以2017年8月库存向过去推算
        df_ = df[date_range[0]:'2017年08月']
        stk = self.stock_dict[pianqu]
        for index in reversed(df_.index):
            df_.at[index, '库存'] = stk
            # 上期库存(在下一次迭代时赋值给['库存']) = 本期库存 - (本期上市 - 本期成交)
            stk -= df_.at[index, '供应面积'] - df_.at[index, '成交面积']

        # 向未来推算
        # df_ = df['2017年01月':date_range[1]]
        # stk = self.stock_dict[pianqu[0]]
        # for index in df_.index:
        #     # 本期库存 = 上期库存 + (本期上市 - 本期成交)
        #     stk += df_.at[index, '供应面积'] - df_.at[index, '成交面积']
        #     df_.at[index, '库存'] = stk

        # 去化周期 = 存量 / 去化速度速度
        df['去化周期'] = df['库存'] / df['去化速度']
        # 面积换算成万方
        df['库存'] = df['库存'] / 1e4
        # 留两位小数
        df = df[['库存', '去化周期']].round(2)

        # 年度
        df_year = df['2007年01月':'2017年08月'].copy()
        for each in df_year.index:
            if ('12月' not in each) and (each != '2017年08月'):
                df_year = df_year.drop(each)

        # 月度
        df_month = df['2017年01月':'2017年08月']

        return df_year, df_month

    def base_stock(self, plate, pianqu):
        date_range = ('2006年08月', '2017年08月')
        df = self.monitor(date_range, ['供应面积', '成交面积'], area2=(plate, [pianqu]))
        stk = df.at['汇总', '供应面积'] - df.at['汇总', '成交面积']
        return stk

    def structure(self, plate, pianqu):
        """成交结构"""
        # 面积段
        print('>>> 面积段...')
        df1 = self.monitor(self.this_year_monthly, ['成交套数'], area2=(plate, [pianqu]), column='面积段')[1:-1]

        # 单价段
        sleep(1)
        print('>>> 单价段...')
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.label('单价段', 1)
        self.monitor_stat()
        df2 = self.download2df()[1:-1]

        # 总价段
        sleep(1)
        print('>>> 总价段...')
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.label('总价段', 1)
        self.monitor_stat()
        df3 = self.download2df()[1:-1]

        return df1, df2, df3


if __name__ == '__main__':
    c = PianQu2()
    excel = ToExcel()

    # 计算基期库存
    # d = dict()
    # for plate in c.area_dict:
    #     for pianqu in c.area_dict[plate]:
    #         try:
    #             d[pianqu] = c.base_stock(plate, pianqu)
    #         except:
    #             print(f'{pianqu}查询失败')
    # json_dump(d, 'stock_dict')

    # 成交结构
    path = r'E:\city_report\合肥\2017-09-12'
    for plate in c.area_dict:
        print('=' * 20, plate, '=' * 20)
        for pianqu in c.area_dict[plate]:
            print(f'{plate}-{pianqu}')
            book_path = f'{path}/{plate}_{pianqu}.xlsx'
            excel.creat_book(book_path)
            # df1, df2, df3 = c.structure(plate, pianqu)
            # excel.df2sheet(book_path, '面积段', df1)
            # excel.df2sheet(book_path, '单价段', df2)
            # excel.df2sheet(book_path, '总价段', df3)
            df_year, df_month = c.stock(plate, pianqu)
            excel.df2sheet(book_path, '年度库存', df_year)
            excel.df2sheet(book_path, '月度库存', df_month)
