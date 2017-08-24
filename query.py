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
        print('>>> 正在保存...')
        wb = self.app.books.open(book)
        sheets = wb.sheets
        sht = sheets.add(sheet, after=len(sheets))
        sht.range('A1').value = df
        wb.save()
        wb.close()
        print('>>> 保存成功！')


class CricQuery(CricSpider):
    """固化一些查询动作"""
    year_range = ('2013年', '2016年')
    month_range = ('2016年01月', '2017年07月')

    def land_yearly(self, area_tuple):
        """跳转至土地统计页面,设置好年度时间,区域及交易方式"""
        # 土地统计页面
        self.land_page()
        # 时间 年度
        self.land_date_range(*self.year_range)
        # 区域 详细到板块
        self.land_area(area_tuple)
        # 交易方式
        self.land_method(['招拍挂土地'])

    def land_structure(self, area_tuple):
        """各属性用地成交结构
        行为每年,列为每种属性,值为成交幅数
        """
        df = pd.DataFrame()
        self.land_yearly(area_tuple)

        for usage in ['纯住宅', '商住', '商办', '综合用地', '工业', '其他']:
            # 土地性质
            self.land_usg([usage])
            # 统计
            self.land_stat()
            # to Dataframe
            df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
            df[usage] = df_['幅数']

        return df.drop('汇总')

    def land_volume_price(self, area_tuple):
        """成交量(幅数)与楼面价特征
        行为每年, 列为量\价
        """
        self.land_yearly(area_tuple)
        # 土地性质
        self.land_usg(['纯住宅', '商住'])
        # 统计
        self.land_stat()
        # to Dataframe
        df = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
        return df[['幅数', '成交楼板价']].drop('汇总')

    def gxj(self, area):
        """供销价，返回一个年度走势df与一个月度走势df"""
        def ajust_df(df):
            # 面积换算成万方
            df.iloc[:, [0, 1]] = round(df.iloc[:, [0, 1]] / 1e4, 2)
            # 均价四舍五入为整数
            df.iloc[:, 2] = round(df.iloc[:, 2]).astype('int')
            return df.drop('汇总')

        # 年度
        self.monitor_page('合肥')
        sleep(1)
        self.date_range(*self.year_range)
        self.area(area)
        self.room_usg(['普通住宅'])
        self.monitor_stat()
        self.monitor_output(['供应面积', '成交面积', '成交均价'])
        df_year = ajust_df(self.download2df())

        # 月度 显示条件后,改变统计时间再统计一次
        sleep(1)
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.date_range(*self.month_range)
        self.monitor_stat()
        df_month = ajust_df(self.download2df())

        return df_year, df_month

    def download2df(self):
        download_path = 'C:/Users/dell/Downloads/CRIC2016.xls'
        # 下载文件
        self.click('导出表格')
        # 读入df
        df = pd.read_excel(download_path, index_col=0, header=0)
        # 删除文件
        os.remove(download_path)

        return df

if __name__ == '__main__':
    excel = ToExcel()
    c = CricQuery()
    area_list = ['蜀山区']
    today = date.today()

    # 遍历每个版块
    for area in area_list:
        print('=' * 20, area, '=' * 20)
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
        df_year, df_month = c.gxj(area)
        excel.df2sheet(book_path, '年度供销价', df_year)
        excel.df2sheet(book_path, '月度供销价', df_month)