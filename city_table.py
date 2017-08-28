
from datetime import date
import pandas as pd
from query import CricQuery, ToExcel

if __name__ == '__main__':
    def land(date_range):
        c.land(area_tuple, date_range)
        c.land_usg(['纯住宅', '商住'])
        c.land_stat()
        return pd.read_html(c.driver.page_source, index_col=0, header=0)[0]


    area_list = ['蜀山区', '庐阳区', '政务区', '高新区', '瑶海区', '滨湖区', '包河区', '经济区', '新站区', '肥东县', '肥西县', '长丰县']
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
    range_2017 = ('2017年01月', '2017年08月')
    range_years = ('2013年', '2016年')
    df_city = pd.DataFrame()
    c = CricQuery()

    # creat excel book
    book_path = f'e:/city_report/合肥/{today}/大表.xlsx'
    excel = ToExcel()
    excel.creat_book(book_path)

    for area in area_list:
        print(area)
        # 变量
        c.stock_base = stock[area]
        area_tuple = ('安徽省', '合肥', area)

        # 2017年含住宅土地成交楼面价
        df = land(range_2017)
        df_city.at[area, '2017年含住宅类成交楼面价'] = df.at['汇总', '成交楼板价']

        # 年均土地成交量
        df = land(range_years)
        df_city.at[area, '年均土地成交量'] = round(df.at['汇总', '建筑面积'] / 1e4 / 4, 2)

        # 存量水平
        df_city.at[area, '存量水平'] = c.stock(area)[0].iat[-1, 0]

    # 2017住宅均价
    df = c.monitor(range_2017, ['成交均价'], index='地域')
    df.columns = ['2017年商品住宅成交均价']
    df_city = pd.concat([df_city, df], axis=1, join='inner')
    df_city['2017年商品住宅成交均价'] = df_city['2017年商品住宅成交均价'].round().astype('int')

    # 住宅年均去化
    df = c.monitor(range_years, ['成交面积'], index='地域')
    df = round(df / 1e4 / 4, 2)
    df.columns = ['住宅年均去化量']
    df_city = pd.concat([df_city, df], axis=1, join='inner')

    c.driver.close()

    # 计算列
    df_city['地楼比'] = df_city['2017年含住宅类成交楼面价'] / df_city['2017年商品住宅成交均价']
    df_city['土地成交/住宅成交'] = df_city['年均土地成交量'] / df_city['住宅年均去化量']

    # 调整保存
    df_city['2017年含住宅类成交楼面价'] = df_city['2017年含住宅类成交楼面价'].astype('int')
    df_city[['年均土地成交量', '存量水平', '住宅年均去化量', '地楼比', '土地成交/住宅成交']] = df_city[
        ['年均土地成交量', '存量水平', '住宅年均去化量', '地楼比', '土地成交/住宅成交']].round(2)

    df_city = df_city[[
        '2017年含住宅类成交楼面价', '2017年商品住宅成交均价', '地楼比',
        '年均土地成交量', '住宅年均去化量', '土地成交/住宅成交',
        '存量水平'
    ]]
    excel.df2sheet(book_path, '大表', df_city)
