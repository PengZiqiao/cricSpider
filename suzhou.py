import pandas as pd
from cric import Query
from consts import areas, stocks


class CricQuery(Query):
    def land_structure(self, area_tuple, date_range):
        df = pd.DataFrame()
        self.land(area_tuple, date_range)

        for usage in ['纯住宅', '商住', '商办', '综合用地', '其他']:
            print(f'>>> {usage}...')
            # 土地性质
            self.land_usg([usage])
            # 统计
            self.land_stat()
            # to Dataframe
            df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
            df[usage] = df_['用地面积']

        # 合并商办、综合、其他
        df['商办及其他'] = df['综合用地'] + df['商办'] + df['其他']
        return df[['纯住宅', '商住', '商办及其他']].drop('汇总')

    def land_house(self, area_tuple, date_range):
        df = pd.DataFrame()

        # 土地可建
        self.land(area_tuple, date_range)
        self.land_stat()

        df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
        df['可建面'] = round(df_['建筑面积'] / 1e4, 2)
        df['楼面价'] = df_['成交楼板价']

        # 住宅成交
        df_ = self.monitor(date_range, ['上市面积', '成交面积', '成交均价'])
        df[['住宅供应', '住宅成交']] = round(df_['上市面积', '成交面积'] / 1e4, 2)
        df['住宅均价'] = df_['成交均价'].round().astype('int')

        df['地价房价比'] = round(df['楼面价'] / df['住宅均价'], 2)
        df['土地消化周期'] = round(df['可建面'] / df['住宅成交'], 2)
        df = df.drop('汇总')

        df1 = df[['可建面', '住宅供应', '住宅成交', '住宅面积', '土地消化周期']]
        df2 = df[['楼面价', '住宅均价', '地价房价比']]

        return df1, df2

    def zhuzhai_gxj(self):
        def ajust_df(df):
            # 面积换算成万方
            df.iloc[:, [0, 1]] = round(df.iloc[:, [0, 1]] / 1e4, 2)
            # 均价四舍五入为整数
            df.iloc[:, 2] = round(df.iloc[:, 2]).astype('int')
            # 丢弃汇总行
            return df.drop('汇总')

        # 年度走势
        df_year = self.monitor(self.year_range, ['供应面积', '成交面积', '成交均价'])
        df_year = ajust_df(df_year)

        # 分板块
        df_plate = self.monitor(self.year_range, ['供应面积', '成交面积', '成交均价'], index='区域')
        df_plate = ajust_df(df_plate)

        return df_year, df_plate

    def stock(self, stk):
        date_range = ('2009年01月', '2017年09月')
        df = self.monitor(date_range, ['供应面积', '成交面积']).drop('汇总')
        df['库存'] = 0
        # 计算滚动12个月月均成交面积
        df['去化速度'] = df['成交面积'].rolling(12).mean()

        # 以最后一期库存向过去推算
        for index in reversed(df.index):
            df.at[index, '库存'] = stk
            # 上期库存(在下一次迭代时赋值给['库存']) = 本期库存 - (本期上市 - 本期成交)
            stk -= df.at[index, '供应面积'] - df.at[index, '成交面积']

        # 去化周期 = 存量 / 去化速度速度
        df['去化周期'] = df['库存'] / df['去化速度']
        # 面积换算成万方
        df['库存'] = df['库存'] / 1e4
        # 只留两列，留两位小数
        df = df[['库存', '去化周期']].round(2)

        # 年度
        df_year = df['2009年01月':'2016年12月'].copy()
        for each in df_year.index:
            if '12月' not in each:
                df_year = df_year.drop(each)
        df_year = df_year.append(df['2017年09月':'2017年09月'])

        return df_year

    def structure(self):
        df_mjd = self.monitor(('2011年01月', '2017年09月'), ['成交套数'], column=['面积段'])
        df_djd = self.monitor(('2011年01月', '2017年09月'), ['成交套数'], column=['单价段'])
        df_zjd = self.monitor(('2011年01月', '2017年09月'), ['成交套数'], column=['总价段'])
        return df_mjd, df_djd, df_zjd
