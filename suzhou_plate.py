import pandas as pd
from cric import Query
from consts import areas, stocks


class CricQuery(Query):
    ###
    # 板块层面
    ###
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
            df[usage] = df_['幅数']

        # 合并商办、综合、其他
        df['商办及其他'] = df['综合用地'] + df['商办'] + df['其他']
        return df[['纯住宅', '商住', '商办及其他']].drop('汇总')

    def land_house(self, area_tuple):
        df = pd.DataFrame()

        # 土地可建
        self.land(area_tuple, self.short_year_range)
        self.land_stat()

        df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
        df['可建面'] = round(df_['建筑面积'] / 1e4, 2)
        df['楼面价'] = df_['成交楼板价']

        # 住宅成交
        plate = area_tuple[2]
        pianqu = areas[plate]
        df_ = self.monitor(self.short_year_range, ['成交面积', '成交均价'], area2=(plate, pianqu))
        df['住宅面积'] = round(df_['成交面积'] / 1e4, 2)
        df['住宅均价'] = df_['成交均价'].round().astype('int')

        df['地价房价比'] = round(df['楼面价'] / df['住宅均价'], 2)
        df = df.drop('汇总')

        df1 = df[['可建面', '住宅面积']]
        df2 = df[['楼面价', '地价房价比']]

        return df1, df2

    ###
    # 片区层面
    ###
    def gxj_year(self, plate, pianqu):
        df = pd.DataFrame()

        # 土地可建
        area_tuple = ('江苏省', self.city, plate, pianqu)
        self.land(area_tuple, self.year_range)
        self.land_stat()
        df_ = pd.read_html(self.driver.page_source, index_col=0, header=0)[0]
        df['可建面'] = round(df_['建筑面积'] / 1e4, 2)

        # 供销价
        df_ = self.gxj(plate, pianqu, self.year_range)

        return df.append(df_)

    def gxj_month(self, plate, pianqu):
        df = self.gxj(plate, pianqu, self.this_year_monthly)
        return df

    def stock(self, plate, pianqu, stk):
        date_range = ('2009年01月', '2017年09月')
        df = self.monitor(date_range, ['供应面积', '成交面积'], area2=(plate, [pianqu])).drop('汇总')
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

        # 月度
        df_month = df['2017年01月':'2017年09月']

        return df_year, df_month
