import pandas as pd
from office import ToExcel
from consts import areas

path = r'E:\city_report\合肥\2017-09-19'
excel = ToExcel()

df_land = pd.read_excel(r'E:\city_report\合肥\land_all.xlsx', 'pivot', index_col=0)
df_land = df_land / 1e4
df_land = df_land.round(2).fillna(0)

for plate in areas:
    print(plate)
    # 可建面
    land = df_land.loc[areas[plate]]
    land.loc['sum'] = land.sum()

    # 板块土地
    book = f'{path}/{plate}土地.xlsx'
    df = pd.read_excel(book, '土地住宅面积', index_col=0)
    df['可建面'] = land.iloc[-1, -5:].tolist()
    excel.rewrite_sheet(book, '土地住宅面积', df)

    for pianqu in areas[plate]:
        print(plate, pianqu)

        book = f'{path}/{plate}_{pianqu}.xlsx'
        df = pd.read_excel(book, '年度库存', index_col=0)
        df = df[2:]
        excel.rewrite_sheet(book, '年度库存', df)

        df = pd.read_excel(book, '年度供销价', index_col=0)
        df = df[2:]
        df['可建面'] = land.loc[pianqu].tolist()
        df = df[['可建面', '供应面积', '成交面积', '成交均价']]
        excel.rewrite_sheet(book, '年度供销价', df)
