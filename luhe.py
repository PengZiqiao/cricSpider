from datetime import date
from query import CricQuery, ToExcel

excel = ToExcel()
c = CricQuery()

stock = {'全选': 5319022}
today = date.today()
area = '全选'

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

c.driver.close()