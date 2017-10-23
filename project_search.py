from cric import Query

if __name__ == '__main__':
    c = Query('苏州')
    # df = c.potential_land_in_project()
    df = c.land_weigongkai_project()
    df.to_excel('e:/weigongkai_land.xlsx')
