import os
from time import sleep

import pandas as pd

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select

from consts import cookies, areas, stocks

import json


class CricSpider:
    def __init__(self, city):
        self.city = city
        # 初始化浏览器、网址并访问
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 10)
        self.url = 'http://2015.app.cric.com/'
        print('>>> opening...')
        self.driver.get(self.url)
        # 添加登陆后的cookie重新访问
        for cookie in cookies:
            self.driver.add_cookie(cookie)
        self.driver.get(self.url)
        print('>>> login...')
        sleep(2)
        print('>>> ok!')

    #################################################
    # 通用功能
    #################################################
    def checkbox(self, id_all, value_list):
        """
        在多选框中选择所需选项
        :param id_all: “全选”的id中包含的字符串
        :param value_list: 由所需选项label名称组织成的列表
        """
        # 单击一下全选，如果单击后为选中状态则再单击一次取消
        chk = self.driver.find_element_by_xpath(f"//input[contains(@id,'{id_all}')]")
        chk.click()
        if chk.is_selected():
            chk.click()

        # 按需要选中
        for each in value_list:
            self.label(each)

    def select(self, id, value):
        """
        下拉列表
        :param name:表单控件的name
        :param value:表单的 value 或visible_text
        """
        s = Select(self.driver.find_element_by_id(id))
        s.select_by_visible_text(value)

    def label(self, name, index=0, pause=0.2):
        """
        单击label元素
        :param name:
        :param index: 出先多个同名label时，同过下标确定，默认为0
        """
        self.driver.find_elements_by_xpath(f"//label[text()='{name}']")[index].click()
        sleep(pause)

    def click(self, name, pause=0.2):
        """
        单击链接元素
        :param name: 链接的link_text
        :param pause: 单击后，等待的时间，以防止页面加载未完成，默认0.2秒
        """
        self.driver.find_element_by_link_text(name).click()
        sleep(pause)

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

    #################################################
    # 市场监测相关
    #################################################
    def monitor_page(self):
        """跳转到市场监测页面"""
        url = f'{self.url}Statistic/MarketMonitor/MarketMonitoringIndex?CityName={self.city}'
        self.driver.get(url)
        sleep(0.2)
        self.wait.until(lambda driver: driver.find_element_by_xpath("//input[@value='统计']"))
        # 直接进入综合分析
        self.click('综合分析')
        sleep(0.2)
        self.wait.until(lambda driver: driver.find_element_by_xpath("//input[@value='统计']"))

    def date(self, date, point):
        """
        选择日期
        :param date: '2017年' or '2017年08月' or '2017年x周/季度'
        :param point: 'Begin' or 'End'
        """
        self.driver.find_element_by_xpath(f"//span[contains(@name, 'Date{point}')]/div/a").click()
        sleep(0.1)
        self.driver.find_element_by_xpath(f"//span[contains(@name, 'Date{point}')]//a[@data-value='{date}']").click()
        sleep(0.1)

    def date_range(self, begin, end):
        """选择统计时间为年度、月度或季度，并同时设置起、止两个时间"""
        if '月' in begin:
            self.label('月度')
        elif '季度' in begin:
            self.label('季度')
        else:
            self.label('年度')

        self.date(begin, 'Begin')
        self.date(end, 'End')

    def area(self, area):
        """区域选择"""
        # 找到“不限”，单击
        self.driver.find_element_by_xpath("//div[@name='regionselecter_region_block']/div/div/p").click()
        sleep(0.5)

        area_xpath = f"//div[@name='regionselecter_region_block']//label[text()='{area}']"
        self.wait.until(lambda driver: driver.find_element_by_xpath(area_xpath).is_displayed())
        self.driver.find_element_by_xpath(area_xpath).click()
        sleep(0.2)

        # 确定
        self.click('确定')

    def area2(self, area, area2=['请选择板块']):
        """板块选择，以查询分片区数据"""
        # 找到“不限”，单击
        self.driver.find_element_by_xpath("//div[@name='regionselecter_area_block']/div/div/p").click()

        # area_xpath为需选中区域的label，待其出现后单击选中
        area_xpath = f"//div[@name='regionselecter_area_block']//label[text()='{area}']"
        self.wait.until(lambda driver: driver.find_element_by_xpath(area_xpath).is_displayed())
        self.driver.find_element_by_xpath(area_xpath).click()
        sleep(0.2)

        # 根据area2选中具体片区，默认为全部片区
        for each in area2:
            self.driver.find_element_by_xpath(
                f"//div[@name='regionselecter_area_block']//label[text()='{each}']").click()
            sleep(0.2)

        # 确定
        self.click('确定')

    def room_usg(self, value_list):
        """市场监测下的物业类型"""
        self.checkbox('RoomUsageAll', value_list)

    def monitor_output(self, value_list):
        """市场监测下的统计输出项"""
        self.checkbox('cbx_title_all', value_list)
        self.click('确定')
        sleep(0.2)

    def monitor_stat(self):
        """按一下市场监测的统计键"""
        self.driver.find_element_by_xpath("//input[@value='统计']").click()
        sleep(0.5)

    def monitor(self, date_range, output, area=[], area2=(), usg=['普通住宅'], index=False, column=False):
        """市场监测综合分析查询动作
        :param date_range: 时间段元组
        :param output: 输出项列表
        :param area:区域 默认不限
        :param area2: 板块
        :param usg : 物业类型列表 默认普通住宅
        :param index: 纵轴 默认时间
        :param column: 横轴 默认无
        :return df
        """
        # 跳转到市场监测-综合分析
        self.monitor_page()

        # 纵轴
        if index:
            self.label(index)
        # 横轴
        if column:
            self.label(column, index=1)

        # 统计时间
        self.date_range(*date_range)

        # 地域选择
        if area:
            self.area(area)
        if area2:
            self.label('板块')
            self.area2(*area2)

        # 物业类型
        self.room_usg(usg)

        # 按下统计
        self.monitor_stat()

        # 选择输出项
        self.monitor_output(output)

        # 下载返回df
        df = self.download2df()
        return df

    #################################################
    # 土地库相关
    #################################################
    def land_page(self):
        """跳转到土地顾问的土地统计页面"""
        url = f'{self.url}land/Statistic/Statistic/StatisticIndex?MenuKey=200001'
        self.driver.get(url)
        sleep(1.5)

    def land_date(self, date, point):
        """
        选择日期
        :param date: '2017年' or '2017年08月' or '2017年第02季度'
        :param point: 'Begin' or 'End'
        """
        if '月' in date:
            by = 'month'
        elif '季度' in date:
            by = 'season'
        else:
            by = 'year'

        self.driver.find_element_by_xpath(f"//div[@id='{by}{point}']//b").click()
        sleep(0.1)
        self.driver.find_element_by_xpath(f"//div[@id='{by}{point}']//a[@title='{date}']").click()
        sleep(0.1)

    def land_date_range(self, begin, end):
        """选择统计时间为年度、月度或季度，并同时设置起、止两个时间"""
        if '月' in begin:
            self.label('月度')
        elif '季度' in begin:
            self.label('季度')
        else:
            self.label('年度')

        self.land_date(begin, 'Begin')
        self.land_date(end, 'End')

    def land_area(self, area_tuple):
        """选择省份,城市,区域"""
        # 分别执行点开菜单,等待选项显示,选中选项
        keys = ['Province', 'City', 'Region']
        if len(area_tuple) == 4:
            keys.append('District')

        for i, key in enumerate(keys):
            self.driver.find_element_by_id(f'select{key}').click()
            self.wait.until(lambda driver: driver.find_element_by_link_text(area_tuple[i]).is_displayed())
            self.click(area_tuple[i])

        # 确定
        self.driver.find_element_by_class_name('areayes').click()

    def land_city(self, city):
        """在自定义城市中选择一个"""
        self.driver.find_element_by_id('cityConfig').click()
        self.select('selconfigcity', city)

    def land_method(self, value_list):
        """出让方式"""
        self.checkbox('allMethod', value_list)

    def land_usg(self, value_list):
        """土地性质"""
        self.checkbox('allProperty', value_list)

    def land_output(self, value_list):
        """土地输出项"""
        # 因为'幅数'这个选项是个奇葩，所以不得不删除它才能进行正常选择,再单独按查id的方法选中
        if '幅数' in value_list:
            value_list.remove('幅数')
            self.checkbox('allColumn', value_list)
            self.driver.find_element_by_id('FUSHU').click()
        else:
            self.checkbox('allColumn', value_list)
        self.driver.find_element_by_id('rowList').click()

    def land_stat(self):
        """按一下土地市场的统计键"""
        self.driver.find_element_by_xpath("//button[text()='统计']").click()
        sleep(0.5)

    def land(self, area_tuple, date_range):
        """跳转至土地统计页面,设置好时间,区域及交易方式"""
        # 土地统计页面
        self.land_page()
        # 时间 年度
        self.land_date_range(*date_range)
        # 区域 详细到板块
        self.land_area(area_tuple)
        # 交易方式
        self.land_method(['招拍挂土地'])


class Query(CricSpider):
    this_year = ('2017年', '2017年')
    this_year_monthly = ('2017年01月', '2017年09月')
    short_year_range = ('2013年', '2017年')
    year_range = ('2009年', '2017年')

    def gxj(self, plate, pianqu, date_range):
        def ajust_df(df):
            # 面积换算成万方
            df.iloc[:, [0, 1]] = round(df.iloc[:, [0, 1]] / 1e4, 2)
            # 均价四舍五入为整数
            df.iloc[:, 2] = round(df.iloc[:, 2]).astype('int')
            # 丢弃汇总行
            return df.drop('汇总')

        df = self.monitor(date_range, ['供应面积', '成交面积', '成交均价'], area2=(plate, [pianqu]))
        return ajust_df(df)

    def pianqu_stock(self, area_tuple):
        plate = area_tuple[0]
        pianqu = area_tuple[1]
        date_range = ('2009年01月', '2017年10月')
        df = self.monitor(date_range, ['供应面积', '成交面积'], area2=(plate, pianqu)).drop('汇总')
        df['库存'] = 0
        # 计算滚动6个月月均成交面积
        df['去化速度'] = df['成交面积'].rolling(6).mean()

        # 以2017年10月库存向过去推算
        stk = sum(list(stocks[pianqu] for pianqu in areas[plate]))
        for index in reversed(df.index):
            df.at[index, '库存'] = stk
            # 上期库存(在下一次迭代时赋值给['库存']) = 本期库存 - (本期上市 - 本期成交)
            stk -= df.at[index, '供应面积'] - df.at[index, '成交面积']

        # # 向未来推算
        # df_ = df['2017年01月':date_range[1]]
        # stk = sum(list(stocks[pianqu] for pianqu in areas[plate]))
        # for index in df_.index:
        #     # 本期库存 = 上期库存 + (本期上市 - 本期成交)
        #     stk += df_.at[index, '供应面积'] - df_.at[index, '成交面积']
        #     df_.at[index, '库存'] = stk

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
        df_year = df_year.append(df['2017年08月':'2017年08月'])

        # 月度
        df_month = df['2017年01月':'2017年08月']

        return df_year, df_month

    def potential_land_in_project(self):
        def get_list():
            projects = self.driver.find_elements_by_xpath(
                '//div[@class="GDataGrid-lock"]/div[@class="GDataGrid-body"]//tr')
            return list(x.text for x in projects)

        df = pd.DataFrame(columns=['片区', '总建', '住宅上市', '未推土地', '未推土地修正'])
        self.driver.get(f'{self.url}Projects/Search/ProjectSearch?CityName={self.city}')
        sleep(2)
        self.click('在售')
        self.click('预售')
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.click('普通住宅')
        sleep(1)

        # 每页显示50条
        self.driver.find_element_by_xpath('//div[@class="g-form-select-show"]').click()
        sleep(0.2)
        self.label('50')
        sleep(1)

        # 按名称排序
        self.driver.find_element_by_xpath('//td[text()="项目名称"]').click()
        sleep(2)

        # 获取页数
        page = self.driver.find_element_by_xpath('//span[text()="页数："]').text
        page = int(page.replace('页数：', ''))

        # 每一页
        zongjian = ''
        for p in range(1, page):
            # 获取项目列表
            projects = get_list()

            pass

            # 每个项目
            for i, proj in enumerate(projects):
                # 点开详情页
                self.click(proj)
                sleep(1)

                # 　切换窗口
                handles = self.driver.window_handles
                self.driver.switch_to.window(handles[1])

                # 总建
                try:
                    zj = self.driver.find_element_by_xpath('//span[starts-with(text(),"总建面积:")]').text
                    # 如果和前一个项目总建相等，则为同一个项目不同期，不重复增加总建
                    zongjian = '0' if zj == zongjian else zj
                    df.at[proj, '总建'] = float(zongjian.replace('总建面积:', '').replace(',', '').replace('㎡', ''))
                except:
                    self.driver.close()
                    self.driver.switch_to.window(handles[0])
                    continue

                # 已上市
                try:
                    shangshi = self.driver.find_element_by_xpath('//label[text()="总建面积"]/../div').text
                    df.at[proj, '住宅上市'] = float(shangshi.replace(',', '').replace('㎡', ''))
                except:
                    df.at[proj, '住宅上市'] = 0

                # 片区
                try:
                    pianqu = self.driver.find_element_by_xpath('//label[contains(text(),"板块")]/../div').text
                    df.at[proj, '片区'] = pianqu
                    print(f"第{p}页，第{i+1}个项目:{proj} 片区:{pianqu} {zongjian} 上市面积:{shangshi}")
                except:
                    self.driver.close()
                    self.driver.switch_to.window(handles[0])
                    continue

                # 回到上级页面
                self.driver.close()
                self.driver.switch_to.window(handles[0])

            # 除非最后一页，否则跳转下一页
            if not p == page:
                self.driver.find_element_by_xpath('//span[@class="arrow arrow-next"]').click()
                sleep(1)

        df['未推土地'] = df['总建'] - df['住宅上市']
        df['未推土地修正'] = df['未推土地'].apply(lambda x: 0 if x < 0 else x / 1e4)
        df = df.replace('蜀山新产业园板块', '蜀山产业园板块')

        return df.round(2)

    def land_weigongkai_project(self):
        def get_list():
            projects = self.driver.find_elements_by_xpath(
                '//div[@class="GDataGrid-lock"]/div[@class="GDataGrid-body"]//tr')
            return list(x.text for x in projects)

        df = pd.DataFrame(columns=['总建', '总建修正'])
        self.driver.get(f'{self.url}Projects/Search/ProjectSearch?CityName={self.city}')
        sleep(2)
        self.click('未公开')
        self.driver.find_element_by_class_name('cric_more_search_condition').click()
        self.click('普通住宅')
        sleep(1)

        # 每页显示50条
        self.driver.find_element_by_xpath('//div[@class="g-form-select-show"]').click()
        sleep(0.2)
        self.label('50')
        sleep(1)

        # 按名称排序
        self.driver.find_element_by_xpath('//td[text()="项目名称"]').click()
        sleep(2)

        # 获取页数
        page = self.driver.find_element_by_xpath('//span[text()="页数："]').text
        page = int(page.replace('页数：', ''))

        # 每一页
        zongjian = ''
        for p in range(1, page):
            # 获取项目列表
            projects = get_list()

            pass

            # 每个项目
            for i, proj in enumerate(projects):
                # 点开详情页
                self.click(proj)
                sleep(1)

                # 　切换窗口
                handles = self.driver.window_handles
                self.driver.switch_to.window(handles[1])

                # 总建
                try:
                    zj = self.driver.find_element_by_xpath('//span[starts-with(text(),"总建面积:")]').text
                    # 如果和前一个项目总建相等，则为同一个项目不同期，不重复增加总建
                    zongjian = '0' if zj == zongjian else zj
                    df.at[proj, '总建'] = float(zongjian.replace('总建面积:', '').replace(',', '').replace('㎡', ''))
                except:
                    self.driver.close()
                    self.driver.switch_to.window(handles[0])
                    continue

                # 片区
                try:
                    pianqu = self.driver.find_element_by_xpath('//label[contains(text(),"板块")]/../div').text
                    df.at[proj, '片区'] = pianqu
                    print(f"第{p}页，第{i+1}个项目:{proj} 片区:{pianqu} {zongjian}")
                except:
                    self.driver.close()
                    self.driver.switch_to.window(handles[0])
                    continue

                # 回到上级页面
                self.driver.close()
                self.driver.switch_to.window(handles[0])

            # 除非最后一页，否则跳转下一页
            if not p == page:
                self.driver.find_element_by_xpath('//span[@class="arrow arrow-next"]').click()
                sleep(1)

        df['总建修正'] = df['总建'] / 1e4
        df = df.replace('蜀山新产业园板块', '蜀山产业园板块')

        return df.round(2)


if __name__ == '__main__':
    c = CricSpider('苏州')
    cookies = c.driver.get_cookies()
    with open('cookies.txt', 'w') as f:
        json.dump(cookies, f)
