import os
from time import sleep
import json
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select


def json_load(filename):
    path = os.getcwd()
    with open(f'{path}/{filename}.json', 'r') as f:
        return json.load(f)


def json_dump(obj, filename):
    path = os.getcwd()
    with open(f'{path}/{filename}.json', 'w') as f:
        json.dump(obj, f)


class CricSpider:
    def __init__(self):
        # 初始化浏览器、网址并访问
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 10)
        self.url = 'http://2015.app.cric.com/'
        print('>>> login...')
        self.driver.get(self.url)
        # 添加登陆后的cookie重新访问
        cookies = json_load('cookies')
        for cookie in cookies:
            self.driver.add_cookie(cookie)
        self.driver.get(self.url)
        print('>>> ok!')

    #################################################
    # 通用功能
    #################################################

    def loaded(self, name='数据说明'):
        """
        等待一个元素出现，以保证页面加载完成
        :param name: 页面上的字，通过link_text查找，默认为“数据说明”
        """
        self.wait.until(lambda driver: driver.find_element_by_link_text(name))

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

    #################################################
    # 市场监测相关
    #################################################

    def monitor_page(self, city):
        """跳转到市场监测页面"""
        url = f'{self.url}Statistic/MarketMonitor/MarketMonitoringIndex?CityName={city}'
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

        # area_xpath为需选中区域的label，待其出现后单击选中
        if area == '全选':
            sleep(0.5)
            for each in ['蜀山区', '庐阳区', '政务区', '高新区', '瑶海区', '滨湖区', '包河区', '经济区', '新站区', '肥东县', '肥西县', '长丰县']:
                area_xpath = f"//div[@name='regionselecter_region_block']//label[text()='{each}']"
                self.driver.find_element_by_xpath(area_xpath).click()
        else:
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

    def monitor(self, date_range, output, area=False, area2=False, usg=['普通住宅'], index=False, column=False):
        """市场监测综合分析查询动作
        :data_range: 时间段元组
        :output: 输出项列表
        :param area:区域 默认不限
        :param area2: 板块
        :param usg : 物业类型列表 默认普通住宅
        :param index: 纵轴 默认时间
        :param column: 横轴 默认无
        :return df
        """
        # 跳转到市场监测-综合分析
        self.monitor_page('合肥')

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
        self.loaded('土地统计')

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
        for i, key in enumerate(['Province', 'City', 'Region']):
            self.driver.find_element_by_id(f'select{key}').click()
            self.wait.until(lambda driver: driver.find_element_by_link_text(area_tuple[i]).is_displayed())
            self.click(area_tuple[i])

        # 全市
        if area_tuple[-1] == '全选':
            self.click('庐江县')
            self.click('巢湖市')

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


if __name__ == '__main__':
    c = CricSpider()
