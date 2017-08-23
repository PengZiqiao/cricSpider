from selenium import webdriver
from time import sleep

cookies = [
    {'domain': '2015.app.cric.com',
     'httpOnly': False,
     'name': 'BIGipServerpool_10.0.7',
     'path': '/',
     'secure': False,
     'value': '34013194.20480.0000'},
    {'domain': '.2015.app.cric.com',
     'httpOnly': False,
     'name': 'Hm_lpvt_dca78b8bfff3e4d195a71fcb0524dcf3',
     'path': '/',
     'secure': False,
     'value': '1503474356'},
    {'domain': '.cric.com',
     'expiry': 1504079106.982804,
     'httpOnly': True,
     'name': 'cric2015',
     'path': '/',
     'secure': False,
     'value': 'C101A9F6D98B733C1ACDA77C8F59A308A206181EC2BA28017D5B5482980ED2E6A61F798D31E2DCE553208E8550D5AACBB0F3C3FA1CC004B4836B8DF95794A14F5AF682E978DB4847EB82FF2D'},
    {'domain': '.2015.app.cric.com',
     'expiry': 1535010356,
     'httpOnly': False,
     'name': 'Hm_lvt_dca78b8bfff3e4d195a71fcb0524dcf3',
     'path': '/',
     'secure': False,
     'value': '1503474313'},
    {'domain': '.cric.com',
     'expiry': 1506066306.982859,
     'httpOnly': False,
     'name': 'cric2015_token',
     'path': '/',
     'secure': False,
     'value': 'username=c7WgHp4zScNtsm7KKQFU/Q==&verifycode=jjKbRj1fAaWKtVZoTqJCYQ==&token=ZT+H74zJs4mnpW93dJ9ZY0FGAy1+y8rRuGTlI9bd2c/2Bj5hjjCIbciiWaiJwSFh&usermobilephone=/xjKEQnYyec5HZvPfeoEsQ==&userid=Fiw/A5cH9X34OaMfzJTZzvuhDhEURUDyXzQkeVFh/K9SzWYASZeECe8KehkOlt37'}
]


class CricSpider:
    def __init__(self):
        self.url = 'http://2015.app.cric.com/'
        self.driver = webdriver.Chrome()
        self.driver.get(self.url)
        for cookie in cookies:
            self.driver.add_cookie(cookie)
        self.driver.get(self.url)

    def checkbox(self, name, value_list):
        """多选框
        :param
            name:表单控件的name
            value:表单的 value
        """
        chk = self.driver.find_elements_by_name(name)

        # 取消所有已选项目
        for each in chk:
            if each.is_selected():
                each.click()

        # 根据value选中需要的项目
        for each in chk:
            if each.get_attribute('value') in value_list:
                each.click()

    def stat_date(self, date, point):
        """选择日期
        date: '2017年' or '2017年08月' or '2017年x周/季度'
        point: 'Start' or 'End'
        """
        self.driver.find_element_by_xpath(f"//div[@id='divTimeRangeChoice_{point}']/div/a").click()
        self.driver.find_element_by_xpath(f"//div[@id='divTimeRangeChoice_{point}']//a[@data-value='{date}']").click()

    def stat_area(self, area):
        """板块选择"""
        self.driver.find_element_by_xpath("//div[@class='area_checked_content']/p").click()
        sleep(1)
        self.driver.find_element_by_xpath(f"//div[@class= 'plate_warp_layer']//label[text()='{area}']").click()
        self.driver.find_element_by_xpath(f"//div[@class= 'plate_warp_layer']//label[text()='请选择板块']").click()
        self.click('确定')

    def monitor_date(self, date, point):
        """选择日期
        date: '2017年' or '2017年08月' or '2017年x周/季度'
        point: 'Begin' or 'End'
        """
        self.driver.find_element_by_xpath(f"//span[contains(@name, 'Date{point}')]/div/a").click()
        self.driver.find_element_by_xpath(f"//span[contains(@name, 'Date{point}')]//a[@data-value='{date}']").click()

    def monitor_area(self, area):
        """板块选择"""
        self.driver.find_element_by_xpath("//div[@name='regionselecter_area_block']/div/div/p").click()
        sleep(1)
        self.driver.find_element_by_xpath(f"//div[@name='regionselecter_area_block']//label[text()='{area}']").click()
        self.driver.find_element_by_xpath(f"//div[@name='regionselecter_area_block']//label[text()='请选择板块']").click()
        self.click('确定')

    def monitor_usage(self, value_list):
        # 按两次全选，取消所有选中
        chk = self.driver.find_element_by_xpath("//input[contains(@name,'RoomUsageAll')]")
        chk.click()
        chk.click()

        # 按需要选中
        chk = self.driver.find_elements_by_xpath("//input[contains(@name,'Dims_RoomUsage')]")
        for each in chk:
            if each.get_attribute('value') in value_list:
                each.click()

    def monitor_radio(self, name):
        self.driver.find_element_by_xpath(f"//input[@data-text='{name}']").click()

    def click(self, name):
        self.driver.find_element_by_link_text(name).click()

    def stat_page(self, city):
        url = f'{self.url}Statistic/StatisticCenter/StatisticCenter?CityName={city}'
        self.driver.get(url)

    def monitor_page(self, city):
        url = f'{self.url}Statistic/MarketMonitor/MarketMonitoringIndex?CityName={city}'
        self.driver.get(url)


if __name__ == '__main__':
    c = CricSpider()

    """
    stat
    """
    c.stat_page('合肥')
    c.click('商品住宅统计')
    sleep(1)
    c.click('月供求')
    c.click('各板块')
    c.stat_date('2016年08月', 'Start')
    c.stat_date('2017年07月', 'End')
    c.stat_area('蜀山区')
    c.driver.find_element_by_xpath("//div[@class='form-item submit tr']/button").click()
    sleep(1)
    c.checkbox('stat_output', ['供应面积', '成交面积', '成交均价'])
    c.click('确定')
    c.driver.find_element_by_xpath("//div[@class='GDataGrid-customBtn']").click()

    """
    monitor
    """
    c.monitor_page('合肥')
    c.click('按分段')
    c.monitor_date('2016年08月', 'Begin')
    c.monitor_date('2017年07月', 'End')
    c.monitor_radio('板块')
    c.monitor_area('蜀山区')
    c.monitor_usage('普通住宅')
    c.driver.find_element_by_xpath("//div[@class='search_sure_btn pt5']/input").click()
    c.checkbox('cbx_title', '成交套数')
    c.click('确定')
    