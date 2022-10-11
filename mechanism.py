#!/usr/bin/env python
#-*- coding:utf-8 -*-


from glob import glob
from lib2to3.pgen2 import driver
from time import sleep
from urllib import request
from selenium import webdriver
from lxml import etree
import xlwt

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import selenium.webdriver.support.ui as ui
import re



chromeDriver_path = 'C:\SoftWares\chromedriver_win32\chromedriver.exe'
profile_directory = r'--user-data-dir=C:\\Users\\mzy\\AppData\\Local\\Google\\Chrome\\User Data'


file = xlwt.Workbook(encoding='utf-8', style_compression=0)
col_header = ['机构名称', '管理资本量', '主投行业', '主投轮次', '近一年直投数量', '最近投资时间', '最近直投项目', '经常合作机构']
sheet_one = file.add_sheet('机构信息', cell_overwrite_ok=True)

# sheet_one = file.add_sheet('管理基金_基金管理人', cell_overwrite_ok=True)
# sheet_two = file.add_sheet('管理基金_直投基金', cell_overwrite_ok=True)
# sheet_three = file.add_sheet('管理基金_母基金', cell_overwrite_ok=True)
# sheet_four = file.add_sheet('管理基金_其他投资主体', cell_overwrite_ok=True)
# sheet_five = file.add_sheet('投资基金', cell_overwrite_ok=True)


main_page_num = 705

headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    'Connection': 'keep-alive',
    'Content-Length': '99',
    'Content-Type': 'application/json',
    'Cookie': 'SESSION=ZWFkYmYyZWQtMDgwMi00MWYzLWEwNTctYjc1YTdjYjA1NTRi; Hm_lvt_0d052dd2e4e34a214d13e86591990a09=1663920386,1663920681; Hm_lpvt_0d052dd2e4e34a214d13e86591990a09=1663920692',
    'Host': 'max.pedata.cn',
    'HTTP-X-TOKEN': '60d2d5e1fc6ed532f175d633240b2075378d8d118446398b2047d169b3333c40',
    'Origin': 'https://max.pedata.cn',
    'Referer': 'https://max.pedata.cn/client/org/active',
    'sec-ch-ua': '"Google Chrome";v="105", "Not)A;Brand";v="8", "Chromium";v="105"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'testxor': 'testxor',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
}



class mechanism:
    def __init__(self, drive):
        self.chrome_driver = drive

        self.main_page = 1
        self.mechanism_num = 1

        self.manage_manager_num = 0
        self.manage_direct_invest_num = 0
        self.manage_mother_fund_num = 0
        self.manage_other_invest_num = 0
        self.invest_data_num = 0

        self.manage_manager_page = 0
        self.manage_direct_invest_page = 0
        self.manage_mother_fund_page = 0
        self.manage_other_invest_page = 0
        self.invest_data_page = 0

        self.main_data = {
            'name' : '',
            'money' : '',
            'industry' : '',
            'rounds' : '',
            'invest_in_one' : '',
            'lasted_time' : '',
            'recent_project' : '',
            'coop_institu': ''
        }
        self.manage_manager_data = {
            'manager' : [],
            'filing_time' : [],
            'filing_numb' : [],
            'establish_time' : [],
            'legal_person' : [],
            'fund' : [],
            'money' : [],
            'invest_or_quit' : [],
            'registration' : []
        }        
        self.manage_direct_invest_data = {
            'fund_name' : [],
            'money' : [],
            'filing_time' : [],
            'filing_numb' : [],
            'establish_time' : [],
            'manager' : [],
            'invest_num' : [],
            'lasted_time' : [],
            'recent_project' : []
        }        
        self.manage_mother_fund_data = {
            'fund_name' : [],
            'money' : [],
            'filing_time' : [],
            'filing_numb' : [],
            'establish_time' : [],
            'manager' : [],
            'invest_num' : [],
            'lasted_time' : []
        }    
        self.manage_other_invest_data = {
            'name' : [],
            'establish_time' : [],
            'money' : [],
            'legal_person' : [],
            'registration' : [],
            'lasted_time' : [],
            'recent_project' : [],
            'invest_num' : []
        } 
        self.invest_data = {
            'name' : [],
            'time' : [],
            'money_out': [],
            'fund_name': [],
            'mechanism' : [],
            'money_tar' : [],
            'other_LP' : []
        }

        self.main_xpath = {
            'name' :            '//table/tbody/tr[%d]/td[1]/div/div/div[2]/a/text()'%self.mechanism_num,
            'money' :           '//table/tbody/tr[%d]/td[2]/div/text()'%self.mechanism_num,
            'industry' :        '//table/tbody/tr[%d]/td[3]/div/div/span/text()'%self.mechanism_num,
            'rounds' :          '//table/tbody/tr[%d]/td[4]/div/div/text()'%self.mechanism_num,
            'invest_in_one' :   '//table/tbody/tr[%d]/td[5]/div/text()'%self.mechanism_num,
            'lasted_time' :     '//table/tbody/tr[%d]/td[6]/div/text()'%self.mechanism_num,
            'recent_project' :  '//table/tbody/tr[%d]/td[7]/div/div/div/div/div/a/text()'%self.mechanism_num,
            'coop_institu':     '//table/tbody/tr[%d]/td[8]/div/div/div/div/div/a/text()'%self.mechanism_num
        }
        self.manage_manager_xpath = {
            'manager' : [],
            'filing_time' : [],
            'filing_numb' : [],
            'establish_time' : [],
            'legal_person' : [],
            'fund' : [],
            'money' : [],
            'invest_or_quit' : [],
            'registration' : []
        }        
        self.manage_direct_invest_xpath = {
            'fund_name' : [],
            'money' : [],
            'filing_time' : [],
            'filing_numb' : [],
            'establish_time' : [],
            'manager' : [],
            'invest_num' : [],
            'lasted_time' : [],
            'recent_project' : []
        }        
        self.manage_mother_fund_xpath = {
            'fund_name' : [],
            'money' : [],
            'filing_time' : [],
            'filing_numb' : [],
            'establish_time' : [],
            'manager' : [],
            'invest_num' : [],
            'lasted_time' : []
        }    
        self.manage_other_invest_xpath = {
            'name' : [],
            'establish_time' : [],
            'money' : [],
            'legal_person' : [],
            'registration' : [],
            'lasted_time' : [],
            'recent_project' : [],
            'invest_num' : []
        } 
        self.invest_xpath = {
            'name' : [],
            'time' : [],
            'money_out': [],
            'fund_name': [],
            'mechanism' : [],
            'money_tar' : [],
            'other_LP' : []
        }

    def update_xpath(self):
        self.main_xpath = {
            'name' :            '//table/tbody/tr[%d]/td[1]/div/div/div[2]/a/text()'%self.mechanism_num,
            'money' :           '//table/tbody/tr[%d]/td[2]/div/text()'%self.mechanism_num,
            'industry' :        '//table/tbody/tr[%d]/td[3]/div/div/span/text()'%self.mechanism_num,
            'rounds' :          '//table/tbody/tr[%d]/td[4]/div/div/text()'%self.mechanism_num,
            'invest_in_one' :   '//table/tbody/tr[%d]/td[5]/div/text()'%self.mechanism_num,
            'lasted_time' :     '//table/tbody/tr[%d]/td[6]/div/text()'%self.mechanism_num,
            'recent_project' :  '//table/tbody/tr[%d]/td[7]/div/div/div/div/div/a/text()'%self.mechanism_num,
            'coop_institu':     '//table/tbody/tr[%d]/td[8]/div/div/div/div/div/a/text()'%self.mechanism_num
        }

    def get_main_data(self):
        page_text = self.chrome_driver.page_source
        page_html = etree.HTML(page_text)
        keys = list(self.main_data.keys())
        for key in keys:
            self.main_data[key] = page_html.xpath(self.main_xpath[key])

    def write_main_header(self, sheet):
        keys = list(self.main_data.keys())
        for key, i in zip(keys, range(8)):
            sheet.write(0, i, key)

    def write_main_data(self, sheet):
        keys = list(self.main_data.keys())
        for key, i in zip(keys, range(8)):
            data = self.main_data[key]
            input_data = ''
            for data_temp in data:
                input_data += str(data_temp) + ' '
            sheet.write((int(self.main_page)-1)*10+int(self.mechanism_num), i, self.main_data[key])


def driver_Init():
    option = webdriver.ChromeOptions()
    option.add_argument(profile_directory)

    chrome_driver = webdriver.Chrome(executable_path=chromeDriver_path, options=option)
    chrome_driver.maximize_window()
    return chrome_driver


def click_button(chrome_driver, xpath):
    wait = ui.WebDriverWait(chrome_driver, 10)
    wait.until(lambda driver:chrome_driver.find_element(by=By.XPATH, value=xpath)).click()
    sleep(0.5)

def focus_on_new_lab(chrome_driver):
    all_handles = chrome_driver.window_handles
    chrome_driver.switch_to.window(all_handles[-1])

def close_new_lab(chrome_driver):
    focus_on_new_lab(chrome_driver)
    chrome_driver.close()
    focus_on_new_lab(chrome_driver)

def get_num(string):
    num = re.findall(r'(.*?)', string)
    return int(num)

def send_request(driver, url, params, method='POST'):
    if method == 'GET':
        parm_str = ''
        for key, value in params.items():
            parm_str = parm_str + key + '=' + str(value) + '&'
        if parm_str.endswith('&'):
            parm_str = '?'+parm_str[:-1]
        driver.get(url + parm_str)
    else:
        jquery = open('.\jquery-2.1.3.min\jquery-2.1.3.min.js', "r").read()
        driver.execute_script(jquery)
        ajax_query = '''
                    $.ajax('%s', {
                    type: '%s',
                    data: %s,
                    headers: %s,
                    module: true,
                    crossDomain: true,
                    xhrFields: {
                     withCredentials: true
                    },
                    success: function(){}
                    });
                    ''' % (url, method, params, headers)

        ajax_query = ajax_query.replace(" ", "").replace("\n", "")
        resp = driver.execute_script("return " + ajax_query)
        return resp



if __name__ == "__main__":
    chrome_driver = driver_Init()
    MECH = mechanism(chrome_driver)

    chrome_driver.get('https://max.pedata.cn/client/org/active')
    sleep(1)

    # MECH.write_main_header(sheet_one)
    for i in range(8):
        sheet_one.write(0, i, col_header[i])

    for MECH.main_page in range(1,main_page_num+1):
        click_button(chrome_driver, '/html/body/div/div/div/div[2]/div[3]/div[2]/ul/li[@title=%s]'%MECH.main_page)
        main_page_text = chrome_driver.page_source
        main_page_html = etree.HTML(main_page_text)
        for MECH.mechanism_num in range(1,11):
            MECH.update_xpath()
            MECH.get_main_data()
            MECH.write_main_data(sheet_one)

            # for key in list(MECH.main_data.keys()):
                # print(MECH.main_data[key], end=None)   
            
            # mech_path = '//table/tbody/tr[%s]/td[1]/div/div/div[1]/a'%MECH.mechanism_num
            # click_button(chrome_driver, mech_path)


            # mech_path = '//table/tbody/tr[%s]/td[1]/div/div/div[1]/a'%MECH.mechanism_num
            # mech_url = 'https://max.pedata.cn' + main_page_html.xpath(mech_path+'/@href')[0]
            # chrome_driver.get(mech_url)

            # 选择机构中的“管理基金”
            # ele = chrome_driver.find_element(by=By.CLASS_NAME, value='ant-tabs-tab-prev ant-tabs-tab-btn-disabled ant-tabs-tab-arrow-show').click()
            # chrome_driver.execute_script("arguments[0].focus();", ele)
            # manager_path = '/html/body/div/div/div/div/div/div[2]/div/div[1]/div/span[2]'
            # click_button(chrome_driver, mech_path)

            # manager_payload = {'module': "机构", 'pagetype': "详情", 'pageitem': "管理基金/基金管理人", 'params': "965763"}
            # print(send_request(chrome_driver, 'https://max.pedata.cn/api/sso/log/page', manager_payload))


    file.save('./fund.xls')
    chrome_driver.close
    chrome_driver.quit