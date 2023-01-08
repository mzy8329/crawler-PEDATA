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
from selenium.webdriver.common.keys import Keys

import selenium.webdriver.support.ui as ui
import re

import pandas as pd
import random


chromeDriver_path = 'D:\Softwares\chromedriver_win32\chromedriver.exe'
profile_directory = r'--user-data-dir=C:\\Users\\YiLab1\\AppData\\Local\\Google\\Chrome\\User Data'


headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    'Connection': 'keep-alive',
    'Content-Length': '49',
    'Content-Type': 'application/json',
    'Cookie': 'Hm_lvt_0d052dd2e4e34a214d13e86591990a09=1672993186; Hm_lpvt_0d052dd2e4e34a214d13e86591990a09=1672994677; SESSION=YmY4ZTUxZTEtZmI0Mi00MWUxLTg3NjAtYTg4ZmQ4NWFhOTFl',
    'Host': 'max.pedata.cn',
    'HTTP-X-TOKEN': 'b9b42240909f825c24ca520d8d28255e2327a009463d6210e9accfb5586dd1bd',
    'Origin': 'https://max.pedata.cn',
    'Referer': 'https://max.pedata.cn/client/org/active',
    'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'testxor': 'testxor',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
}

xpaths = {
    'list_name': '/html/body/div/div/div/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/table/thead/tr/th',
    'export_but': '/html/body/div/div/div/div[2]/div[2]/div[2]/span[1]',
    'extend_but': '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/span[2]',
    'extend_list_but': '/html/body/div/div/div/div[2]/div[3]/div[3]/div[1]/div[2]/div/div[1]/div/div[1]/div/label[%s]/span[1]',
    'extend_list_name': '/html/body/div/div/div/div[2]/div[3]/div[3]/div[1]/div[2]/div/div[1]/div/div[1]/div/label%s',
    'extend_ok_but': '/html/body/div/div/div/div[2]/div[3]/div[3]/div[1]/div[3]/span',
    'pages': '/html/body/div/div/div/div[2]/div[3]/div[2]/ul/li%s',
    'headers': '/html/body/div[1]/div/div/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/table/thead/tr/th%s',
    'headers_name': '/html/body/div[1]/div/div/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[1]/table/thead/tr/th%s/div[1]/span[1]',
    'fond_datas': '/html/body/div/div/div/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[2]/table/tbody/tr%s',
    'fond_data' : '/html/body/div/div/div/div[2]/div[3]/div[1]/div/div/div[2]/div[1]/div[2]/table/tbody/tr%s/td',
    'to_page_input': '/html/body/div/div/div/div[2]/div[3]/div[2]/ul/li[10]/div[2]/input'
}

xpath_fond_format = ['/div', '/div/a', '/div/div', '/div/div/div/div/div/a', '/div/div/div', '/div/div/div[2]/a']

urls = {
    '正在募集基金': 'https://max.pedata.cn/client/fund/raise?menuItem=1',
    '美元基金': 'https://max.pedata.cn/client/fund/dollar?menuItem=4',
    '延期基金': 'https://max.pedata.cn/client/fund/defered?menuItem=6',
    '限售解禁基金': 'https://max.pedata.cn/client/fund/ipounlock?menuItem=7',
    '备案基金库': 'https://max.pedata.cn/client/fund/list?menuItem=5',
}

paths = {
    '正在募集基金': 'data/正在募集基金.xlsx',
    '美元基金': 'data/美元基金.xlsx',
    '延期基金': 'data/延期基金.xlsx',
    '限售解禁基金': 'data/限售解禁基金.xlsx',
    '备案基金库': 'data/备案基金库.xlsx'
}

max_pages = {
    '正在募集基金': 220,
    '美元基金': 81,
    '延期基金': 590,
    '限售解禁基金': 1039,
    '备案基金库': 7212
}



class Crawler:
    def __init__(self):
        pass


    def init(self, origin_url, extend_need, max_page, name, begin_page):
        self.chrome_driver = driver_Init()
        self.origin_url = origin_url
        self.max_page = max_page
        self.name = name
        self.begin_page = begin_page
        self.page = begin_page

        self.chrome_driver.get(self.origin_url)
        sleep(1)
        self.goto_page(begin_page)
        
        if(extend_need):
            click_button(self.chrome_driver, xpaths['extend_but'])
            main_page_text = self.chrome_driver.page_source
            main_page_html = etree.HTML(main_page_text)
            self.label_size = len(main_page_html.xpath(xpaths['extend_list_name']%''))
            
            for label in range(self.label_size):
                if(main_page_html.xpath(xpaths['extend_list_name']%'')[label].attrib['class'] == 'ant-checkbox-wrapper'):
                    click_button(self.chrome_driver, xpaths['extend_list_but']%(label+1))
                    sleep(0.1)
            click_button(self.chrome_driver, xpaths['extend_ok_but'])

        else:
            main_page_text = self.chrome_driver.page_source
            main_page_html = etree.HTML(main_page_text)
            self.label_size = len(main_page_html.xpath(xpaths['list_name']))

    def stop_driver(self):
        self.chrome_driver.close
        self.chrome_driver.quit
        del self.chrome_driver

        while(1):
            try:
                self.chrome_driver.execute_script('javascript:void(0);')
            except:
                break
            else:
                self.chrome_driver.quit
            sleep(0.1)


    def goto_page(self, page):
        self.chrome_driver.find_element(by=By.XPATH, value=xpaths['to_page_input']).send_keys(page)
        self.chrome_driver.find_element(by=By.XPATH, value=xpaths['to_page_input']).send_keys(Keys.ENTER)
        
    def get_data(self):
        page_text = self.chrome_driver.page_source
        page_html = etree.HTML(page_text)

        for page in range(self.begin_page, self.max_page+1):
            if(click_button(self.chrome_driver, xpaths['pages']%'[@title=%s]'%page) == False):
                return self.page
            self.page += 1
            page_text = self.chrome_driver.page_source
            page_html = etree.HTML(page_text)

            fond_num = len(page_html.xpath(xpaths['fond_datas']%''))
            for fond_index in range(fond_num):                
                for data_index in range(self.label_size):                   
                    find = False
                    for format in xpath_fond_format:
                        try:
                            data = page_html.xpath(xpaths['fond_data']%[fond_index+1]+'%s'%[data_index+1]+format)[0].text
                        except:
                            pass
                        else:
                            if data != None:
                                self.Output_datas[self.Headers[data_index]].append(data)
                                find = True
                                break
                    if find == False:
                        self.Output_datas[self.Headers[data_index]].append(None)

            sleep(5+random.randint(0,10)*0.1)
        
        return 0

    def file_init(self):
        page_text = self.chrome_driver.page_source
        page_html = etree.HTML(page_text)

        self.Headers = []
        for head_index in range(self.label_size):
            self.Headers.append(page_html.xpath(xpaths['headers_name']%'')[head_index].text)
        self.Output_datas = {head:[] for head in self.Headers}

    def file_write(self):
        df = pd.DataFrame(self.Output_datas)
        df.to_excel(paths[self.name])



def driver_Init():
    option = webdriver.ChromeOptions()
    option.add_argument(profile_directory)

    chrome_driver = webdriver.Chrome(executable_path=chromeDriver_path, options=option)
    chrome_driver.maximize_window()
    return chrome_driver



def click_button(chrome_driver, xpath):
    try:
        wait = ui.WebDriverWait(chrome_driver, 20)
        wait.until(lambda driver:chrome_driver.find_element(by=By.XPATH, value=xpath)).click()

    except:
        return False
    else:
        sleep(0.1)
        return True

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
    name = '备案基金库'

    MECH = Crawler()
    MECH.init(urls[name], True, max_pages[name], name, 1)
    MECH.file_init()

    over_page = MECH.get_data()
    while over_page != 0:
        MECH.stop_driver()
        print(over_page)
        
        sleep(1)

        MECH.init(urls[name], True, max_pages[name], name, over_page)
        over_page = MECH.get_data()

    MECH.file_write()
    MECH.stop_driver()