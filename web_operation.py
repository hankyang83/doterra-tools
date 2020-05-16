# -*- coding: utf-8 -*-

from selenium import webdriver
import pandas as pd
import time

class connection():
    def __init__(self, driver, timeout):
        self.driver = driver
        self.timeout = timeout
    
    def log_in(self, account, password):
        # check ready
        start = time.time()
        while True:
            if (time.time() - start) > self.timeout:
                print('web not ready...')
                break
            try:
                keyword = self.driver.find_element_by_xpath('//div[@class="title"]').text
                if keyword == 'Sign in':
                    break
            except:
                pass
        
        self.driver.find_element_by_name('uid').clear()
        self.driver.find_element_by_name('uid').send_keys(account)
        self.driver.find_element_by_name('password').clear()
        self.driver.find_element_by_name('password').send_keys(password)
        self.driver.find_element_by_xpath('//button[@type="submit"]').submit()
        # check log in
        start = time.time()
        while True:
            if (time.time() - start) >  self.timeout:
                print('logging timeout...')
                break
            try:
                self.driver.find_element_by_link_text('首頁').click()
                print('Logging success...')
                break
            except:
                pass
    
    def log_out(self):
        start = time.time()
        # check log out
        while True:
            if (time.time() - start) > self.timeout:
                print('log out timeout...')
                break
            try:
                self.driver.find_element_by_xpath('//a[@id="SignOut"]').click()
                print('Log out success')
                break
            except:
                pass
 
    def get_point(self):
        start = time.time()
        while True:
            if (time.time() - start) > self.timeout:
                print('getting point Timeout')
                break
            try:
                self.driver.find_element_by_link_text('首頁').click()
                self.driver.find_element_by_xpath('//a[@href="/index.cfm?tabsel=MainMenu"]').click()
                time.sleep(2)
                point = self.driver.find_element_by_xpath('//div[@class="ten wide middle aligned column"]').text.split('\n')[0]
                break
            except:
                pass
        return point

if __name__ == '__main__':
    start = time.time()
    timeout = 60
    file_name = 'account_info.xlsx'
    driver = webdriver.Chrome(executable_path='./chromedriver_win32/chromedriver.exe')
    driver.get('https://login.doterra.com/tw/zh-tw/sign-in')
    df = pd.read_excel(file_name)
    
    action = connection(driver, timeout)
    for idx, row in df.iterrows():
        account = str(int(row['ID']))
        passwd = row['PassWord']
        print(account)
        try:
            action.log_in(account, passwd)
        except:
            print('Login fail')
            print('-' * 50)
            continue
        point = action.get_point()
        print(point)
        df.loc[idx, 'Point'] = point
        try:
            action.log_out()
        except:
            print('sign out fail')
            break
        print('-' * 50)
    df.to_excel(file_name, index=False)
    print('Total time: {} sec'.format(time.time() - start))