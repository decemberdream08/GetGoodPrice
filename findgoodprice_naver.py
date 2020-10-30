import datetime
import os
import re
import shutil
import time

import telepot
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

###################################################################
#   Define List, Variable
###################################################################
item_list = []
diff = []
Cur_price_column = 7
Old_price_column = 8
Cur_URL_column = 17
Old_URL_column = 18
Delivery_price_column = 9
row_number = 4
item_numbers = 0

###################################################################
#   Working Directory
###################################################################

### Excel File ###

#PATH = 'D:/Auto_Program/'
PATH = 'D:/01_MS_Work/02_Office/01_MS_Global/01_Global_Selling/'
File_Name = '판매상품관리'
File_extension = '.xlsx'
Excel_PATH = PATH + File_Name + File_extension
#print(datetime.date.today())
date = str(datetime.date.today())
Excel_PATH2 = PATH + File_Name + '_' + date + File_extension
excel = win32com.client.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(Excel_PATH)
ws = wb.Worksheets('판매상품목록')

### Log File ###
def write_log(msg):
    f = open(PATH + 'auto.log', 'a')
    f.write('[%s] %s\n' % (str(datetime.datetime.now()), msg))


### Looking for number of item from Excel file ###


for i in range(row_number, 100000):
    cell_value = ws.Cells(i, 3).Value
    if cell_value == None:
        break
    else:
        item_list.append(cell_value)
        ws.Cells(i, Old_price_column).Value = ws.Cells(
            i, Cur_price_column).Value
        ws.Cells(i, Old_URL_column).Value = ws.Cells(i, Cur_URL_column).Value
        ws.Cells(i, Cur_price_column).Value = ''
        ws.Cells(i, Cur_URL_column).Value = ''

opts = webdriver.ChromeOptions()
opts.add_argument('headless')
opts.add_argument('window-size=1920x1080')
opts.add_argument('disable-gpu')
#driver = webdriver.Chrome('D:/03_Study/01_Python/01_Code/02_Auto/chromedriver')
driver = webdriver.Chrome('D:/03_Study/01_Python/01_Code/02_Auto/chromedriver', options=opts)

driver.get('https://shopping.naver.com/')
sleep_time = 3

for item_name in item_list:
    try:
        try:
            write_log('%d. 아이템명 검색' % row_number)
            print('%d. 아이템명 검색' % row_number)
            
            if row_number == 4:
                elem = driver.find_element_by_class_name('co_srh_input')
            else:
                elem = driver.find_element_by_class_name('searchInput_search_input__3ZswN')


            #elem.clear()
            #elem.click()
            elem.send_keys(Keys.CONTROL + "a")
            elem.send_keys(Keys.DELETE)
            elem.click()


        except Exception as e:
            try:
                print(e)
            finally:
                e = None
                del e

    finally:
        print(item_name)
        elem.send_keys(item_name)
        elem.send_keys(Keys.RETURN)
        write_log('아이템명 검색 실시')
        time.sleep(sleep_time)

        ############################################################################################
        #### 가격비교 메뉴에서 가격비교 가능한 아이템 수를 확인 하고 가격비교를 선택하여 페이지 이동 ####
        ############################################################################################

        ### 가격비교 메뉴를 찾기 / 가격비교 가능한 아이템 수 찾기 ###
        try:
            try:
                elem = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, 'style_content__2T20F')))
                elem1 = elem.find_element_by_class_name('seller_filter_area')
                item_numbers_list = elem1.find_elements_by_xpath('./ul/li')  # re.findall('\\d+', elem.text)
                
                ## 가격비교 탭에서 비교 가능 숫자를 구한다.
                item_numbers = 0

                for li in item_numbers_list:
                    atag = li.find_element_by_class_name('subFilter_filter__3Y-uy')
                    result = atag.text.split("\n")
                    
                    if(result[0] == '가격비교'):
                        item_numbers = int(result[1])
                        break 

                li.click() ## 가격비교 탭을 클릭하여 가격비교 페이지로 이동
                
                if item_numbers == 1: ### 판매중단인지 Check(가격비교 해당 아이템이 1행만 있을때 체크)
                        time.sleep(sleep_time)
                        elem2 = elem.find_element_by_class_name('basicList_price_area__1UXXR') 
                        
                        if '판매중단' in elem2.text:
                            item_numbers = 0


                if item_numbers == 0:
                    write_log('item이 없습니다.')
                else:
                    write_log('아이템 개수 : %s' % item_numbers)
                
            except Exception as e:
                try:
                    write_log(e)
                    time.sleep(5)
                    write_log('Exception : 가격 비교 Click 재실시')

                    elem = driver.find_element_by_class_name('seller_filter_area')
                    item_numbers_list = elem.find_elements_by_xpath('./ul/li')  # re.findall('\\d+', elem.text)
                
                    ## 가격비교 탭에서 비교 가능 숫자를 구한다.
                    item_numbers = 0

                    for li in item_numbers_list:
                        atag = li.find_element_by_class_name('subFilter_filter__3Y-uy')
                        result = atag.text.split("\n")
                        
                        if(result[0] == '가격비교'):
                            item_numbers = int(result[1])
                            break 
                    
                    li.click() ## 가격비교 탭을 클릭하여 가격비교 페이지로 이동
                
                    if item_numbers == 1: ### 판매중단인지 Check(가격비교 해당 아이템이 1행만 있을때 체크)
                            time.sleep(sleep_time)
                            elem2 = elem.find_element_by_class_name('basicList_price_area__1UXXR') 
                            
                            if '판매중단' in elem2.text:
                                item_numbers = 0


                    if item_numbers == 0:
                        write_log('item이 없습니다.')
                    else:
                        write_log('아이템 개수 : %s' % item_numbers)

                    
                finally:
                    e = None
                    del e

        finally:
            write_log('가격 비교 클릭 완료')

        ### 가격비교 페이지에서 첫번째 가격비교 아이템을 클릭
        if item_numbers != 0:
            time.sleep(sleep_time)
            try:
                try:
                    elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'basicList_link__1MaTN')))
                    elem.click()
                    time.sleep(sleep_time)
                except Exception as e:
                    try:
                        write_log(e)
                        time.sleep(5)
                        elem = driver.find_element_by_xpath('//li[@data-expose-rank="1"]/div')
                        elem.click()
                        time.sleep(5)
                        write_log(elem)
                        write_log('goto exception ~~~!!!!!')
                    finally:
                        e = None
                        del e

            finally:
                write_log('가격 비교 창에서 첫번째 item을 선택/가격비교 새창 열림')
                
            window_num = 1

            driver.switch_to.window(driver.window_handles[window_num]) ### 새로운 창이 열리므로, 새창으로 Switch

            ### 가격 비교 새창에서 배송비 포함 클릭 ###
            try:
                try:
                    ##elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//a[@data-filter-name="withFee"]')))
                    elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[text()="배송비포함"]/a')))
                    elem.click()
                    time.sleep(sleep_time)
                except Exception as e:
                    try:
                        write_log(e)
                        time.sleep(5)
                        elem = driver.find_element_by_xpath('//a[@data-filter-name="withFee"]')
                        elem.click()
                        time.sleep(5)
                    finally:
                        e = None
                        del e

            finally:
                write_log('가격 비교 창에서 배송비 포함 선택')

            ### 배송비 및 배송비 포함 가격 확인 ###
            try:
                try:
                    elem = driver.find_element_by_class_name('productMain_price_blue__sTy-3')
                    elem = elem.find_element_by_tag_name('tbody')
                    elem = elem.find_element_by_tag_name('tr')
                    price_class = elem.find_element_by_class_name('productMain_price__1Gv9M')
                    a_tag = price_class.find_element_by_tag_name('a')
                    item_price = a_tag.text.replace(',', '')
                    ##item_price = item_price.replace('최저', '')
                    item_price = int(item_price)
                    item_url = a_tag.get_attribute('href')
                    gift_class = elem.find_element_by_class_name('productMain_gift__Ngvlo')

                    if gift_class.text.find('무료배송') == -1:
                        delivery_price = gift_class.text.replace(',', '')
                        delivery_price = delivery_price.replace('원', '')
                        delivery_price = int(delivery_price)
                    else:
                        delivery_price = 0
                    
                    write_log('아이텝 %d번 가격 :' % row_number)
                    write_log('price : %d' % item_price)
                    
                    write_log('delivery price : %d' % delivery_price)
                    
                    write_log('link : %s' % item_url)
                    ws.Cells(row_number, Cur_price_column).Value = item_price
                    ws.Cells(row_number, Delivery_price_column).Value = delivery_price
                    ws.Cells(row_number, Cur_URL_column).Value = item_url
                    
                    if ws.Cells(row_number, Old_price_column).Value != item_price:
                        diff.append((row_number, item_name, int(ws.Cells(row_number, Old_price_column).Value), item_price))
                except Exception as e:
                    try:
                        write_log(e)
                    finally:
                        e = None
                        del e

            finally:
                write_log('최저가 가격 획득')

            driver.close()
            window_num = 0
            
            driver.switch_to.window(driver.window_handles[window_num])
            
        else: ### 가격비교 가능한 아이템이 없는 경우 ###
            ws.Cells(row_number, 7).Value = ''
            ws.Cells(row_number, 9).Value = ''
            ws.Cells(row_number, 33).Value = ''
        
        row_number += 1
 
driver.quit()
write_log('엑셀 파일에 데이터를 저장 후 종료 합니다.')

if diff:
    write_log('변동된 가격 정보를 텔레그램으로 전송 합니다.')
    bot = telepot.Bot('1146194999:AAED43PhvHMme3ibW80Fnlgq9XiIXqvugHI')
    msg = ''
    for info in diff:
        msg += '- %s.%s\n%s => %s\n' % info

    print(msg)
    bot.sendMessage('714653402', msg)

wb.Save()
excel.Quit()
date = str(datetime.date.today())
New_Excel_PATH = PATH + File_Name + '_' + date + File_extension
shutil.copy(Excel_PATH, New_Excel_PATH)
