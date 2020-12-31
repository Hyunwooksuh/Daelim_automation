from selenium import webdriver
import time
from bs4 import BeautifulSoup
import openpyxl
wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(["item", "state", "dollar", "weight"])

headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'}

# 로그인 및 수입까지 설정
driver = webdriver.Chrome('C:/Users/User/Desktop/chromedriver_win32/chromedriver.exe')
driver.implicitly_wait(2)
driver.get('https://www.bandtrass.or.kr/login.do?returnPage=M')
driver.find_element_by_xpath('//*[@id="id"]').send_keys('daeco')
driver.find_element_by_xpath('//*[@id="pw"]').send_keys('daeco941001!')
driver.find_element_by_xpath('//*[@id="page-wrapper"]/div/div/div[2]/div/table/tbody/tr[1]/td[2]/button').click()
driver.find_element_by_xpath('//*[@id="pass_change"]/div/div/div[2]/button[2]').click()
driver.find_element_by_xpath('//*[@id="1_search"]/div[2]/div/div[1]/a[2]').click()
driver.find_element_by_xpath('//*[@id="tr1"]/td/div[2]/label').click()
driver.find_element_by_xpath('//*[@id="tr2"]/td[1]/div[2]/label').click()

for element in ('2916124000', '2917121000', '2916111000', '2914110000', '2914221000',
                '2915339000', '2916123000', '2909430000', '2902110000', '2909410000',
                '2915310000', '2916121000', '2905122090', '2916122000', '2905310000',
                '2914120000', '2905110000', '2914130000', '2916141000', '2905320000',
                '2902500000', '2902300000', '2902440000', '2815120000', '2909192000',
                '2928009020', '2915320000', '2929101000', '2909491000', '2920903000',
                '2915393000'):
    # HS CODE 입력하기
    driver.find_element_by_xpath('//*[@id="SelectCd"]').send_keys(element)
    driver.find_element_by_xpath('//*[@id="form"]/div/div[1]/div[3]/button').click()
    time.sleep(20)
    html = driver.page_source # html을 문자열로 가져온다.
            # driver.close() # 크롬드라이버 닫기
    soup = BeautifulSoup(html, 'html.parser')
    chemicals = soup.select('#table_list_1 > tbody > tr')
    for chem in chemicals[1:]:
         item = chem.select_one('#table_list_1 > tbody > tr > td:nth-child(3) > font').text
         state = chem.select_one('#table_list_1 > tbody > tr > td:nth-child(4)').text
         dollar = chem.select_one('#table_list_1 > tbody > tr > td:nth-child(5) > font').text
         weight = chem.select_one('#table_list_1 > tbody > tr > td:nth-child(7) > font').text
         if len(weight) >= 9:
             import_data = {'item': item, 'state':state, 'dollar':dollar, 'weight':weight}
             sheet.append([item, state, dollar, weight])

    driver.find_element_by_xpath('//*[@id="SelectCd"]').send_keys(u'\ue009' + u'\ue003');
    elem = driver.find_element_by_xpath('//*[@id="SelectCd"]')
    elem.clear()

wb.save("importdata.xlsx")

