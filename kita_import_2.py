from selenium import webdriver
import time
from bs4 import BeautifulSoup
import openpyxl
from selenium.webdriver.support.ui import Select

wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(["국가", "수출금액(천$)", "수입금액(천$)"])

headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'}

# 로그인 및 수입까지 설정
driver = webdriver.Chrome('C:/Users/User/Desktop/chromedriver_win32/chromedriver.exe')
driver.implicitly_wait(2)
driver.get('https://stat.kita.net/')
driver.find_element_by_xpath('//*[@id="header"]/div/div[1]/ul/li[1]/a').click()
driver.implicitly_wait(2)
driver.find_element_by_xpath('//*[@id="userId"]').send_keys('g8880916')
driver.find_element_by_xpath('//*[@id="pwd"]').send_keys('daeco30979')
driver.find_element_by_xpath('//*[@id="loginBtn"]').click()
time.sleep(5)
driver.find_element_by_xpath('//*[@id="gnb"]/ul/li[1]/a').click()
driver.find_element_by_xpath('//*[@id="gnb"]/ul/li[1]/div/div/ul/li[1]/ul/li[3]/a').click()
driver.find_element_by_xpath('//*[@id="contents"]/div[1]/div[1]/ul/li[2]/a').click()
elem = driver.find_element_by_xpath('//*[@id="s_item_value"]')
elem.clear()

for element in ('2804291000', '2804293000', '2804294000'):
    driver.find_element_by_xpath('//*[@id="s_item_value"]').send_keys(element)
    time.sleep(5)
    driver.find_element_by_xpath('//*[@id="contents"]/div[2]/form/fieldset/div[3]').click()
    time.sleep(2)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    dollars = soup.select('#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr')
    for dollar in dollars[2:]:
        country = dollar.select_one('#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr > td:nth-child(5)').text
        exportdollar = dollar.select_one('#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr > td:nth-child(11)').text
        importdollar = dollar.select_one('#mySheet1 > tbody > tr:nth-child(3) > td > div > div.GMPageOne > table > tbody > tr > td:nth-child(13)').text
        print(country, exportdollar, importdollar)
        dollar_data = {'국가': country, '수출금액(천$)': exportdollar, '수입금액(천$)': importdollar}
        sheet.append([country, exportdollar, importdollar])
    print('--------------------------------------------')

    driver.find_element_by_xpath('//*[@id="s_item_value"]').send_keys(u'\ue009' + u'\ue003');
    elem = driver.find_element_by_xpath('//*[@id="s_item_value"]')
    elem.clear()

wb.save("dollar_data.xlsx")

'''weightoption = driver.find_element_by_xpath('//*[@id="contents"]/div[2]/form/fieldset/div[2]/div[4]/select')
selector = Select(weightoption)
selector.select_by_value('당월')
'''