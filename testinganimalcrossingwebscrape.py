import requests, re, csv, os
from bs4 import BeautifulSoup
from selenium import webdriver

driver = webdriver.Chrome(executable_path= r'C:\Users\cphil\PycharmProjects\chromedriver.exe')
urlpage= r'https://animalcrossing.gamepress.gg/furniture'
driver.get(urlpage)

items = driver.find_elements_by_css_selector('div[class=view-furniture] div')
all_data = []
for item in items:
    data = str.split(item.text, '\n')
    try:
        if 'craft' in str.lower(data[2]):
            try:
                all_data.append([data[0], data[1], data[2]])
            except:
                pass
    except:
        pass
with open(os.path.join(os.curdir, 'animalcrossing.csv'), 'w', newline = '') as my_csv:
    csv.writer(my_csv).writerows(all_data)


