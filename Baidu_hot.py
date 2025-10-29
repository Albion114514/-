import requests
from bs4 import BeautifulSoup
import openpyxl

# 发起HTTP请求获取百度热搜页面内容
url = 'https://top.baidu.com/board?tab=realtime'
response = requests.get(url)
html = response.content

# 使用BeautifulSoup解析页面内容
soup = BeautifulSoup(html, 'html.parser')

# 提取热搜数据
hot_searches = []
for item in soup.find_all('div', {'class': 'c-single-text-ellipsis'}):
    hot_searches.append(item.text)

# 保存热搜数据到Excel
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = 'Baidu Hot Searches'

# 设置标题
sheet.cell(row=1, column=1, value='百度热搜排行榜')

# 写入热搜数据
for i in range(len(hot_searches)):
    sheet.cell(row=i + 2, column=1, value=hot_searches[i])

workbook.save('百度热搜.xlsx')
print('热搜数据已保存到 百度热搜.xlsx')