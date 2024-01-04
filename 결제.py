import requests
import json
from bs4 import BeautifulSoup as bs
from openpyxl import Workbook
from datetime import datetime
import time

all_datas = []

page = 1
while page <= 27:
    page += 1
    f = open(f'coupang/coupang{page}.html')

    soup = bs(open(f"coupang/coupang{page}.html"), "html.parser")
    text = soup.select_one("#__NEXT_DATA__").get_text()
    json_object = json.loads(text)
    shopping_list = json_object['props']['pageProps']['domains']['desktopOrder']['orderList']

    datas = []

    for i in range(len(shopping_list)):
        title = shopping_list[i]['title']
        for j in shopping_list[i]['deliveryGroupList']:
            productName_lists = shopping_list[i]['deliveryGroupList'][0]
            productName_list = productName_lists['productList'][0]
            productName = productName_list['vendorItemName']
            quantity_lists = shopping_list[i]['deliveryGroupList'][0]
            quantity_list = quantity_lists['productList'][0]
            quantity = quantity_list['quantity']
            detail = str(productName) + " : " + str(quantity) + "개"
            print(detail)
        price = shopping_list[i]['totalProductPrice']
        date = str(shopping_list[i]['orderedAt'])[:10]
        date_time = datetime.fromtimestamp(int(date)).strftime('%Y-%m-%d %H:%M:%S')
        address = shopping_list[i]['deliveryDestination']['addressDetail']
        datas.append([title, detail, price, date_time, address])
    # print(datas)
    all_datas.append(datas)
print(all_datas)


write_wb = Workbook()
write_ws = write_wb.create_sheet('내역')
write_ws.append(['상품명','세부내역 및 수량', '가격', '구매일', '지점'])

for datas in all_datas:
    for item in datas:
        write_ws.append(item)

write_wb.save('./쿠팡 결제 내역(12월-1월).xlsx')
f.close()
