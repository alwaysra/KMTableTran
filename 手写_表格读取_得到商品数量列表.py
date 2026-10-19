import numbers
import  os,re
import openpyxl
from bs4 import BeautifulSoup

with open("2.xls", "r", encoding="utf-8") as f:
    html_content = f.read()
#得到商品和数量列表
# 使用BeautifulSoup解析html文本，html.parser是python自带解析器不用额外装
soup = BeautifulSoup(html_content, "html.parser")
# 获取tbody下面所有的<tr>表格行
tr_list = soup.find_all("tr")
j=0
for tr in tr_list:
    tds=tr.find_all("td")
    sp_qty_list = []
    for td in tds[3:]:
        #print(td.get_text(strip=True))
        if j<len(tr_list)-2:
            sp_qty_list.append(td.get_text(strip=True))
    wb = openpyxl.load_workbook("1_1.xlsx")
    sheet = wb['Sheet1']
    i=0
    while i<len(sp_qty_list):
        celltext=sp_qty_list[i]
        #使用正则表达式只需要整数部分
        if j==0:
            sheet.cell(i+2,j+1).value=celltext
            print(celltext)
            i = i + 1
        if j>0:
            celltext=celltext.replace("瓶","")  #替换单位汉字
            sheet.cell(i+2,j+2).value=celltext #从C列开始插入第一列是商品,第二列是小计
            i=i+1
        ## 指定不同的文件名，可以另存为别的文件
    wb.save("1_1.xlsx")
    j=j+1

hj=tr_list[len(tr_list)-2]
hj_list=[]
for td in hj:
    tds=hj.find_all("td")
    hj_list.append(td.get_text(strip=True))


wb = openpyxl.load_workbook("1_1.xlsx")
sheet = wb['Sheet1']
i=0
while i<len(hj_list):
    celltext=hj_list[i]
    #使用正则表达式只需要整数部分
    celltext=celltext.replace("瓶","")  #替换单位汉字
    sheet.cell(i+1,2).value=celltext #第二列B 是合计
    i=i+1
    ## 指定不同的文件名，可以另存为别的文件
wb.save("1_1.xlsx")

kh_list=[] #客户列表
tr_list = soup.find_all("tr")
#合计行 和  凑整行不要 减去两行
for tr_i in range(len(tr_list)-2):
    if tr_i==0:
        continue
    td_list=tr_list[tr_i].find_all("td")
    kh_list.append(td_list[1].get_text(strip=True))

wb = openpyxl.load_workbook("1_1.xlsx")
sheet = wb['Sheet1']
i=0
while i<len(kh_list):
   print(kh_list[i])
   sheet.cell(1,i+3).value=kh_list[i]
   i=i+1

## 指定不同的文件名，可以另存为别的文件
wb.save("1_1.xlsx")
print("客户处理完成")