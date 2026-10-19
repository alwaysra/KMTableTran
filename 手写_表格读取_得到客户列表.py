import  os
import openpyxl
from bs4 import BeautifulSoup

with open("1.xls","r",encoding="utf-8") as f:
    html_content=f.read()

#得到客户列表
# 使用BeautifulSoup解析html文本，html.parser是python自带解析器不用额外装
soup = BeautifulSoup(html_content, "html.parser")
# 获取tbody下面所有的<tr>表格行
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

