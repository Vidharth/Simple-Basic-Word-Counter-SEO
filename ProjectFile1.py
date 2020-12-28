from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter
import sqlite3
import json

try:
   url1=input('Enter the URL:')
   urlresp1=urlopen(url1)

except HTTPError as e:
   print("Check URL or network settings.")

finally:
   data1=urlresp1.read()
   urlresp1.close()
   print("URL Response Received Successfully!")

keyword=input("Enter the keywords:")
keyList=keyword.split()

soup1=BeautifulSoup(data1,'html.parser')

for x in soup1(['script', 'style']):
   x.extract()

text1=soup1.get_text()
list1=text1.split()

newDict1={}

for a in list1:
   for b in keyList:
      if a == b :
         if a not in newDict1:
            newDict1[a]=1
         else:
            newDict1[a]+=1

connection = sqlite3.connect("SEOFreq.db")
connection.execute("CREATE TABLE IF NOT EXISTS SEOFreq (Word TEXT, Count INT, Frequency REAL)")
print("Database Connection Established!")
print("Table Created Successfully!")

newDict2={}

for x,y in newDict1.items():
   if x not in newDict2:
      newDict2[x]=y/len(list1)
      connection.execute("INSERT INTO SEOFreq (Word,Count,Frequency) VALUES (?,?,?)" , (x, y,newDict2[x]))

connection.commit()
connection.close()

print("Values Inserted Into Database Successfully!")
   
workbook = xlsxwriter.Workbook("SEOFreq.xlsx")
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format({'bold': True})
cell_format.set_text_wrap()
cell_format.set_font_size(19)
worksheet.write(0,0 , 'Keywords',cell_format)
worksheet.write(0, 1, 'Word Count',cell_format)
worksheet.write(0, 2, 'Word Frequency',cell_format)

cell_format1 = workbook.add_format()
cell_format.set_text_wrap()
cell_format1.set_font_size(15)

chart = workbook.add_chart({'type':'bar'})
chart1 = workbook.add_chart({'type':'bar'})

row = 1
col = 0
for word in newDict1.keys():
   worksheet.write(row, col, word,cell_format1)
   row +=1

def two():
   row = 1
   col = 1 
   for value in newDict1.values():
      worksheet.write(row, col, value,cell_format1)
      row +=1
   return row

def three():
   row = 1
   col = 2 
   for value in newDict2.values():
      worksheet.write(row, col, value,cell_format1)
      row +=1
   return row

two()
chart.add_series({
       'categories': ['Sheet1', 1, 0, row, 0],
       'values':     ['Sheet1', 1, 1, row, 1],
       'line':       {'color': 'red'},
       'name': ['Sheet1', 0, 1]
   })
chart.set_x_axis({'name': 'Word Count'})
chart.set_y_axis({'name': 'KeyWord'})

three()
chart1.add_series({
    'categories': ['Sheet1', 1, 0, row, 0],
    'values':     ['Sheet1', 1, 2, row, 2],
    'line':       {'color': 'red'},
    'name': ['Sheet1', 0, 2]
})
chart1.set_x_axis({'name': 'Frequency'})
chart1.set_y_axis({'name': 'KeyWord'})

with open('SEOWC.json', 'w') as fp:
    json.dump(newDict1, fp)

with open('SEOFreq.json', 'w') as fp:
    json.dump(newDict2, fp)

print("Json File Created and Values Updated!")

worksheet.insert_chart('M2', chart)
worksheet.insert_chart('M19', chart1)
workbook.close()

print("Excel Created and Values Mapped and Graphed!")
print("Please View the Data in the Respective Files and DB")
