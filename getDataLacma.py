import requests
from lxml import html
from bs4 import BeautifulSoup
import xlsxwriter
import regex
import pandas as pd
import unicodedata
import string
from datetime import date
from time import strptime

def month_converter(month):
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    return months.index(month) + 1

def callme(path,row,col,deep):
    workbook = xlsxwriter.Workbook('UserDatabase.xlsx')
#    worksheet.set_column(0, 2, 30)
    worksheet.set_column(26, 26, 150)
    worksheet.set_column(6, 6, 50)
    response2 = requests.get(path)
    parsed_body2 = html.fromstring(response2.content)
    title = parsed_body2.xpath('//h1[@class="title gutter"]//text()')
    str1 = ''.join(title)
    print str1.replace("', u'", "")
    description = parsed_body2.xpath('//div[@class="field-item even"]//p//text()')
    description = [item.replace(u'\xa0', "") for item in description]
    description = [item.replace(u'\u2019s', "") for item in description]
    description = [item.replace(u'\u201c', "") for item in description]
    description = [item.replace(u'\u201d', "") for item in description]
    description = [item.replace(u'\u2014', "") for item in description]
    description = [item.replace(u'\xa0', "") for item in description]
    description = [item.replace(u'\xe9', "") for item in description]
    description = [item.replace(u'\u2013', "") for item in description]
    description = [item.replace(u"\u2018", "")for item in description]
    description = [item.replace(u"\u2019", "'") for item in description]
    str2 = ''.join(description)
    str2.replace("', u'", "")
    print str2
    dateE = parsed_body2.xpath('//div[@class="field field-type-date field-name-field-exhibition-date"]/div/text()')
    str3 = ''.join(dateE)
    print str3.replace("', u'", "")
#    row += 1
    worksheet.write_string(row,col,"1683")
    worksheet.write_string(row, col+1, "Museum Associates dba Los Angeles County Museum of Art")
    worksheet.write_string(row, col + 27, deep[row-1])
    worksheet.write_string(row, col + 6, str1)
    worksheet.write_string(row, col + 26, str2)
    test = deep[row-1]
    name = test.encode('utf8', 'replace')
    test2 = name.split()
    abd = str(test2[2])
    if "Undetermined" in abd:
        s_year2 = int(abd[0:4])
        s_day2 = int(test2[1].replace(',', ''))
        s_month2 = test2[0]
        month_num3 = month_converter(s_month2)
        start_date2 = str(month_num3) + '/' + str(s_day2) + '/' + str(s_year2)
        end_date2="NA"
        worksheet.write_string(row, col + 2, start_date2)
        worksheet.write_string(row, col + 3, end_date2)
        worksheet.write_string(row, col + 4, "NA")
        worksheet.write_string(row, col + 5, str(s_year2))
    else:
        s_year = int(abd[0:4])
        s_day = int(test2[1].replace(',', ''))
        s_month = test2[0]
        month_num1 = month_converter(s_month)
        start_date = str(month_num1)+'/'+str(s_day)+'/'+str(s_year)
        if len(test2) > 3:
            e_year = int(test2[4])
            e_day = int(test2[3].replace(',', ''))
            e_month = (abd[7:len(abd)])
            month_num2 = month_converter(e_month)
            d0 = date(int(s_year), int(month_num1), int(s_day))
            d1 = date(int(e_year), int(month_num2), int(e_day))
            delta = d1 - d0
            end_date = str(month_num2)+'/'+str(e_day)+'/'+str(e_year)
            worksheet.write_string(row, col + 4, str(delta.days))
            worksheet.write_string(row, col + 3, end_date)
        else:
            worksheet.write_string(row, col + 4, "NA")
            worksheet.write_string(row, col + 3, "NA")
        worksheet.write_string(row, col +2, start_date)
        worksheet.write_string(row, col + 5, str(s_year))
    workbook.close()
#    print '\n'
    return

row = 0
col = 0
workbook = xlsxwriter.Workbook('LACMA_DATA_TILL_2013.xlsx')
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format({'bold': True, 'italic': True})
green_format = workbook.add_format()
green_format.set_pattern(1)
green_format.set_bg_color('#008000')
green_format.set_bold()
green_format.set_italic()
worksheet.write(row, col, 'orgid',green_format)
worksheet.write(row, col+1, 'orgname',green_format)
worksheet.write(row, col+2, 'startdate',green_format)
worksheet.write(row, col+3, 'enddate',green_format)
worksheet.write(row, col+4, 'exhlength',green_format)
worksheet.write(row, col+5, 'year',green_format)
worksheet.write(row, col+6, 'exhtitle',green_format)
worksheet.write(row, col+7, 'singleormultiple',green_format)
worksheet.write(row, col+8, 'artistnumber',green_format)
worksheet.write(row, col+9, 'artistnum',green_format)
worksheet.write(row, col+10, 'artistnum_n',green_format)
worksheet.write(row, col+11, 'featuredartist',green_format)
worksheet.write(row, col+12, 'otherpurpose',green_format)
worksheet.write(row, col+13, 'lecture',green_format)
worksheet.write(row, col+14, 'publication',green_format)
worksheet.write(row, col+15, 'publicationtype',green_format)
worksheet.write(row, col+16, 'numofcurator',green_format)
worksheet.write(row, col+17, 'numofassistants',green_format)
worksheet.write(row, col+18, 'individual',green_format)
worksheet.write(row, col+19, 'foundation',green_format)
worksheet.write(row, col+20, 'corporate',green_format)
worksheet.write(row, col+21, 'government',green_format)
worksheet.write(row, col+22, 'travelout',green_format)
worksheet.write(row, col+23, 'travelin',green_format)
worksheet.write(row, col+24, 'numofworks',green_format)
worksheet.write(row, col+25, 'note',green_format)
worksheet.write(row, col+26, 'description',green_format)
worksheet.write(row, col+27, 'EXHIBITION DATE',green_format)
worksheet.set_row(0, 20)
deep = []
app = []

for i in range(0, 7):
    abc = "http://www.lacma.org/art/exhibitions/past?page="
    dab = abc+str(i)
    response2 = requests.get(dab)
    doc=response2.content
    soup = BeautifulSoup(''.join(doc))
    href = soup.find_all("div", class_="view-content")
    test = str(href)
    soup2 = BeautifulSoup(''.join(test))

    for a in soup.find_all("span", class_="date-display-start"):
        deep.append(a.get_text())
    for a in soup2.find_all('a', href=True):
        row += 1
        callme('http://www.lacma.org/'+a['href'],row,col,deep)
workbook.close()