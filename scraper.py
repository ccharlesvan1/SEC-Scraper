import requests
import re
import urllib
import lxml
import xlsxwriter
from bs4 import BeautifulSoup
from datetime import datetime

def make_url(base, components):
    url = base
    for part in components:
        if part[0] == "/":
            url += part
        else:
            url += "/" + part
    return url

def standardize_cik(cik):
    integer = int(cik)
    return str(integer)

def quarter_number(month):
    if month <= 3:
        quarter_num = 1
    elif month <= 6:
        quarter_num = 2
    elif month <= 9:
        quarter_num = 3
    else:
        quarter_num = 4
    return quarter_num

print("Please enter company ticker:")
ticker = input()

#finding company CIK number

document_type = "10-K"
base_url = r"https://www.sec.gov"
url = base_url + "/cgi-bin/browse-edgar?action=getcompany&CIK=" + ticker + "&type=10-K&dateb=&owner=exclude&count=10"
s = BeautifulSoup(requests.get(url).content, 'xml')
for span in s.find("div", {"id" : 'contentDiv'}).find_all('span'):
    if span['class'] == 'companyName':
        name_and_cik = span.text.split("CIK#:")
company_name = str.strip(name_and_cik[0])
company_cik = standardize_cik(name_and_cik[1].split("(")[0])

#finding 10-K filing date, year, and quarter

regex = "([0-9]{4}-[0-9]{2}-[0-9]{2})"
date = s.table.find('td', text=re.compile(regex)).text
filing_date = date.replace("-", "")
datetime_obj = datetime.strptime(date, '%Y-%m-%d')
year = str(datetime_obj.year)
quarter = "QTR" + str(quarter_number(datetime_obj.month))

#from SEC's master archive, find company's filing summary document

master_file = "master." + filing_date + ".idx"
url = make_url(base_url, ["Archives/edgar/daily-index", year, quarter, master_file])
content = requests.get(url).content
splitted = content.decode('utf-8').split('--------------------------------------------------------------------------------')
data = splitted[1].replace("\n", "|").split("|")
txt_link = ""
for index, elem in enumerate(data):
    if (elem == company_cik and data[index + 2] == document_type):
        txt_link = data[index + 4]
        break
fixed_link = txt_link.replace("-","").replace(".txt", "")
updated_url = make_url(base_url, ["Archives", fixed_link, "index.json"])
content = requests.get(updated_url).json()
for file in content['directory']['item']:
    if file['name'] == 'FilingSummary.xml':
        summary_url = make_url(base_url, [content['directory']['name'], file['name']])
        break
new_base_url = summary_url.replace("FilingSummary.xml", "")
content = requests.get(summary_url).content

#getting all the reports from the summary document

soup_obj = BeautifulSoup(content, "lxml")
reports = soup_obj.find('myreports')
all_reports = []
for report in reports.find_all('report')[:-1]:
    if str.lower(report.menucategory.text) == 'statements':
        report_dict = {}
        report_dict['name'] = report.shortname.text
        report_dict['url'] = new_base_url + report.htmlfilename.text
        all_reports.append(report_dict)

#search for the three main statements amongst all reports

statements_url = []
balance_sheet = ["consolidated balance sheets", "consolidated balance sheet"]
cash_flow = ["consolidated statements of cash flows", "consolidated statement of cash flows"]
income = ['operations', 'income', 'comprehensive income']
statements = balance_sheet + cash_flow

for report_dict in all_reports:
    lower_name = str.lower(report_dict['name'])
    if lower_name in statements:
        statements_url.append(report_dict['url'])

for report_dict in all_reports:
    lower_name = str.lower(report_dict['name'])
    if income[0] in lower_name:
        statements_url.append(report_dict['url'])
        break
    elif income[1] in lower_name and not 'comprehensive' in lower_name:
        statements_url.append(report_dict['url'])
        break
    elif income[2] in lower_name:
        statements_url.append(report_dict['url'])
        break

#parse and store data for each financial statement

statement_data_dict = {}
for index, statement in enumerate(statements_url):
    line_items = {}
    line_items['header'] = []
    line_items['line'] = []
    content = requests.get(statement).content
    statement_soup = BeautifulSoup(content, 'lxml')
    for row in statement_soup.table.find_all('tr'):
        cols = row.find_all('td')
        if len(row.find_all('th')) != 0:
            row_list = [elem.text.strip() for elem in row.find_all('th')]
            line_items['header'].append(row_list)
        else:
            row_list = [elem.text.strip().replace('$',"").replace(',',"").replace("\n", "").replace("(", "-").replace(")","") for elem in cols]
            line_items['line'].append(row_list)
    statement_data_dict[index] = line_items

#creating and writing to an Excel workbook

wb = xlsxwriter.Workbook(company_name + '.xlsx', {'strings_to_numbers': True})
for statement_num in statement_data_dict:
    statement_data = statement_data_dict[statement_num]
    sheet = wb.add_worksheet('Sheet ' + str(statement_num))
    for line_type in statement_data:
        line_items = statement_data[line_type]
        if line_type == "header":
            title = line_items[0][0]
            if len(line_items) == 1:
                year_ends = line_items[0][1:]
            else:
                year_ends = line_items[1]
            sheet.write(0, 0, title, wb.add_format({'bold': True}))
            for index, year in enumerate(year_ends):
                sheet.write(1, index+ 1, year)
        else:
            for col_index, line in enumerate(line_items):
                for row_index, cell_content in enumerate(line):
                    sheet.write(col_index + 2, row_index, cell_content)
wb.close()
