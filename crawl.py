import requests
from bs4 import BeautifulSoup
import json
import xlrd
import sys


'''
searchIndustry(industry, data)
This function takes name and json data as arguments.
It returns the index of name in data. otherwise -1
'''


def searchIndustry(industry, data):
    for i in data:
        if industry.lower() == i['industry_name'].lower():
            return data.index(i)

    return -1


'''
this function fetches the url and parses it to collect
the industry name, no of companies, url of industry, perm_id, heirarchical id, an array of all the companies
. In the array of all the companies, it fetches the ticker, name of company, market capitalization, ttm sales, number of employees, and the details of each employee in a separate array.
'''


def parse_industries(url, input_url, sheet_data):
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    # search for the heading. assume data does not exist if heading is not
    # there.
    try:
        heading = soup.find(id="sectionTitle").text.strip()
    except:
        return []

    no_of_companies = int(soup.find(id="pageEnd").text.strip())

    # the outermost array for storing the industry data object
    data = []

    # The object to store the industry data
    content = {}
    content['industry'] = heading

    # searching for perm_id and heirarchical_id in sheet data
    index = searchIndustry(heading, sheet_data)
    if index >= 0:
        content['perm_id'] = sheet_data[index]['perm_id']
        content['heirarchical_id'] = sheet_data[index]['heirarchical_id']

    content['url'] = input_url
    content['no_of_companies'] = no_of_companies
    content['companies'] = []

    # Start fetching company data
    table_content = soup.select("table#dataTable tr")
    table_content.pop(0)
    table_content.pop(0)

    for table_row in table_content:
        cells = table_row.findAll('td')
        company_detail = {}
        company_detail['ticker'] = cells[0].text.strip()
        company_detail['name'] = cells[1].text.strip()

        company_detail['market_cap'] = cells[2].text.strip()

        company_detail['ttm_sales'] = cells[3].text.strip()

        # convert 23,512 to 23512
        try:
            no_of_employees = int(cells[4].text.strip().replace(',', ''))
        except:
            no_of_employees = -1

        company_detail['no_of_employees'] = no_of_employees
        company_detail['url'] = "https://www.reuters.com" + \
            cells[1].find('a')['href']

        # print("https://www.reuters.com/finance/stocks/company-officers/" +
        #       cells[0].text.strip())
        company_detail['executives'] = getExecutiveData(
            url="https://www.reuters.com/finance/stocks/company-officers/" + cells[0].text.strip())

        content['companies'].append(company_detail)

    data.append(content)
    return data

'''
this function will get the data from excel sheet and aggregate data of industry name and its perm_id and heirarchical_id
'''


def parse_sheet():
    workbook = xlrd.open_workbook(
        'C-TR-Business-Classification-Index-1.xlsx')

    # Selecting sheet 'TRBC'
    worksheet = workbook.sheet_by_name('TRBC')

    industry_name = []
    perm_id = []
    heirarchical_id = []

    for name in worksheet.col_values(3):
        if name:
            industry_name.append(name.strip())

    # Removing first row of the sheet
    industry_name.pop(0)

    for id in worksheet.col_values(4):
        if id:
            try:
                perm_id.append(int(id))
            except:
                pass

    for id in worksheet.col_values(5):
        if id:
            try:
                heirarchical_id.append(int(id))
            except:
                pass

    data = []
    for i in range(len(industry_name)):
        details = {}
        details['industry_name'] = industry_name[i]
        details['perm_id'] = perm_id[i]
        details['heirarchical_id'] = heirarchical_id[i]
        data.append(details)

    return data

'''
This function will search for name for data. Return index if name exists.
'''


def search_name(name, data):
    # Remove non breaking spaces from the string
    new_name = name.text.strip().replace("\xa0", " ")
    for i in data:
        name1 = i['name'].replace("\u00A0", " ")
        if name1 == new_name:
            return data.index(i)

    return -1


'''
This function will take the company executive's url and fetches the details of all the executive including name, age, since, position and description
'''


def getExecutiveData(url):
    data = []
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    # print(url)
    table_content = soup.findAll(class_="dataTable")

    # check if table exist in the page
    try:
        table1 = table_content[0]
    except:
        return data

    # check if description table exist in the page
    flag = 0
    try:
        table2 = table_content[1]
        flag = 1
    except:
        pass

    cells = table1.findAll('td')
    for name, age, since, position in zip(*[iter(cells)] * 4):
        executive_details = {}

        executive_details['name'] = name.text.strip().replace("\u00A0", " ")
        executive_details['age'] = age.text.strip()
        executive_details['since'] = since.text.strip()
        executive_details['current_pos'] = position.text.strip()

        try:
            executive_url = "https://www.reuters.com" + name.find('a')['href']
        except:
            executive_url = ""
        executive_details['url'] = executive_url
        executive_details['description'] = ""
        data.append(executive_details)

    # if flag = 0, then description table does not exist
    if flag:
        cells = table2.findAll('td')
        for name, desc in zip(*[iter(cells)] * 2):
            index = search_name(name, data)
            if index >= 0:
                data[index]['description'] = desc.text.strip()

    return data


'''
Check for command line arguments, if no argument then runs script with default url. Otherwise uses the user supplied url
'''


def main():
    if len(sys.argv) == 1:
        input_url = "https://www.reuters.com/sectors/industries/rankings?industryCode=179"
    elif len(sys.argv) == 2:
        input_url = sys.argv[1]
    else:
        print("Enter only one url")
        exit(1)

    url = input_url + "&view=size&page=-1&sortby=mktcap&sortdir=DESC"

    print("URL is: %s" % input_url)
    sheet_data = parse_sheet()
    data = parse_industries(url, input_url, sheet_data)

    # This code snippet can be used to fetch the content from all the industries with range from 1 to 200, if the page exists.

    # for i in range(1, 200):
    #     url = "https://www.reuters.com/sectors/industries/rankings?industryCode=" + \
    #         str(i) + "&view=size&page=-1&sortby=mktcap&sortdir=DESC"
    #     data = parse_industries(url, input_url, sheet_data)

    #     with open('data' + str(i) + '.txt', 'w') as outfile:
    #         json.dump(data, outfile, indent=2, ensure_ascii=False)

    with open('data.txt', 'w') as outfile:
        json.dump(data, outfile, indent=2, ensure_ascii=False)

'''
Start the execution here
'''

if __name__ == '__main__':
    main()
