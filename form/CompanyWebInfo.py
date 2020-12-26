import openpyxl
import requests
from lxml import etree


def parse_html(html):
    page = etree.HTML(html)
    person = page.xpath('//h2[@class="seo font-20"]/text()')[0].strip()
    first_row = page.xpath(
        '//section[@id="Cominfo"]/table[@class="ntable"]//tr')[0]
    founded_time = first_row[5].xpath('text()')[0].strip()
    second_row = page.xpath(
        '//section[@id="Cominfo"]/table[@class="ntable"]//tr')[1]
    register_fund = second_row[1].xpath('text()')[0].strip()
    fifth_row = page.xpath(
        '//section[@id="Cominfo"]/table[@class="ntable"]//tr')[5]
    employee_total = fifth_row[1].xpath('text()')[0].strip()
    sixth_row = page.xpath(
        '//section[@id="Cominfo"]/table[@class="ntable"]//tr')[7]
    register_address = sixth_row[1].xpath('text()')[0].strip()
    return person, founded_time, register_fund, employee_total, register_address


if __name__ == '__main__':
    data_book = openpyxl.load_workbook('/data/companies.xlsx', data_only=True)
    data_sheet = data_book.active
    company_column = 'A'
    url_column = 'B'
    business_scope_column = 'C'
    person_column = 'D'
    founded_time_column = 'E'
    register_fund_column = 'F'
    employ_total_column = 'G'
    register_address_column = 'H'
    company_info = {}
    headers = {
        'User-Agent': '''Mozilla/5.0 (X11; Linux x86_64)
                         AppleWebKit/536.5
                         (KHTML, like Gecko)
                         Chrome/19.0.1084.9 Safari/536.5'''}
    error_urls = []
    for r in range(1, data_sheet.max_row + 1):
        service_company = data_sheet[company_column + str(r)].value
        company_info['url'] = data_sheet[url_column + str(r)].value
        print((service_company, company_info['url']))
        response = requests.get(company_info['url'], headers=headers)
        try:
            info = parse_html(html=response.text)
            data_sheet[person_column + str(r)] = info[0]
            data_sheet[founded_time_column + str(r)] = info[1]
            data_sheet[register_fund_column + str(r)] = info[2]
            data_sheet[employ_total_column + str(r)] = info[3]
            data_sheet[register_address_column + str(r)] = info[4]
        except Exception:
            error_urls.append((service_company, company_info['url']))
    data_book.save('C:\\Users\\zhang\\Documents\\Tools\\data\\companies.xlsx')
    print(error_urls)
    pass
