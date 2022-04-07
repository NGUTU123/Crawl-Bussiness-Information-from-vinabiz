from asyncore import write
import logging
import sys
import argparse
import urllib.parse
import requests
from bs4 import BeautifulSoup
import xlwt 
from xlwt import Workbook 
import xlrd
from xlutils.copy import copy
from xlrd import open_workbook

from company import Company

cookie = 'ASP.NET_SessionId=veab5j3herzozn4lsriy4vll; __RequestVerificationToken=Z3TrM6ogea5owXgxvQ_GWoeXPna9eabJTRdTjZEQjHaVn2n_rd4-vcsmRGpmQai0e37iLTnhjd4Qg4INgsAfPMXPv8XV-NpzrWUufWoY8Yk1; .AspNet.ApplicationCookie=velUXQ80tPCP117PQFhEbZY37rsu9JxoIlMqpLqod0Mp3t5YDmFqEiPnCFJVofFKszplUGf6ccHdlrvmuj3V3Jul3oAYWlFQD0yBkBh926lNQT3E8ngP-XF19TgSX6iOMyqDDZiYbzK32a29wmYl5zbtHct9fjzYySzodf0reiE6YpzpFI4x0lpUaFlIyw_NDzuqefQug7tndNRYe7wRk1Nn17j9jusjwTl3ttCXoJvSLRKARJpD4qmh3AOUFF4TrObvZvs1I_7QXgpEuSQ63RT_z9TuPjxHEh7Cyjayv5GzwkdkGmMmMybnO0SUNmUYXSCkC1cA9FLIqW218ARFtKFTq4j6TCg1WtjlaBagHmRshmAU-71iNGeccPmHX887699po0hnYLrJSldVFjgHWhWyw-ZhnWUYlXs-ffhYjE2ELNnl6JpLtS2ameuw3RorDIcucZmXXK1wE4-82zLJvVQtJvwV8sShtQrD6Izah0M4nGLxYUY4FnXucVkeFItQHxce09i93Yzz0tSaQxeBFw'
log_format = '\x1b[0m[%(levelname)s] - %(message)s'
logging.basicConfig(level='INFO', format=log_format)
write_index = 0

parser = argparse.ArgumentParser()
parser.add_argument("--url", "-u", help="base url")
parser.add_argument("--start", "-s", help="start page")
parser.add_argument("--end", "-e", help="end page")
parser.add_argument("--out", "-o", help="output file")
args = parser.parse_args()

company_arr = []


def check_input():
    if args.url is None:
        logging.error('Please enter base url');
        sys.exit(0)
    if args.start is None:
        logging.error('Please enter start page');
    if int(args.start) <= 0:
        logging.error('Please enter start page > 0');
        sys.exit(0)
    if args.end is None:
        logging.error('Please enter end page');
        sys.exit(0)
    if int(args.start) > int(args.end):
        logging.error('Please enter start page < end page');
        sys.exit(0)
    if args.out is None:
        logging.error('Please enter output file');
        sys.exit(0)


def r(e, t):
    r = e[t:t+2]
    return int(r, base=16)


def decode(n, c):
    o = ''
    a = r(n, c)
    i = c + 2
    xs = i
    for x in range(i, len(n)):
        if xs in range(i, len(n)):
            l = r(n, xs) ^ a
            o += chr(l)
            xs = xs + 2
        else:
            break
    try:
        o = urllib.parse.unquote(urllib.parse.quote(o))
        return o
    except Exception as e:
        logging.error(str(e))


def request_list_company(page):
    company_url_list = []
    logging.info("getting list of company in page " + str(page))
    url = args.url
    if url[:-1] != '/' : url = url + '/'
    if int(page) > 1 : url = url + str(page)
    response = requests.get(url, headers={'Cookie': cookie})
    soup = BeautifulSoup(response.content, 'html.parser')
    list_of_company_div = soup.find_all("div", class_= "row margin-right-15 margin-left-10")
    for company_div in list_of_company_div:
        if company_div.find('a')['href'] : company_url_list.append(company_div.find('a')['href'])
    logging.info('Get total ' + str(len(company_url_list)) + ' company url')
    return company_url_list


def parse_company_detail(rows):
    emailCode = None
    company = Company()
    company.official_name = rows[1].find_all('td')[1].get_text().strip()
    company.trading_name = rows[1].find_all('td')[3].get_text().strip()
    company.bussiness_code = rows[2].find_all('td')[1].get_text().strip()
    company.date_of_license = rows[2].find_all('td')[3].get_text().strip()
    company.start_working_date = rows[3].find_all('td')[3].get_text().strip()
    company.status = rows[4].find_all('td')[1].find_all('div', class_='alert alert-success fade in')[0].get_text().strip()
    company.address = rows[7].find_all('td')[1].get_text().strip()
    company.phone = rows[8].find_all('td')[1].get_text().strip()
    if rows[9].find_all('td')[1].find('span', class_='__cf_email__'): emailCode = rows[9].find_all('td')[1].find('span', class_='__cf_email__')['data-cfemail']
    if emailCode is not None: company.email = decode(emailCode, 0)
    else: company.email = ''
    #company.author = rows[10].find_all('td')[1].get_text().strip()
    company.director = rows[12].find_all('td')[1].get_text().strip()
    company.director_phone = rows[12].find_all('td')[3].get_text().strip()
    company.accountant = rows[14].find_all('td')[1].get_text().strip()
    company.accountant_phone = rows[14].find_all('td')[3].get_text().strip()
    return company


def get_company_details(url):
    url = 'https://vinabiz.us' + url
    logging.info('\x1b[1;32mGet company details in ' + url + '\x1b[0m')
    response = requests.get(url, headers=
    {
        'cookie': cookie, 
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-language': 'en-US,en;q=0.9,vi;q=0.8',
        'content-type': 'text/html; charset=utf-8',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.39'
    })

    soup = BeautifulSoup(response.content, 'html.parser')
    rows = soup.find_all("table", class_= "table-bordered")[0].find_all('tr')
    
    company = parse_company_detail(rows)
    if soup.find("div", {"id": "hr2"}) is not None:
         company.business_lines = soup.find("div", {"id": "hr2"}).get_text()
    else:
         company.business_lines = ''
    
    company_arr.append(company)


def write_sheet_header(sheet):
    sheet_header = ['Tên chính thức', 'Tên giao dịch', 'Mã doanh nghiệp', 'Ngày cấp', 'Ngày bắt đầu hoạt động', 
    'Trạng thái', 'Địa chỉ', 'Điện thoại', 'Email', 'Giám đốc', 'SĐT giám đốc', 
    'Kế toán', 'SĐT kế toán', 'Nghành nghề']
    #sheet_header = ['Tên chính thức', 'Địa chỉ', 'Điện thoại', 'Người đại diện', 'Giám đốc']

    for header in sheet_header:
        sheet.write(0, sheet_header.index(header), header)


def write_sheet_data(data):
    global write_index
    file = args.out + '.xls'
    rb = open_workbook(file, formatting_info=True)
    wb_c = copy(rb)

    #r_sheet = rb.sheet_by_index(0) #original
    w_sheet = wb_c.get_sheet(0)

    for company in data:
        attributes_arr = list(company.__dict__.keys())
        for att in attributes_arr:
            w_sheet.write(write_index, attributes_arr.index(att), str(getattr(company, att)))
        write_index += 1
    wb_c.save(file)
    write_index += 1
    logging.info('Saved file ' + file)


#def write_result(data):
#    file = args.out + '.xls'
#    logging.info('Save result to file')
#    wb = Workbook()
#    sheet = wb.add_sheet('Data')
#    write_sheet_header(sheet)
#    write_sheet_data(sheet, data)
#    wb.save(file)
#    logging.info('Saved to ' + file)

def creat_result_file():
    global write_index
    file = args.out + '.xls'
    wb = Workbook()
    sheet = wb.add_sheet(args.out)
    write_sheet_header(sheet)
    wb.save(file)
    write_index = 1
    logging.info('Created file ' + file)


def craw():
    total = 0
    creat_result_file()
    company_arr.clear()
    for i in range(int(args.start), int(args.end) + 1):
        company_url_list = request_list_company(i)
        for company_url in company_url_list:
            try:
                get_company_details(company_url)
            except Exception as e:
                logging.error('url: ' + company_url)
                logging.error(str(e))
        write_sheet_data(company_arr)
        total += len(company_arr)
        company_arr.clear()
    logging.info('Get information of total ' + str(len(company_arr)) + ' companies')


def main():
      check_input()
      craw()


if __name__== "__main__":
  main()

