import openpyxl

from form.PayInfoForm import generate_pay_info_form

if __name__ == '__main__':
    data_book = openpyxl.load_workbook('./data/payinfos.xlsx', data_only=True)
    data_sheet = data_book.active

    pay_service_company_column = 'A'
    pay_year_column = 'B'
    actual_pay_column = 'C'

    for r in range(1, data_sheet.max_row + 1):
        service_company = data_sheet[pay_service_company_column + str(r)].value
        pay_year = data_sheet[pay_year_column + str(r)].value
        actual_pay = data_sheet[actual_pay_column + str(r)].value
        print(service_company + ' ' + str(pay_year) + '年度 支付金额: ' + str(actual_pay))
        pay_info = {'apply_pay': str(actual_pay), 'actual_pay': str(actual_pay)}
        generate_pay_info_form(company=service_company, year=str(pay_year), pay_info=pay_info)
