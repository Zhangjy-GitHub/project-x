import openpyxl

from form.CompanyInfoForm import generate_company_info_form

if __name__ == '__main__':
    data_book = openpyxl.load_workbook('./data/companies.xlsx', data_only=True)
    data_sheet = data_book.active

    name_column = 'A'
    business_scope_column = 'C'
    legal_person_column = 'D'
    founded_time_column = 'E'
    register_fund_column = 'F'
    employee_quantity_column = 'G'
    register_address_column = 'H'
    bank_name_column = 'I'
    bank_account_column = 'J'

    for r in range(1, data_sheet.max_row + 1):
        company_info = {'company': data_sheet[name_column + str(r)].value,
                        'person':
                            data_sheet[legal_person_column + str(r)].value,
                        'register_fund':
                            data_sheet[register_fund_column + str(r)].value,
                        'founded_time':
                            data_sheet[founded_time_column + str(r)].value,
                        'employee_total':
                            data_sheet[employee_quantity_column + str(r)].value,
                        'business_scope':
                            data_sheet[business_scope_column + str(r)].value,
                        'register_address':
                            data_sheet[register_address_column + str(r)].value,
                        'bank_name':
                            data_sheet[bank_name_column + str(r)].value,
                        'bank_account':
                            data_sheet[bank_account_column + str(r)].value}
        print(company_info['company'])
        generate_company_info_form(companies_info=company_info)

    pass
