import openpyxl
from form.TContract import generate_contract

if __name__ == '__main__':
    data_book = openpyxl.load_workbook('./data/contracts.xlsx',
                                       data_only=True)
    data_sheet = data_book.active

    company_name_column = 'A'
    contract_year_column = 'B'
    company_area_column = 'C'
    legal_person_column = 'D'
    comm_address_column = 'E'
    contract_area_column = 'F'

    for r in range(1, data_sheet.max_row + 1):
        company_name = data_sheet[company_name_column + str(r)].value
        contract_year = data_sheet[contract_year_column + str(r)].value
        company_area = data_sheet[company_area_column + str(r)].value
        legal_person = data_sheet[legal_person_column + str(r)].value
        comm_address = data_sheet[comm_address_column + str(r)].value
        contract_area = data_sheet[contract_area_column + str(r)].value

        contract_info = {
            'company_name': company_name,
            'company_area': company_area,
            'legal_person': legal_person,
            'comm_address': comm_address,
            'contract_area': contract_area,
            'contract_year': str(contract_year)
        }
        print(contract_info)
        generate_contract(contract_info)
        pass
