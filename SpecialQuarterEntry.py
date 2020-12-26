import time

import openpyxl
from form.SpecialQuarterConfirmForm \
    import generate_special_quarter_confirm_form


def add_quarter_data(quarter: str, service: str, confirm_infos: dict):
    if quarter not in confirm_infos:
        confirm_infos[quarter] = {}
    if service not in confirm_infos[quarter]:
        confirm_infos[quarter][service] = [0, 0, '']


if __name__ == '__main__':
    data_book = openpyxl.load_workbook(
        './data/special_projects.xlsx', data_only=True)
    data_sheet = data_book.active

    year_column = 'A'
    project_id_column = 'AI'
    project_name_column = 'AJ'
    service_company_column = 'B'
    service_price_1_column = 'V'
    service_price_2_column = 'X'
    unit_type_column = 'N'
    service_content_1_column = 'Q'
    service_type_1_column = 'P'

    service_total_1_column = 'M'
    service_amount_1_column = 'Z'

    quarter_1_amount_1_column = 'BT'
    quarter_1_total_1_column = 'BW'

    quarter_2_amount_1_column = 'DC'
    quarter_2_total_1_column = 'DF'

    quarter_3_amount_1_column = 'EL'
    quarter_3_total_1_column = 'EO'

    quarter_4_amount_1_column = 'FV'
    quarter_4_total_1_column = 'FY'

    company_quarter_confirm = {}
    for r in range(1, data_sheet.max_row + 1):
        start = time.time()
        year = data_sheet[year_column + str(r)].value
        service_company = data_sheet[service_company_column + str(r)].value
        project_id = data_sheet[project_id_column + str(r)].value
        project_name = data_sheet[project_name_column + str(r)].value
        service_amount_1 = data_sheet[service_amount_1_column + str(r)].value
        service_total_1 = data_sheet[service_total_1_column + str(r)].value
        service_price_1 = data_sheet[service_price_1_column + str(r)].value
        unit_type = data_sheet[unit_type_column + str(r)].value
        service_content_1 = data_sheet[service_content_1_column + str(r)].value
        service_type_1 = data_sheet[service_type_1_column + str(r)].value

        quarter_1_amount_1 = data_sheet[
            quarter_1_amount_1_column + str(r)
        ].value
        quarter_2_amount_1 = data_sheet[
            quarter_2_amount_1_column + str(r)
        ].value
        quarter_3_amount_1 = data_sheet[
            quarter_3_amount_1_column + str(r)
        ].value
        quarter_4_amount_1 = data_sheet[
            quarter_4_amount_1_column + str(r)
        ].value

        quarter_1_total_1 = data_sheet[quarter_1_total_1_column + str(r)].value
        quarter_2_total_1 = data_sheet[quarter_2_total_1_column + str(r)].value
        quarter_3_total_1 = data_sheet[quarter_3_total_1_column + str(r)].value
        quarter_4_total_1 = data_sheet[quarter_4_total_1_column + str(r)].value

        start_quarter = time.time()
        company_quarter_key = service_company + '+' + str(year)
        if project_name is None:
            project_name = ''
        project_key = str(project_id) + '+' + project_name
        # 季度确认
        if company_quarter_key not in company_quarter_confirm:
            company_quarter_confirm[company_quarter_key] = {}
        confirm_info = company_quarter_confirm[company_quarter_key]

        if quarter_1_amount_1 != 0 and quarter_1_amount_1 is not None:
            add_quarter_data('1', service_type_1, confirm_info)
            confirm_info['1'][service_type_1][0] += quarter_1_amount_1
            confirm_info['1'][service_type_1][1] += quarter_1_total_1
            confirm_info['1'][service_type_1][2] = unit_type

        if quarter_2_amount_1 != 0 and quarter_2_amount_1 is not None:
            add_quarter_data('2', service_type_1, confirm_info)
            confirm_info['2'][service_type_1][0] += quarter_2_amount_1
            confirm_info['2'][service_type_1][1] += quarter_2_total_1
            confirm_info['2'][service_type_1][2] = unit_type

        if quarter_3_amount_1 != 0 and quarter_3_amount_1 is not None:
            add_quarter_data('3', service_type_1, confirm_info)
            confirm_info['3'][service_type_1][0] += quarter_3_amount_1
            confirm_info['3'][service_type_1][1] += quarter_3_total_1
            confirm_info['3'][service_type_1][2] = unit_type

        if quarter_4_amount_1 != 0 and quarter_4_amount_1 is not None:
            add_quarter_data('4', service_type_1, confirm_info)
            confirm_info['4'][service_type_1][0] += quarter_4_amount_1
            confirm_info['4'][service_type_1][1] += quarter_4_total_1
            confirm_info['4'][service_type_1][2] = unit_type

    print(company_quarter_confirm)
    for (key, confirm_info) in company_quarter_confirm.items():
        company_name = key.split('+', 1)[0]
        year = key.split('+', 1)[1]
        for (q, confirm) in company_quarter_confirm[key].items():
            generate_special_quarter_confirm_form(company_name, year,
                                                  q, confirm)
    pass
