import time

import openpyxl

from form.QuarterConfirmForm import generate_quarter_confirm_form


def add_quarter(quarter: str, project_key_info: str, confirm_infos: dict):
    if quarter not in confirm_infos:
        confirm_infos[quarter] = {}
    confirm_infos[quarter][str(project_key_info)] = []


if __name__ == '__main__':
    data_book = openpyxl.load_workbook(
        './data/projects.xlsx', data_only=True)
    data_sheet = data_book.active

    year_column = 'A'
    project_id_column = 'AI'
    project_name_column = 'AJ'
    service_company_column = 'B'
    service_price_1_column = 'V'
    service_price_2_column = 'X'
    unit_type_column = 'N'
    service_content_1_column = 'Q'
    service_content_2_column = 'Q'

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
    company_year_confirm = {}
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

        quarter_1_amount_1 = data_sheet[quarter_1_amount_1_column +
                                        str(r)].value
        quarter_2_amount_1 = data_sheet[quarter_2_amount_1_column +
                                        str(r)].value
        quarter_3_amount_1 = data_sheet[quarter_3_amount_1_column +
                                        str(r)].value
        quarter_4_amount_1 = data_sheet[quarter_4_amount_1_column +
                                        str(r)].value

        quarter_1_total_1 = data_sheet[quarter_1_total_1_column + str(r)].value
        quarter_2_total_1 = data_sheet[quarter_2_total_1_column + str(r)].value
        quarter_3_total_1 = data_sheet[quarter_3_total_1_column + str(r)].value
        quarter_4_total_1 = data_sheet[quarter_4_total_1_column + str(r)].value

        print('项目号: ' + str(project_id))
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
            add_quarter('1', project_key, confirm_info)
            confirm_info['1'][project_key]\
                .append((str(quarter_1_amount_1) + unit_type,
                        str(service_price_1), quarter_1_total_1,
                        service_content_1))

        if quarter_2_amount_1 != 0 and quarter_2_amount_1 is not None:
            add_quarter('2', project_key, confirm_info)
            confirm_info['2'][project_key]\
                .append((str(quarter_2_amount_1) + unit_type,
                        str(service_price_1), quarter_2_total_1,
                        service_content_1))

        if quarter_3_amount_1 != 0 and quarter_3_amount_1 is not None:
            add_quarter('3', project_key, confirm_info)
            confirm_info['3'][project_key]\
                .append((str(quarter_3_amount_1) + unit_type,
                         str(service_price_1), quarter_3_total_1,
                         service_content_1))

        if quarter_4_amount_1 != 0 and quarter_4_amount_1 is not None:
            add_quarter('4', project_key, confirm_info)
            confirm_info['4'][project_key]\
                .append((str(quarter_4_amount_1) + unit_type,
                         str(service_price_1), quarter_4_total_1,
                         service_content_1))

        start_year = time.time()
        # 年度确认
        if service_company not in company_year_confirm:
            company_year_confirm[service_company] = {}
        if str(year) not in company_year_confirm[service_company]:
            company_year_confirm[service_company][str(year)] = {}
        company_year_info = company_year_confirm[service_company][str(year)]
        if service_total_1 is not None and service_total_1 != 0:
            if project_key not in company_year_info:
                company_year_info[project_key] = []
            company_year_info[project_key]\
                .append((str(service_amount_1) + unit_type,
                         str(service_price_1), service_total_1,
                         service_content_1))
        print('读取时间: ' + str(start_quarter - start) + ' 季度处理时间: ' +
              str(start_year - start_quarter) +
              ' 年度处理时间: ' + str(time.time() - start_year))

    for (c, company_info) in company_quarter_confirm.items():
        company_year = c.split('+', 1)
        for (q, project_info) in company_info.items():
            generate_quarter_confirm_form(
                company_year[0], company_year[1], q, project_info)

    """for (c, company_info) in company_year_confirm.items():
        for (y, project_info) in company_info.items():
            generate_year_confirm_form(c, y, project_info)"""
