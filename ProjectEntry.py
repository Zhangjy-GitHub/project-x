import datetime
from math import fabs

import openpyxl

from form.ApplyBudgeForm import generate_apply_budget_form
from form.ApplyForm import generate_apply_form
from form.ProjectConfirmForm import generate_project_confirm_form
from form.TaskDelegateForm import generate_task_delegate_form

if __name__ == '__main__':

    data_book = openpyxl.load_workbook(
        './data/projects.xlsx', data_only=True)
    data_sheet = data_book.active

    year_column = 'A'
    service_company_column = 'B'
    service_content_1_column = 'Q'
    service_content_2_column = 'Q'
    service_type_column = 'P'
    unit_type_column = 'N'

    project_id_column = 'AI'
    project_name_column = 'AJ'
    contract_time_column = 'AM'
    service_budget_column = 'AB'
    service_price_1_column = 'V'
    service_price_2_column = 'X'
    service_amount_1_column = 'Z'
    service_budget_amount_1_column = 'AC'
    project_area_column = 'AL'
    project_end_column = 'GE'

    original_service_content = ['外场支持协助类',
                                '驻场技术支持类', '设备验货支持类', '项目文件评审类', '专业技术支持类']
    service_budget_predicted = 100000

    for r in range(1, data_sheet.max_row + 1):
        apply_info = {}
        year = data_sheet[year_column + str(r)].value
        project_id = str(data_sheet[project_id_column + str(r)].value)
        project_name = data_sheet[project_name_column + str(r)].value
        service_company = data_sheet[service_company_column +
                                     str(r)].value
        service_type = data_sheet[service_type_column + str(r)].value
        unit_type = data_sheet[unit_type_column + str(r)].value
        print(project_id)
        content = ''
        i = 1
        for sc in original_service_content:
            if sc == service_type:
                content = content + u'\u2713 ' + str(i) + '）' + sc + '\n'
            else:
                content = content + '□ ' + str(i) + '）' + sc + '\n'
            i = i + 1
        apply_info['service_type'] = content
        service_budget = data_sheet[service_budget_column + str(r)].value / 10000
        if service_budget is None or fabs(service_budget) < 1e-6:
            continue
        need_all_files = True
        if service_type == '专业技术支持类' or service_type == '项目文件评审类' and service_budget < service_budget_predicted:
            need_all_files = False
        apply_info['service_budget'] = str(service_budget)
        contract_time_delta = data_sheet[contract_time_column + str(r)].value
        contract_time = datetime.date(1900, 1, 1) + datetime.timedelta(contract_time_delta)

        service_month = contract_time.month
        service_year = year
        if contract_time.year < year:
            service_month = 1
        apply_info['apply_time'] = str(service_year) + '年' + str(service_month) + '月'
        apply_info['service_time'] = str(service_year) + '年' + str(service_month) + '月' \
                                     + ' —— ' + str(year) + '年' + '12月'
        apply_info['apply_dep'] = '工程管理中心'
        if need_all_files:
            generate_apply_form(project_id=project_id, project_name=project_name, company=service_company,
                                year=str(year), apply_info=apply_info)
        service_amount_1 = data_sheet[service_amount_1_column + str(r)].value
        service_budget_amount_1 = data_sheet[service_budget_amount_1_column + str(r)].value
        budget_list = []
        task_list = []
        confirm_list = []
        total_budget = 0
        if service_amount_1 != 0 and service_amount_1 is not None:
            service_price_1 = data_sheet[service_price_1_column + str(r)].value
            service_content_1 = data_sheet[service_content_1_column + str(r)].value
            total = service_price_1 * service_amount_1
            budget = service_price_1 * service_budget_amount_1
            budget_list.append((service_content_1, str(service_budget_amount_1) + unit_type, str(service_price_1), str(budget)))
            task_list.append((service_content_1, str(service_budget_amount_1) + unit_type, service_content_1))
            confirm_list.append(
                (str(service_amount_1) + unit_type, str(service_price_1), str(total), service_content_1))
            total_budget = total_budget + budget
        budget_info = {'budget_list': budget_list, 'total_budget': str(total_budget)}
        if need_all_files:
            generate_apply_budget_form(project_id=project_id, project_name=project_name
                                       , company=service_company, year=str(year), budget_info=budget_info)
        project_area = data_sheet[project_area_column + str(r)].value
        task_info = {'service_company': service_company, 'service_time': apply_info['service_time'],
                     'service_area': project_area, 'task_list': task_list}
        if need_all_files:
            generate_task_delegate_form(project_id=project_id, project_name=project_name
                                        , company=service_company, year=str(year), task_infos=task_info)
        project_end_month = data_sheet[project_end_column + str(r)].value
        if project_end_month is not None and project_end_month != 0:
            confirm_time = str(service_year) + '年' + str(service_month) + '月 —— ' + str(year) + '年' + str(
                project_end_month) + '月'
            confirm_info = {'service_company': service_company,
                            'service_time': confirm_time, 'confirm_list': confirm_list}
            generate_project_confirm_form(project_id=project_id, project_name=project_name
                                          , company=service_company, year=str(year), confirm_info=confirm_info)
