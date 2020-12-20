import os

from docx import Document


def save_project_form_file(company: str, year: str, project_id: str, project_name: str, form_name: str,
                           doc: Document):
    file_name_suffix = project_id + '_' + project_name
    if not os.path.exists('./生成表格/' + company):
        os.mkdir('./生成表格/' + company)
    if not os.path.exists('./生成表格/' + company + '/' + year):
        os.mkdir('./生成表格/' + company + '/' + year)
    if not os.path.exists('./生成表格/' + company + '/' + year + '/' + file_name_suffix):
        os.mkdir('./生成表格/' + company + '/' + year + '/' + file_name_suffix)
    doc.save('./生成表格/' + company + '/' + year + '/' + file_name_suffix + '/' + form_name + '.docx')


def save_company_form_file(company: str, year: str, form_name: str, doc: Document):
    if not os.path.exists('./生成表格/' + company):
        os.mkdir('./生成表格/' + company)
    if year is None:
        doc.save('./生成表格/' + company + '/' + form_name + '.docx')
    else:
        if not os.path.exists('./生成表格/' + company + '/' + year):
            os.mkdir('./生成表格/' + company + '/' + year)
        doc.save('./生成表格/' + company + '/' + year + '/' + form_name + '.docx')


def save_special_company_form_file(company: str, year: str, form_name: str, doc: Document):
    if not os.path.exists('./生成表格/'):
        os.mkdir('./生成表格/')
    if not os.path.exists('./生成表格/special'):
        os.mkdir('./生成表格/special')
    if not os.path.exists('./生成表格/special/' + company):
        os.mkdir('./生成表格/special/' + company)
    if year is None:
        doc.save('./生成表格/special/' + company + '/' + form_name + '.docx')
    else:
        if not os.path.exists('./生成表格/special/' + company + '/' + year):
            os.mkdir('./生成表格/special/' + company + '/' + year)
        doc.save('./生成表格/special/' + company + '/' + year + '/' + form_name + '.docx')


def save_company_exists_form_file(company: str, year: str, form_name: str, doc: Document):
    if not os.path.exists('./生成表格/' + company):
        return
    if year is None:
        return
    else:
        if not os.path.exists('./生成表格/' + company + '/' + year):
            os.mkdir('./生成表格/' + company + '/' + year)
        doc.save('./生成表格/' + company + '/' + year + '/' + form_name + '.docx')
