# -*- coding: windows-1251 -*-
from docxtpl import DocxTemplate
import win32com.client

context = {}
subcontext = {}

BASE = 'База'
METOD = 'Этапы'
ART = 'ПИ '
sheet_excel = {}
sheet_ = {}


def up_context(tag, value):
    context[tag] = value


def save_doc():
    doc = DocxTemplate(ART+"шаблон.docx")
    doc.render(context)
    doc.save(ART + str(context['titul_number']) + " " + str(context['subsystem']) + ".docx")


def get_tag(sheet_v, diap_t, revers):
    if sheet_v == METOD:
        range_tag = 1
    else:
        range_tag = 2
    tag = excel_get_data(range_tag, diap_t, sheet_v, revers)
    return tag


def get_value(rang, range_val, sheet_v, revers):
    val = excel_get_data(rang, range_val, sheet_v, revers)
    return val


def get_len(sheet_v):
    if sheet_ != sheet_v:
        excel_connection(sheet_v)

    if sheet_ == METOD:
        leng_h = 2
        leng_v = 2
    else:
        leng_h = 1
        leng_v = 1
    vals = 0
    while vals is not None:
        vals = sheet_excel.Cells(1, leng_h).value
        leng_h = leng_h + 1

    vals = 0
    while vals is not None:
        vals = sheet_excel.Cells(leng_v, 1).value
        leng_v = leng_v + 1
    leng = [leng_h - 2, leng_v - 2]
    return leng


def excel_connection(sheet_v):
    global sheet_
    sheet_ = sheet_v
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(u'C:\\Users\\krasr\\PycharmProjects\\pythonProject\\base_v1.xlsx')
    global sheet_excel
    sheet_excel = wb.WorkSheets(sheet_)


def excel_get_data(diap, rang_m, sheet_v, revers):
    if sheet_ != sheet_v:
        excel_connection(sheet_v)

    if sheet_v == METOD and revers:
        vals = [r[0].value for r in sheet_excel.Range(sheet_excel.Cells(2, diap), sheet_excel.Cells(rang_m, diap))]
    elif sheet_v == METOD and not revers:
        vals = [r[0].value for r in sheet_excel.Range(sheet_excel.Cells(diap, 1), sheet_excel.Cells(diap, rang_m))]
    else:
        for i in range(diap, diap + 1):
            vals = [r[0].value for r in sheet_excel.Range("A" + str(i) + ":T" + str(i))]

    return vals


def main():
    # Добавление методик
    leng_m = get_len(METOD)

    tag_subsystem = get_tag(METOD, leng_m[0], False)
    tag_metod = get_tag(METOD, leng_m[1], True)

    for i in range(2, leng_m[0] + 1):
        value = get_value(i, leng_m[1], METOD, True)
        k = 0
        context_metod = {}

        while k < len(value):
            if value[k] is not None:
                context_metod[tag_metod[k]] = value[k]
                k = k + 1
            else:
                break
        subcontext[tag_subsystem[i - 1]] = context_metod
    print(subcontext['RT1'])
    # Количество элеменотов
    leng_b = get_len(BASE)

    # Добавление основной информации в словарь
    tag_b = get_tag(BASE, leng_b[0], True)
    for j in range(3, leng_b[1]):
        value = get_value(j, i, BASE, True)
        for i in range(1, leng_b[0] + 1):
            up_context(str(tag_b[i - 1]), str(value[i - 1]))
        subsystem = subcontext[context['subsystem']]
        for i in range(0, len(subsystem)):
            up_context(tag_metod[i], subsystem[tag_metod[i]])
        save_doc()


if __name__ == '__main__':
    main()
