from email.policy import default

from docxtpl import DocxTemplate
import win32com.client
import PySimpleGUI as sg

context = {}
subcontext = {}
pathExcel = ""
pathTemplate = ""
countSheet = 0
# sheet1 = ''
revers_ = ''
reversSW = ''
# sheet2 = ''
# sheet3 = ''
sheet_excel = {}
sheet_ = {}


def add_context(tag, value):
    context[tag] = value


# def save_doc():
#     doc = DocxTemplate(pathTemplate)  # "ЖПР шаблон.docx")
#     doc.render(context)
#     doc.save("ЖПР " + str(context['titul_number']) + " " + str(context['subsystem']) + ".docx")


def get_tag(range_table):
    if revers_:
        range_tag = 1
    else:
        range_tag = 2
    tag = excel_get_data(range_tag, range_table)
    return tag


def get_value(range_, range_val):
    val = excel_get_data(range_, range_val)
    return val


def get_len(revers):
    excel_connection()

    if revers:
        length_h = 2
        length_v = 2
    else:
        length_h = 1
        length_v = 1
    vals = 0
    while vals is not None:
        vals = sheet_excel.Cells(1, length_h).value
        length_h = length_h + 1
    vals = 0
    while vals is not None:
        vals = sheet_excel.Cells(length_v, 1).value
        length_v = length_v + 1
    length = [length_h - 2, length_v - 2]
    return length


def excel_connection():
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(pathExcel)  # u'C:\\Users\\krasr\\PycharmProjects\\pythonProject\\base_v1.xlsx')
    global sheet_excel
    sheet_excel = wb.WorkSheets(sheet_)


def excel_get_data(range_, rangMethod):
    if not reversSW:
        vals = [r[0].value for r in
                sheet_excel.Range(sheet_excel.Cells(2, range_), sheet_excel.Cells(rangMethod, range_))]
    elif reversSW:
        vals = [r[0].value for r in
                sheet_excel.Range(sheet_excel.Cells(range_, 1), sheet_excel.Cells(range_, rangMethod))]
    else:
        for i in range_(range_, range_ + 1):
            vals = [r[0].value for r in sheet_excel.Range("A" + str(i) + ":T" + str(i))]

    return vals


def main():
    layout = [
        [sg.Text('Количество Листов в базе для обработки(MAX 3)'), sg.InputText()],
        [sg.Submit('Выбрать'), sg.Cancel('Выйти')]
    ]

    window = sg.Window('File Compare', layout)
    event, values = window.read()
    window.close()
    global countSheet
    countSheet = int(values[0])
    while True:
        match countSheet:
            case 1:
                screen = [
                    [sg.Text('Выберите файл БАЗЫ:'), sg.InputText(), sg.FileBrowse()],
                    [sg.Text('Выберите файл шаблона:'), sg.InputText(), sg.FileBrowse()],
                    [sg.Text('Название Листа 1(Основная база):'), sg.InputText(), sg.Checkbox('Обработка по столбцам'),
                     sg.Checkbox('Один столбец')],
                    [sg.Output(size=(88, 20))],
                    [sg.Submit('Выполнить'), sg.Cancel('Выйти')]
                ]
                break
            case 2:
                screen = [
                    [sg.Text('Выберите файл БАЗЫ:'), sg.InputText(), sg.FileBrowse()],
                    [sg.Text('Выберите файл шаблона:'), sg.InputText(), sg.FileBrowse()],
                    [sg.Text('Название Листа 1(Доп. база 1):'), sg.InputText(), sg.Checkbox('Обработка по столбцам'),
                     sg.Checkbox('Один столбец')],
                    [sg.Text('Название Листа2 (Основная база):'), sg.InputText(), sg.Checkbox('Обработка по столбцам'),
                     sg.Checkbox('Один столбец')],
                    [sg.Output(size=(88, 20))],
                    [sg.Submit('Выполнить'), sg.Cancel('Выйти')]
                ]
                break
            case 3:
                screen = [
                    [sg.Text('Выберите файл БАЗЫ:'), sg.InputText(), sg.FileBrowse()],
                    [sg.Text('Выберите файл шаблона:'), sg.InputText(), sg.FileBrowse()],
                    [sg.Text('Название Листа 1 (Доп. база 1):'), sg.InputText(), sg.Checkbox('Обработка по столбцам'),
                     sg.Checkbox('Один столбец')],
                    [sg.Text('Название Листа 2 (Доп. база 2):'), sg.InputText(), sg.Checkbox('Обработка по столбцам'),
                     sg.Checkbox('Один столбец')],
                    [sg.Text('Название Листа 3 (Основная база):'), sg.InputText(), sg.Checkbox('Обработка по столбцам'),
                     sg.Checkbox('Один столбец')],
                    [sg.Output(size=(88, 20))],
                    [sg.Submit('Выполнить'), sg.Cancel('Выйти')]
                ]
                break
            case _:
                if countSheet > 3:
                    sg.popup("Слишком много листов")
                break

    if countSheet < 4:
        window = sg.Window('File Compare', screen)

        while True:  # The Event Loop
            event, values = window.read()
            if event == 'Выполнить':
                for i in range(2, len(values), 3):
                    global pathExcel
                    global pathTemplate
                    global sheet_
                    global revers_
                    pathExcel = values[0]
                    pathTemplate = values[1]
                    sheet_ = values[i]
                    revers_ = values[i + 1]
                    oneRaw = values[i + 2]

                    if sheet_ != "" and revers_:
                        # теги углом
                        length_m = get_len(revers_)
                        global reversSW
                        reversSW = revers_
                        tag_subsystem = get_tag(length_m[0])
                        reversSW = not revers_
                        tag_method = get_tag(length_m[1])

                    for j in range(2, length_m[0]):
                        value_raw = get_value(j, length_m[1])
                        k = 0
                        context_method = {}

                        while k < len(value_raw):
                            if value_raw[k] is not None:
                                context_method[tag_method[k]] = value_raw[k]
                                k = k + 1
                            else:
                                break
                        subcontext[tag_subsystem[i - 1]] = context_method

                    # Добавление информации в словарь
                    if sheet_ != "" and not revers_:
                        length_b = get_len(revers_)

                        tag_b = get_tag(length_b[0])
                        for j in range(3, length_b[1]):
                            value_raw = get_value(j, i)
                            for x in range(1, length_b[0] + 1):
                                add_context(str(tag_b[x - 1]), str(value_raw[x - 1]))
                            subsystem = subcontext[context['subsystem']]
                            for x in range(0, len(subsystem)):
                                add_context(tag_method[x], subsystem[tag_method[x]])
                # save_doc()

            if event in (None, 'Exit', 'Cancel'):
                break


if __name__ == '__main__':
    main()
