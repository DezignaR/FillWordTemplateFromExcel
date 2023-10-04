import pandas as pd
from docxtpl import DocxTemplate
from tkinter import *
from tkinter import ttk
from tkinter import filedialog

path_template_docx = ""
path_base_xlsx = ""
tab_vr_dict = {}
gen_dict = {}
gen_skip_row = 0
sec_skip_row = 0
tag_vr = ""
var_table_align_vr = []
var_table_sheet_vr = ""
var_table_sheet_hr = ""
title_name = ""
first_name = ""
sec_name = ""
name_doc_file = ""

tab = pd.DataFrame()


def saveDoc():
    doc = DocxTemplate(path_template_docx)
    doc.render(gen_dict)
    doc.save("%s %s %s.docx" % (title_name, str(gen_dict[first_name]), str(gen_dict[sec_name])))


def getDataFrame(sheet, skip):
    global tab
    tab = pd.read_excel(io=path_base_xlsx, sheet_name=sheet, skiprows=skip, engine='openpyxl')


def getTag(tag_vr_lc):
    global tab
    global tab_vr_dict
    tag = tab.to_dict()
    tag = list(tag[tag_vr_lc].values())
    for col_name, data in tab.items():
        tab_vr_dict[col_name] = dict(zip(tag, list(data.dropna(axis=0).to_dict().values())))
    del tab_vr_dict[tag_vr_lc]


def dictMerge():
    for index in var_table_align_vr:
        gen_dict.update(tab_vr_dict[gen_dict[index]])


def getRowTable(index):
    global gen_dict
    gen_dict = tab.iloc[index].to_dict()
    # print(genDict)


def getNameCol(sheet, skip, flag=False):
    if flag:
        getDataFrame(skip=int(skip), sheet=sheet)
        return str(tab.columns[0])
    else:
        getDataFrame(skip=int(skip), sheet=sheet)
        return tuple(tab.columns.values.tolist())


def getSheets():
    return tuple(pd.ExcelFile(path_base_xlsx).sheet_names)


def main():
    def setVarTableAlignVr(event):
        global var_table_align_vr
        var_table_align_vr.append(cmb_column_vr.get())

    def setVarTableSheetHr(event):
        global var_table_sheet_hr
        var_table_sheet_hr = cmb_sheet_hz.get()

    def setVarTableSheetVr(event):
        global var_table_sheet_vr
        var_table_sheet_vr = cmb_sheet_vr.get()

        global tag_vr
        tag_vr = getNameCol(var_table_sheet_vr, spin_vr.get(), True)

    def openFileDialogXL(event):
        global path_base_xlsx
        path_base_xlsx = filedialog.Open(root).show()
        if path_base_xlsx == "":
            return
        entry_xls.insert(1, path_base_xlsx)
        sheet_list = getSheets()
        cmb_sheet_vr['values'] = sheet_list
        cmb_sheet_hz['values'] = sheet_list

    def openFileDialogDC(event):
        global path_template_docx
        path_template_docx = filedialog.Open(root).show()
        if path_template_docx == "":
            return
        entry_doc.insert(1, path_base_xlsx)

    def addLabel(text_label, child):
        label = Label(child)
        label['text'] = text_label
        return label

    def addCombobox(child):
        return ttk.Combobox(child, state='readonly')

    def addNotebook(frame, text, state="normal"):
        notebook.add(frame, text=text, state=state)

    def addCheckbox(child, text):
        checkbox = Checkbutton(child, text=text)
        return checkbox

    def addEntry(child):
        return Entry(child)

    def checkboxChange(note, check):
        if check.get() == 1:
            notebook.tab(note, state="normal")
            return
        elif check.get() == 0:
            notebook.tab(note, state="disabled")
            return

    def btnContinueXl():
        notebook.select(1)

    def btnContinueHr():
        if enable_vertical.get() == 1:
            notebook.select(2)
        else:
            notebook.select(3)

        name_col = getNameCol(var_table_sheet_hr, spin_hr.get())
        cmb_column_vr['value'] = name_col
        cmb_name_first['values'] = name_col
        cmb_name_sec['values'] = name_col

    def btnContinueVr():
        notebook.select(3)

    def btnAddVrTag():
        getDataFrame(skip=int(spin_vr.get()), sheet=var_table_sheet_vr)
        getTag(tag_vr)
        label_msg['text'] = "Добавлено таблица: %s, связь: %s." % (str(cmb_sheet_vr.get()), str(cmb_column_vr.get()))

        print(tab_vr_dict)

    def btnApply():
        getDataFrame(skip=int(spin_hr.get()), sheet=var_table_sheet_hr)
        count = len(tab.index)
        for index in range(0, count):
            nonlocal prog_bar_var
            prog_bar_var.set(index * 100 / count)
            getRowTable(index)
            dictMerge()
            saveDoc()

    def fillName(event):
        msg = 0
        entry_name_full.delete(0, END)
        entry_name_full.insert(0, "%s %s %s.docx" % (entry_name.get(), cmb_name_first.get(), cmb_name_sec.get()))
        global title_name, first_name, sec_name
        title_name = entry_name.get()
        first_name = cmb_name_first.get()
        sec_name = cmb_name_sec.get()

    root = Tk()
    root.title("Отчётогенератор V1.0")
    root.geometry("500x320")
    root.resizable(False, False)

    notebook = ttk.Notebook()

    enable_vertical = IntVar()
    prog_bar_var = DoubleVar()

    # Области размещения элементов
    path_frame = Frame(notebook, bg='gray')
    horizontal_frame = Frame(notebook, bg='gray')
    vertical_frame = Frame(notebook, bg='gray')
    template_frame = Frame(notebook, bg='gray')

    addNotebook(path_frame, text="Xslx")
    addNotebook(horizontal_frame, "Hr ЛИСТ 1")
    addNotebook(vertical_frame, text="Vr ЛИСТ 1", state="disabled")
    addNotebook(template_frame, text="Docx")

    # Элементы формы - объявление
    spin_hr = Spinbox(horizontal_frame, from_=0.0, to=5.0)
    spin_vr = Spinbox(vertical_frame, from_=0.0, to=5.0)

    btn_xls = Button(path_frame, text='Выбрать')
    btn_doc = Button(template_frame, text='Выбрать')
    btn_continue_xl = Button(path_frame, text="Продолжить")
    btn_continue_hz = Button(horizontal_frame, text="Продолжить")
    btn_continue_vr = Button(vertical_frame, text="Продолжить")
    btn_add_vr = Button(vertical_frame, text="Добавить")
    btn_apply = Button(template_frame, text="Выполнить")

    prog_bar = ttk.Progressbar(template_frame, orient='horizontal', length=100, variable=prog_bar_var,
                               mode='determinate')
    # Упаковщики

    notebook.pack(fill=BOTH, expand=True)

    ##############################################################################################

    # Первая закладка (Выбор файла ХЛ)

    label_xl = addLabel("Выберите файл .Xlsx", path_frame)
    label_xl.pack(fill=X)
    entry_xls = addEntry(path_frame)
    entry_xls.pack(fill=X)

    btn_xls.pack(anchor='e')
    btn_xls.bind('<Button-1>', openFileDialogXL)

    btn_continue_xl['command'] = btnContinueXl
    btn_continue_xl.pack(side="bottom", anchor='e')

    # Вторая закладка (Горизонтальная главная)

    label_sheet_hz = addLabel("Выберите лист таблицы с данными по горизонтали", horizontal_frame)
    label_sheet_hz.pack(fill=X)
    cmb_sheet_hz = addCombobox(horizontal_frame)
    cmb_sheet_hz.pack(fill=X)
    cmb_sheet_hz.bind('<<ComboboxSelected>>', setVarTableSheetHr)

    label_spin_hr = addLabel("Укажите количество строк для пропуска", horizontal_frame)
    label_spin_hr.pack(fill=X)
    spin_hr.pack(anchor='w')

    check_align = addCheckbox(horizontal_frame, text="Добавить вертикальную таблицы")
    check_align['variable'] = enable_vertical
    check_align['command'] = lambda: checkboxChange(2, enable_vertical)
    check_align.pack(fill=X)

    btn_continue_hz['command'] = btnContinueHr
    btn_continue_hz.pack(side="bottom", anchor='e')

    # Третья закладка (Вертикальная)

    label_sheet_vr_cmb = addLabel("Выберите лист таблицы с данными по вертикале", vertical_frame)
    label_sheet_vr_cmb.pack(fill=X)
    cmb_sheet_vr = addCombobox(vertical_frame)
    cmb_sheet_vr.pack(fill=X)
    cmb_sheet_vr.bind('<<ComboboxSelected>>', setVarTableSheetVr)

    label_column_vr_cmb = addLabel("Выберите параметр для слияния таблиц", vertical_frame)
    label_column_vr_cmb.pack(fill=X)
    cmb_column_vr = addCombobox(vertical_frame)
    cmb_column_vr.pack(fill=X)
    cmb_column_vr.bind('<<ComboboxSelected>>', setVarTableAlignVr)

    label_spin_vr = addLabel("Укажите количество строк для пропуска", vertical_frame)
    label_spin_vr.pack(fill=X)
    spin_vr.pack(anchor='w')

    btn_add_vr['command'] = btnAddVrTag
    btn_add_vr.pack(anchor='e')

    label_msg = addLabel("", vertical_frame)
    label_msg.pack(fill=X)

    btn_continue_vr['command'] = btnContinueVr
    btn_continue_vr.pack(side="bottom", anchor='e')

    # Четвертая (Выбор шаблона)

    label_dox = addLabel("Выберите файл шаблон", template_frame)
    label_dox.pack(fill=X)
    entry_doc = addEntry(template_frame)
    entry_doc.pack(fill=X)
    btn_doc.pack(anchor='e')
    btn_doc.bind('<Button-1>', openFileDialogDC)

    label_name_title = addLabel("Выберите заголовок файла", template_frame)
    label_name_title.pack(fill=X)
    entry_name = addEntry(template_frame)
    entry_name.pack(anchor="w")

    label_name_title = addLabel("Выберите первый тег", template_frame)
    label_name_title.pack(fill=X)
    cmb_name_first = addCombobox(template_frame)
    cmb_name_first.pack(anchor="w")
    cmb_name_first.bind('<<ComboboxSelected>>', fillName)

    label_name_title = addLabel("Выберите второй тег", template_frame)
    label_name_title.pack(fill=X)
    cmb_name_sec = addCombobox(template_frame)
    cmb_name_sec.pack(anchor="w")
    cmb_name_sec.bind('<<ComboboxSelected>>', fillName)

    label_name_title = addLabel("Маска имени файла", template_frame)
    label_name_title.pack(fill=X)
    entry_name_full = addEntry(template_frame)
    entry_name_full.pack(fill=X)

    btn_apply.pack(side="bottom", anchor='e')
    btn_apply['command'] = btnApply

    prog_bar.pack(side="bottom", anchor='w')

    #########################################################################################

    root.mainloop()


if __name__ == '__main__':
    main()
