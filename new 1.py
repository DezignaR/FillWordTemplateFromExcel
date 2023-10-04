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
tab = pd.DataFrame()


def saveDoc(title):
    doc = DocxTemplate(path_template_docx)
    doc.render(gen_dict)
    doc.save(title + str(gen_dict['titul_number']) + " " + str(gen_dict['subsystem']) + ".docx")


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

    # def setVarTableAlignVrSec(event):
    #     global var_table_align_vr_sec
    #     var_table_align_vr_sec = cmb_column_vr_sec.get()

    def setVarTableSheetHr(event):
        global var_table_sheet_hr
        var_table_sheet_hr = cmb_sheet_hz.get()

    def setVarTableSheetVr(event):
        global var_table_sheet_vr
        var_table_sheet_vr = cmb_sheet_vr.get()
        cmb_column_vr['value'] = getNameCol(var_table_sheet_hr, spin_hr.get())
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
            notebook.select(4)

    def btnContinueVr():
        notebook.select(3)

    def btnAddVrTag():
        getDataFrame(skip=int(spin_vr.get()), sheet=var_table_sheet_vr)
        getTag(tag_vr)
        label_msg['text'] = "Добавлено таблица: %s, связь: %s." % (str(cmb_sheet_vr.get()), str(cmb_column_vr.get()))

        print(tab_vr_dict)

    def btnApply():
        getDataFrame(skip=int(spin_hr.get()), sheet=var_table_sheet_hr)
        for index in range(0, 2):
            getRowTable(index)
            dictMerge()
            saveDoc("ПИ_ПИ_тест")

    root = Tk()
    root.title("Отчётогенератор V1.0")
    root.geometry("500x300")
    root.resizable(False, False)

    notebook = ttk.Notebook()

    enable_vertical = IntVar()

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

    entry_xls = Entry(path_frame)
    btn_xls = Button(path_frame, text='Выбрать')

    entry_doc = Entry(template_frame)
    btn_doc = Button(template_frame, text='Выбрать')

    spin_hr = Spinbox(horizontal_frame, from_=0.0, to=5.0)
    spin_vr = Spinbox(vertical_frame, from_=0.0, to=5.0)

    btn_continue_xl = Button(path_frame, text="Продолжить")
    btn_continue_hz = Button(horizontal_frame, text="Продолжить")
    btn_continue_vr = Button(vertical_frame, text="Продолжить")
    btn_add_vr = Button(vertical_frame, text="Добавить")
    btn_apply = Button(template_frame, text="Выполнить")
    # Упаковщики

    notebook.pack(fill=BOTH, expand=True)

    ##############################################################################################

    # Первая закладка (Выбор файла ХЛ)

    label_xl = addLabel("Выберите файл .Xlsx", path_frame)
    label_xl.pack(fill=X)
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

    addLabel("Выберите файл шаблон", template_frame)
    entry_doc.pack(fill=X)
    btn_doc.pack(anchor='e')
    btn_doc.bind('<Button-1>', openFileDialogDC)

    btn_apply.pack(side="bottom", anchor='e')
    btn_apply['command'] = btnApply

    #########################################################################################

    root.mainloop()


if __name__ == '__main__':
    main()
