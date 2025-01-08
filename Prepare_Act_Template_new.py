from tkinter.ttk import Combobox
import numpy as np
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from tkinter import *
from tkinter import filedialog
from docx import Document
import openpyxl
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
from docxtpl import DocxTemplate
from tkinter.messagebox import OK, INFO, showinfo
from tkinter import ttk
from tkinter import font

def open_spec():
    showinfo(title="Информация", message="Выберете спецификацию в формате xlsx")

def make_act():
    showinfo(title="Информация", message="Проверьте что спецификация загружена!")


def safe_act():
    make_act()
    if xl_arr.shape[0] == 1:
        return
    filepath = filedialog.asksaveasfilename()
    if filepath != "" and xl_arr.shape[0]>1:


        headers = ('№ ', 'Поз.', 'Наименование', 'Тип, марка\nматериал\nТехническая\nдокументация',
                   'Завод -\nизготовитель', 'Кол-\nво,\nшт')

        global document

        # Данные для заполнения шаблона
        context = {
            'station': station.get(),
            'calc': calc.get(),
            'company': company.get(),
            'object': obj.get(),
            'address': address.get(),
            'number': number.get(),
            'name': name.get(),
            'data': data.get()
        }

        # Заполнение шаблона данными
        document.render(context)

        all_tables = document.tables
        new_table = all_tables[0]
        cols_number = len(headers)

        if type_choies.get() == 'Акт основного оборудования':

            f_list = np.empty((1, len(xl_arr.tolist()[0]))).tolist()
            f_list.clear()

            f_list += [x for x in xl_arr.tolist() if 'теплообменник' in str(x).lower()]
            f_list += [x for x in xl_arr.tolist() if 'насос' in str(x).lower()]
            f_list += [x for x in xl_arr.tolist() if 'регулирующий' in str(x).lower()]
            f_list += [x for x in xl_arr.tolist() if 'регулятор давления' in str(x).lower()]
            j = 1
            for i in range(0, len(f_list)):
                f_list[i].insert(0, j)
                j += 1
            j = 1

            for row in f_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])
        elif type_choies.get() == 'Акт вспомогательного оборудования':

            s_list = np.empty((1, len(xl_arr.tolist()[0]))).tolist()
            s_list.clear()

            s_list += [x for x in xl_arr.tolist() if 'фильтр' in str(x).lower()]
            s_list += [x for x in xl_arr.tolist() if 'обратный' in str(x).lower()]
            s_list += [x for x in xl_arr.tolist() if 'вентиль' in str(x).lower()]
            s_list += [x for x in xl_arr.tolist() if 'шаровой' in str(x).lower()]
            j = 1
            for i in range(0, len(s_list)):
                s_list[i].insert(0, j)
                j += 1
            j = 1

            for row in s_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])
        elif type_choies.get() == 'Акт КИПиА':

            th_list = np.empty((1, len(xl_arr.tolist()[0]))).tolist()
            th_list.clear()

            th_list += [x for x in xl_arr.tolist() if 'Манометр' in str(x)]
            th_list += [x for x in xl_arr.tolist() if 'термометр' in str(x).lower()]
            th_list += [x for x in xl_arr.tolist() if 'термостат погружной' in str(x).lower()]
            th_list += [x for x in xl_arr.tolist() if 'датчик' in str(x).lower()]
            th_list += [x for x in xl_arr.tolist() if 'реле' in str(x).lower()]
            th_list += [x for x in xl_arr.tolist() if 'прессостат' in str(x).lower()]

            j = 1
            for i in range(0, len(th_list)):
                th_list[i].insert(0, j)
                j += 1
            j = 1

            for row in th_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])
        else:

            xl_list = xl_arr.tolist()

            j = 1
            for i in range(0, len(xl_arr)):
                xl_list[i].insert(0, j)
                j += 1
            j = 1

            for row in xl_list:
                row_cells = new_table.add_row().cells
                for i in range(cols_number):
                    row_cells[i].text = str(row[i])

        document.save(filepath)


def open_table():
    open_spec()
    table_path=filedialog.askopenfilename()
    global xl_arr
    if table_path !="":
        df = pd.read_excel(table_path, sheet_name='Table 1', skiprows=2)
        df_cleaned=df.dropna()
        print(df_cleaned.head())
        print(len(df_cleaned.index))
        print(df_cleaned.columns)
        print(len(df_cleaned.to_numpy()))
        for i in range(0,5):
            print(df_cleaned.to_numpy()[i])
        xl_arr=df_cleaned.to_numpy()

def create_table(document, headers, rows, style='Table Grid'):
    cols_number = len(headers)

    table = document.add_table(rows=1, cols=cols_number)
    table.style = style

    hdr_cells = table.rows[0].cells
    for i in range(cols_number):
        hdr_cells[i].text = headers[i]

    for row in rows:
        row_cells = table.add_row().cells
        for i in range(cols_number):
            row_cells[i].text = str(row[i])

    return table

root=Tk()
root.title('Подготовка Актов')
root.geometry('600x600+100+100')
root.iconbitmap(default='brend.ico')
font1 = font.Font(family= "Times New Roman", size=11, weight="bold", slant="roman", underline=False, overstrike=False)
font2 = font.Font(family= "Times New Roman", size=11, weight="normal", slant="roman", underline=False, overstrike=False)

name_form=Label(root, text='Заполните данные шапки акта', font=("Arial", 11, "bold"))
name_form.place(x=20, y=20)

station =Entry(root, font=font1)
station.place(x=20, y=60, width=650)
station.insert(0,'Введите название установки')

calc =Entry(root, font=font1)
calc.place(x=20, y=100, width=650)
calc.insert(0,'Введите номер расчета')

company =Entry(root, font=font2)
company.place(x=20, y=140, width=650)
company.insert(0,'Введите название компании')

obj =Entry(root, font=font2)
obj.place(x=20, y=180, width=650)
obj.insert(0,'Введите название обьекта')

address =Entry(root, font=font2)
address.place(x=20, y=220, width=650)
address.insert(0,'Введите название адреса')

number =Entry(root, font=font1)
number.place(x=20, y=260, width=650)
number.insert(0,'Введите номер акта')

name =Entry(root, font=font1)
name.place(x=20, y=300, width=650)
name.insert(0,'Введите название акта')

data =Entry(root, font=font2)
data.place(x=20, y=340, width=650)
data.insert(0,'Введите дату')

acttype=Label(root, text='Выберите тип акта', font=("Arial", 11, "bold"))
acttype.place(x=20, y=380)

type_acts=['Акт основного оборудования', 'Акт вспомогательного оборудования', 'Акт КИПиА']
# по умолчанию будет выбран первый элемент из languages
type_var = StringVar(value=type_acts[0])

type_choies=Combobox(textvariable=type_var, values=type_acts, state="readonly")
type_choies.place(x=20, y=410)


file_button=Button(text='Открыть спец', command=open_table, font=("Arial", 12, "bold"))
file_button.place(x=400, y=20)

btn=Button(text='Создать Акт', command=safe_act, font=("Arial", 12, "bold"))
btn.place(x=400, y=400)


# Загрузка шаблона
document = DocxTemplate("Шаблон.docx")
xl_arr=np.zeros((1,5))

root.mainloop()