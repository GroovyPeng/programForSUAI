from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup as BS
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from schedule import ScheduleTable


HOME_URL = 'https://rasp.guap.ru/'
HTML = BS(requests.get(HOME_URL).content, "lxml")
ADRESSES = [number['value'] for number in HTML.find(
    attrs={'name': 'ctl00$cphMain$ctl07'}).find_all('option')]
# Удаление Дистант и РИМР
ADRESSES.remove('4')
ADRESSES.remove('7')
CABINETS = [number['value'] for number in HTML.find(
    attrs={'name': 'ctl00$cphMain$ctl08'}).find_all('option')]


def parse(URL: str):
    return BS(requests.get(URL).content, "lxml")


def create_docx(HTML, PATH):
    document = Document()
    sections = document.sections

    for section in sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    title = HTML.find("h2").text
    header = document.add_heading(title)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER

    result = HTML.find(class_='result')
    ScheduleTable(document, result)
    document.save(PATH)


def GUI():
    def safe():
        adress = ADRESSES[adress_names.index(adr_cmbox.get())]
        cabinet = CABINETS[cabinet_numbers.index(cbnt_cmbox.get())]
        URL = f'{HOME_URL}?b={adress}&r={cabinet}'
        HTML = parse(URL)

        if HTML.find('h3') != None:
            files_type = [('MS Word', '*.docx'),
                          ('All files', '.')]
            file = filedialog.asksaveasfile(
                mode='w',
                filetypes=files_type,
                defaultextension=files_type,
                initialfile=f'{cbnt_cmbox.get()} ({adr_cmbox.get()})')

            PATH = file.name
            create_docx(HTML, PATH)
        else:
            messagebox.showinfo('Расписание отсутствует',
                                'По заданным данным расписание отсутствует!')

    def is_chosen() -> None:
        check_adress = adress_names.index(adr_cmbox.get()) != 0
        check_cabinet = cabinet_numbers.index(cbnt_cmbox.get()) != 0

        if check_adress and check_cabinet:
            safe()
        else:
            msg = 'Пожалуйста, выберете '
            if not check_adress and not check_cabinet:
                msg += 'адрес и номер кабинета.'
            elif check_adress:
                msg += 'номер кабинета.'
            else:
                msg += 'адрес учреждения.'
            messagebox.showinfo('Неполные данные', msg)

    adress_names = [HTML.find(attrs={'name': 'ctl00$cphMain$ctl07'}).find(
        attrs={'value': id}).text for id in ADRESSES]
    cabinet_numbers = [HTML.find(attrs={'name': 'ctl00$cphMain$ctl08'}).find(
        attrs={'value': id}).text for id in CABINETS]

    root = Tk()
    # Размещение окна в центре экрана
    x = (root.winfo_screenwidth() - root.winfo_reqwidth()) / 2
    y = (root.winfo_screenheight() - root.winfo_reqheight()) / 2
    root.wm_geometry("+%d+%d" % (x, y))

    root.title('Расписание ГУАП')
    root.geometry('300x300')
    root.resizable(False, False)

    adress_frame = Frame(root)
    adress_frame.place(rely=.1, relx=.25, relwidth=.50)
    lbl = Label(adress_frame, text='Адрес учреждения:')
    lbl.pack()
    adr_cmbox = ttk.Combobox(adress_frame, values=adress_names)
    adr_cmbox.current(0)
    adr_cmbox.pack()

    cabinet_frame = Frame(root)
    cabinet_frame.place(rely=.25, relx=.25, relwidth=.50)
    lbl = Label(cabinet_frame, text='Номер кабинета:')
    lbl.pack()
    cbnt_cmbox = ttk.Combobox(cabinet_frame, values=cabinet_numbers)
    cbnt_cmbox.current(0)
    cbnt_cmbox.pack()

    ttk.Button(root, text='Сохранить', width=20,
               command=is_chosen).pack(side=BOTTOM, pady=50)
    root.mainloop()


if __name__ == '__main__':
    GUI()
