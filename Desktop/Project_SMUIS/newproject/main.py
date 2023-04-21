import openpyxl
import pathlib
from selenium import webdriver
from time import sleep
import json
from tkinter import *
from tkinter import ttk
from tkinter.ttk import Combobox
from tkinter import filedialog as fd
import threading

file_name=''
name_directory=''
list_name = ''
cells = ''


root = Tk()
root.title("Patent-loader")
root.geometry("400x350")
root["bg"] = "gray22"
lbl1 = Label(root, background="gray22", fg='white', text="Одиночный поиск", font=("TkTextFont", 12))
lbl1.grid(column=1)
lbl2 = Label(root, background="gray22", fg='white', text="Поисковые параметры", font=("TkTextFont", 10))
lbl2.grid(column=0, row=2)

combo = Combobox(root)
combo['values'] = ["Патентные документы", "Товарные знаки", "Промышленные образцы", "Программы для ЭВМ", "БД", "ТИМС"]
combo.grid(column=1, row=2, ipadx=30)

lbl3 = Label(root, background="gray22", fg='white', text="Номер регистрации", font=("TkTextFont", 10))
lbl3.grid(column=0, row=3)
txt = Entry(root, width=23)
txt.grid(column=1, row=3, ipadx=30)
txt.focus()

lbl4 = Label(root, background="gray22", fg='white', text="Дата регистрации", font=("TkTextFont", 10))
lbl4.grid(column=0, row=4)
txt2 = Entry(root, width=23)
txt2.grid(column=1, row=4, ipadx=30)
txt2.focus()

lbl5 = Label(root, background="gray22", fg='white', text="Автор (ФИО)", font=("TkTextFont", 10))
lbl5.grid(column=0, row=5)
txt3 = Entry(root, width=23)
txt3.grid(column=1, row=5, ipadx=30)
txt3.focus()

lbl6 = Label(root, background="gray22", fg='white', text="Автор (ФИО)", font=("TkTextFont", 10))
lbl6.grid(column=0, row=5)
txt4 = Entry(root, width=23)
txt4.grid(column=1, row=5, ipadx=30)
txt4.focus()

lbl7 = Label(root, background="gray22", fg='white', text="Правообладатель", font=("TkTextFont", 10))
lbl7.grid(column=0, row=5)
txt5 = Entry(root, width=23)
txt5.grid(column=1, row=5, ipadx=30)
txt5.focus()

lbl8 = Label(root, background="gray22", fg='white', text="Название", font=("TkTextFont", 10))
lbl8.grid(column=0, row=6)
txt6 = Entry(root, width=23)
txt6.grid(column=1, row=6, ipadx=30, ipady=15)
txt6.focus()

lbl9 = Label(root, background="gray22", fg='white', text="Групповой поиск", font=("TkTextFont", 12))
lbl9.grid(column=1, row=7)

lbl10 = Label(root, background="gray22", fg='white', text="Книга Excel либо txt-файл", font=("TkTextFont", 10))
lbl10.grid(column=0, row=8)

lbl12 = Label(root, background="gray22", fg='white', text="Книга Excel", font=("TkTextFont", 11))
lbl12.grid(column=0, row=9)

lbl11 = Label(root, background="gray22", fg='white', text="Название листа книги", font=("TkTextFont", 10))
lbl11.grid(column=0, row=10)
txt7 = Entry(root, width=23, textvariable=list_name)
txt7.grid(column=1, row=10, ipadx=30)
txt7.focus()

lbl13 = Label(root, background="gray22", fg='white', text="Ячейки (формат-A1:A3)", font=("TkTextFont", 10))
lbl13.grid(column=0, row=11)
txt8 = Entry(root, width=23, textvariable=cells)
txt8.grid(column=1, row=11, ipadx=30)
txt8.focus()

def open_text_file():
    # Specify the file types
    filetypes = (('text files', '*.txt'),
                 ('All files', '*.*'))
    fd.askopenfile(filetypes=filetypes,
                       initialdir="D:/Downloads")

def dialog_papka():
    fd.askdirectory()

open_button = ttk.Button(root, text='Open a File', command=open_text_file, textvariable=file_name )
open_button.grid(column=1, row=8, padx=20, sticky='w', ipadx=50)

lbl14 = Label(root, background="gray22", fg='white', text="Директория для выгрузки", font=("TkTextFont", 12))
lbl14.grid(column=1, row=12)

open_button1 = ttk.Button(root, text='Open a Directory', command=dialog_papka, textvariable=name_directory)
open_button1.grid(column=1, row=13, padx=20, sticky='w', ipadx=40)

root.mainloop()



file_extension = pathlib.Path(file_name).suffix
if file_extension == '.xlsx':

    work_bk = openpyxl.load_workbook(file_name)
    sheet = work_bk[list_name]
    vals = [v[0].value for v in sheet[cells]]

    chrome_options = webdriver.ChromeOptions()
    settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
                "selectedDestinationId": "Save as PDF", "version": 2}
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                'savefile.default_directory': name_directory}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    browser = webdriver.Chrome(r"chromedriver.exe", options=chrome_options)
    for numb in vals:
        browser.get(f'https://new.fips.ru/registers-doc-view/fips_servlet?DB=RUPAT&DocNumber={numb}&TypeFile=html')
        sleep(1)
        browser.execute_script('window.print();')

elif file_extension == '.txt':
    vals2 = []
    with open('patents.txt', 'r') as file:
        for line in file:
            number = int(line.strip())
            vals2.append(number)
    chrome_options = webdriver.ChromeOptions()
    settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
                "selectedDestinationId": "Save as PDF", "version": 2}
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
             'savefile.default_directory': 'C:\\Users\\Екатерина\\Desktop\\Project_SMUIS'}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    browser = webdriver.Chrome(r"chromedriver.exe", options=chrome_options)
    for numb in vals2:
        browser.get(f'https://new.fips.ru/registers-doc-view/fips_servlet?DB=RUPAT&DocNumber={numb}&TypeFile=html')
        sleep(1)
        browser.execute_script('window.print();')

