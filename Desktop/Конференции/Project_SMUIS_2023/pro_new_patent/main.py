import openpyxl
import pathlib
from selenium import webdriver
from time import sleep
import json
from tkinter import *
from tkinter import ttk
from tkinter import StringVar

def extensionCheck():
    file_name_str = file_name.get()
    list_name_str = list_name.get()
    cells_str = cells.get()
    directory_str = directory.get()

    file_extension = pathlib.Path(file_name_str).suffix
    if file_extension == '.xlsx':
        work_bk = openpyxl.load_workbook(file_name_str)
        sheet = work_bk[list_name_str]
        vals = [v[0].value for v in sheet[cells_str]]
        chrome_options = webdriver.ChromeOptions()
        settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
                    "selectedDestinationId": "Save as PDF", "version": 2}
        prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                 'savefile.default_directory': directory_str}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--kiosk-printing')
        browser = webdriver.Chrome(r"chromedriver.exe", options=chrome_options)
        for numb in vals:
            browser.get(f'https://new.fips.ru/registers-doc-view/fips_servlet?DB=RUPAT&DocNumber={numb}&TypeFile=html')
            sleep(1)
            browser.execute_script('window.print();')
    elif file_extension == '.txt':
        vals2 = []
        with open(file_name_str, 'r') as file:
            for line in file:
                number = int(line.strip())
                vals2.append(number)
            chrome_options = webdriver.ChromeOptions()
            settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
                        "selectedDestinationId": "Save as PDF", "version": 2}
            prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                     'savefile.default_directory': directory_str}
            chrome_options.add_experimental_option('prefs', prefs)
            chrome_options.add_argument('--kiosk-printing')
            browser = webdriver.Chrome(r"chromedriver.exe", options=chrome_options)
            for numb in vals2:
                browser.get(f'https://new.fips.ru/registers-doc-view/fips_servlet?DB=RUPAT&DocNumber={numb}&TypeFile=html')
                sleep(1)
                browser.execute_script('window.print();')

def set_file_name(value):
    global file_name
    file_name = value


def startScreen():
    lbl9 = Label(root, background="gray22", fg='white', text="Групповой поиск", font=("TkTextFont", 12))
    lbl9.pack(side=TOP, padx=5, pady=5)

    lbl10 = Label(root, background="gray22", fg='white', text="Название файла", font=("TkTextFont", 10))
    lbl10.pack(side=TOP, padx=5, pady=5)

    entryFile = ttk.Entry(textvariable=file_name)
    entryFile.pack(side=TOP, padx=5, pady=5, ipadx=20)

    lbl11 = Label(root, background="gray22", fg='white', text="Название листа", font=("TkTextFont", 10))
    lbl11.pack(side=TOP, padx=5, pady=5)

    entryList = ttk.Entry(textvariable=list_name)
    entryList.pack(side=TOP, padx=5, pady=5, ipadx=20)

    lbl12 = Label(root, background="gray22", fg='white', text="Диапазон ячеек (A2:A7)", font=("TkTextFont", 10))
    lbl12.pack(side=TOP, padx=5, pady=5)

    entryCells = ttk.Entry(textvariable=cells)
    entryCells.pack(side=TOP, padx=5, pady=5, ipadx=20)

    lbl13 = Label(root, background="gray22", fg='white', text="Директория", font=("TkTextFont", 10))
    lbl13.pack(side=TOP, padx=5, pady=5)

    entryDirectory = ttk.Entry(textvariable=directory)
    entryDirectory.pack(side=TOP, padx=5, pady=5, ipadx=20)

    button = Button(root, text="Отправить", command=lambda: extensionCheck())
    button.pack(side=TOP, padx=5, pady=15, ipadx=15)

root = Tk()
root.title("Patent-loader")
root.geometry("400x350")
root["bg"] = "gray22"

file_name = StringVar()
list_name = StringVar()
cells = StringVar()
directory = StringVar()

startScreen()

root.mainloop()
