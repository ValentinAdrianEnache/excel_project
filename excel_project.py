"""
Excel automation project
"""

import datetime
import time
import tkinter as tk
import random
from tkinter.filedialog import askopenfilename
from tkinter import *
from datetime import datetime
import openpyxl

FILE1=''
FILE2=''
FILE3 = 'date_hour.xlsx'

WB = openpyxl.load_workbook(FILE3)
SHEET = WB.active


#GUI
WINDOW = tk.Tk()
WINDOW.title('AppX v1.0')
WINDOW.geometry("1000x550")
WINDOW.configure(background="#1D6F42")
background_text=Label(WINDOW,text='KEEP\nCALM\nIT'+"'"+'S JUST AN\nEXCEL\nFILE',
           bg='#1D6F42',
           font=("Arial", 35, "bold"),
           fg="white")
background_text.pack()
background_text.place(x=20, y=85)


# click_btn= PhotoImage(file='browse.png')

BUTON_FILE_1 = tk.Button(
    # image=click_btn,
    text="Load",
    # width=180,
    width=15,
    # height=100,
    height=3,
    font=("Arial", 15, "bold"),
    bg='#abf77e',
    fg="black",
)
BUTON_FILE_1.pack()
BUTON_FILE_1.place(x=450,y=20)

BUTON_FILE_2 = tk.Button(
    text="Report",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#abf77e",
    fg="black",
)
BUTON_FILE_2.pack()
BUTON_FILE_2.place(x=550,y=120)

BUTON_EXECUTION = tk.Button(
    text="Execute",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#abf77e",
    fg="black",
)
BUTON_EXECUTION.pack()
BUTON_EXECUTION.place(x=450,y=220)

BUTON_DATE_RECORDER = tk.Button(
    text="Rec",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#3EB489",
    fg="black",
)
BUTON_DATE_RECORDER.pack()
BUTON_DATE_RECORDER.place(x=550,y=320)

BUTON_RESET_DATA = tk.Button(
    text="Delete",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#a80505",
    fg="black",
)
BUTON_RESET_DATA.pack()
BUTON_RESET_DATA.place(x=450,y=420)


def file_1(event):
    """
    :param event:
    :return:FILE1 name
    """
    global FILE1
    FILE1 = askopenfilename(
        filetypes=[('text', '*.xlsx'), ('all files', '*.*')]
    )
    print(FILE1,' loaded!')



def file_2(event):
    """
        :param event:
        :return:FILE2 name
        """
    global FILE2
    FILE2 = askopenfilename(
        filetypes=[('text', '*.xlsx'), ('all files', '*.*')]
    )
    print(FILE2,' loaded!')

def execute(event):
    """
    :param event:
    :return: create report
    """
    # load and read from file 1
    WB = openpyxl.load_workbook(FILE1)
    SHEET = WB.active
    no_list = []
    t_list = []
    for i in range(1, 22):
        no_cell = SHEET['A' + str(i)].value
        no_list.append(no_cell)
        temp_cell = SHEET['B' + str(i)].value
        t_list.append(temp_cell)
    # load and read from FILE3
    WB = openpyxl.load_workbook(FILE3)
    SHEET = WB.active
    h_list = []
    d_list = []
    for i in range(1, 22):
        hour_cell = SHEET['B' + str(i)].value
        h_list.append(hour_cell)
        date_cell = SHEET['C' + str(i)].value
        d_list.append(date_cell)
    # write report
    line_2 = [str(i) + '2' for i in range_letter("A", "Z")]
    line_3 = [str(i) + '3' for i in range_letter("A", "Z")]
    line_4 = [str(i) + '4' for i in range_letter("A", "Z")]
    line_5 = [str(i) + '5' for i in range_letter("A", "Z")]

    WB = openpyxl.load_workbook(FILE2)
    SHEET = WB.active

    for i in range(21):
        # start_no_list= start index for each list
        start_no_list = no_list[i]
        start_t_list = t_list[i]
        start_h_list = h_list[i]
        start_d_list = d_list[i]

        SHEET[line_2[i]].value = start_no_list
        SHEET[line_3[i]].value = start_t_list
        SHEET[line_4[i]].value = start_h_list
        SHEET[line_5[i]].value = start_d_list

    print('Report has been created! Report path: ', FILE2)
    WB.save(FILE2)

def range_letter(start, stop):
    """
    :param start: first letter
    :param stop: last letter
    :return: range
    """
    return (chr(n) for n in range(ord(start), ord(stop) + 1))

def date_recorder(event):
    """
    :param event:
    :return: record data
    """
    global file3

    # write date and hour:
    WB = openpyxl.load_workbook(FILE3)
    SHEET = WB['Sheet']

    date_now = datetime.now()  # current date and hour
    hour_now = date_now.strftime('%H:%M:%S')  # hour format
    day_now = date_now.strftime('%d''-''%m''-''%y')  # day format
    epoch = time.time()  # epoch time in seconds (from 01.01.1970)

    numero_cell = ['A' + str(i) for i in range(1, 22)]
    hour_cell = ['B' + str(i) for i in range(1, 22)]
    date_cell = ['C' + str(i) for i in range(1, 22)]
    epoch_cell = ['D' + str(i) for i in range(1, 22)]

    count = 1
    delta = 6
    wait_sec = 5

    for i in epoch_cell[1:21]:
        if SHEET[i].value is not None:
            time_float = float(SHEET[i].value)
            delta = float(epoch) - time_float
            count += 1
            if count == 21:
                print('Full memory! Press delete for reset data!')
            continue

        elif delta > wait_sec:
            SHEET[i].value = epoch
            SHEET[date_cell[count]].value = day_now
            SHEET[hour_cell[count]].value = hour_now
            SHEET[numero_cell[count]].value = str(count) + '.'
            print('Values has been recorded!')
            deg_reg() #write temperature data
            break
        else:
            print(f'Please wait {wait_sec - int(delta)} seconds until the next record! ')
            break
    WB.save(FILE3)

def reset_data(event):
    """
    :param event:
    :return: clear table from all files
    """
    # read and load FILE3
    global FILE3
    WB = openpyxl.load_workbook(FILE3)
    SHEET = WB.active
    # delete columns
    for i in range(4):  #using for loop with range(4) to delete column one four times
        SHEET.delete_cols(1)  # when deleted column 1 it's been replaced with column 2 and so on
    # write
    numero = SHEET.cell(row=1, column=1)
    hour = SHEET.cell(row=1, column=2)
    date = SHEET.cell(row=1, column=3)
    time = SHEET.cell(row=1, column=4)

    number_title = 'No.'
    hour_title = 'Hour'
    date_title = 'Date'
    epoch_title = 'Epoch'

    numero.value = number_title
    hour.value = hour_title
    date.value = date_title
    time.value = epoch_title
    WB.save(FILE3)

    global FILE1
    WB = openpyxl.load_workbook(FILE1)
    SHEET = WB.active

    for i in range(2):
        SHEET.delete_cols(1)
        numero = SHEET.cell(row=1, column=1)
        degree = SHEET.cell(row=1, column=2)
        number = 'No.'
        d1 = 'Temperature(°C)'
        numero.value = number
        degree.value = d1
    WB.save(FILE1)
    print('Data has been deleted!')




def deg_reg():
    """
    :return: temperature data
    """
    global FILE1
    WB = openpyxl.load_workbook(FILE1)
    SHEET = WB['Sheet']
    degree = random.randint(18, 22)
    numero_cell = ['A' + str(i) for i in range(1, 22)]
    degree_cell = ['B' + str(i) for i in range(1, 22)]
    count = 0
    for i in numero_cell[1:21]:
        count += 1
        if SHEET[i].value is None:
            SHEET[i].value = str(count) + '.'
            SHEET[degree_cell[count]].value = str(degree) + '°'
            break
    WB.save(FILE1)


BUTON_FILE_1.bind("<Button-1>", file_1)
BUTON_FILE_2.bind("<Button-1>", file_2)
BUTON_EXECUTION.bind("<Button-1>", execute)
BUTON_RESET_DATA.bind("<Button-1>", reset_data)
BUTON_DATE_RECORDER.bind("<Button-1>", date_recorder)
WINDOW.mainloop()

