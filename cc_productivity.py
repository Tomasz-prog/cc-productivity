import logging
import tkinter as tk
from tkinter import ttk
import pandas as pd
import os
# import datetime
from datetime import datetime
# from datetime import timedelta
# import time
import datetime
from tkinter import *

global back

try:
    filepath = "adres.txt"
    f = open(filepath, "r")
    adres1 = (f.readline())
    adres1 = adres1[:-1]
    f.close()

    df_cc = pd.read_excel('liczenie.xlsx', usecols=['Date planned', 'Kind of count', 'Location from', 'Start time 1st count','End time 1st count','User code 1st count', 'Sub batch status nr'])
    kodys = pd.read_excel('Usercodes list.xlsx', usecols=['USERCODE','NAAM'], index_col='USERCODE')
except FileNotFoundError as e:
    logging.exception(("brak pliku 'liczenie.xlsx' lub 'adres.txt'"))
# ---------------------- clear functions -------------------------------
def clear_right_excell_side():
    wynik_label = tk.StringVar()
    wynik_label.set("")
    wynik_label_print = tk.Label(root, textvariable=wynik_label, font=("Times New Roman", 15))
    wynik_label_print.place(x=10, y=70, height=2000, width=2000)
def passs():
    pass
def clear_up_header_without_back_btt():
    wynik_label = tk.StringVar()
    wynik_label.set("")
    wynik_label_print = tk.Label(root, textvariable=wynik_label, font=("Times New Roman", 15))
    wynik_label_print.place(x=70, y=10, height=40, width=900)
def clear():
    name_f1 = tk.StringVar()
    name_f1.set(" ")
    name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
    name_f1_print.place(x=10, y=50, height=700, width=1500)
# -------------------------- widgets labels functions -------------------------
def label_reports_output(h_min, m_min, h_max, m_max, break_time):
    print(break_time)
    clear()
    df_cc = pd.read_excel('liczenie.xlsx',
                          usecols=['Date planned', 'Kind of count', 'Location from', 'Start time 1st count',
                                   'End time 1st count', 'User code 1st count', 'Sub batch status nr'])
    kodys = pd.read_excel('Usercodes list.xlsx', usecols=['USERCODE', 'NAAM'], index_col='USERCODE')
    frame = pd.read_excel('./Details_data.xlsx',
                          usecols=["Date_Trunc", "UserCode", "UserName", "Zone", "Measure Names",
                                   "Measure Values"])

    date = '07.07.2021'

    seria_user = frame['UserCode']
    seria_user = seria_user.drop_duplicates()
    seria_user = list(seria_user)
    print(seria_user)

    axis = 100
    data = {'Name': [], 'User': [], 'Start': [], 'Stop': [], 'Full Time': [], 'Item Qty': [],
            'UPH': []}
    cc_report = pd.DataFrame(data)
    avr = 0
    loc = 0
    list_of_names = []
    list_of_users = []
    list_of_timestart = []
    list_of_timestop = []
    list_of_fulltime = []
    list_of_numbers_of_items = []
    list_of_norma = []

    # filtr_user = (frame["UserCode"]) == int(user)
    for i in range(len(seria_user)):
        try:
            user = seria_user[i]
            try:
                name = kodys.loc[user, 'NAAM']
            except:
                name ='NO DATA'

            filtr_user = (frame["UserCode"]) == int(user)
            filtr_date = (frame["Date_Trunc"]) == date
            filtr_items = (frame["Measure Names"]) == "QtyStockAfter"
            filtr_time = (frame["Measure Names"]) == "TimeDiff"

            # set a sorted dataframe for prepare sum of items

            frame_items = frame[filtr_date & filtr_user & filtr_items]
            sum_items = sum(frame_items['Measure Values'])

            # set a sorted dataframe for prepare time
            frame_time = frame[filtr_date & filtr_user & filtr_time]
            times = list(frame_time['Measure Values'])

            h, m, s = 0, 0, 0
            time, result = 0, 0
            all_time, old_value = 0, 0
            for i in range(len(times)):
                # print(times[i].hour,times[i].minute,times[i].second)
                # result = times[i].hour * 3600 + times[i].minute * 60 + times[i].second
                # all_time += result

                old_value = times[i]
                time = old_value * 24
                h = int(time)
                m = (time * 60) % 60
                s = (time * 3600) % 60
                print("%d:%02d.%02d" % (h, m, s))
                result = h * 3600 + m * 60 + s
                all_time += result

            all_time = all_time / 3600

            # output



            list_of_names.append(name)
            list_of_users.append(user)
            list_of_fulltime.append(all_time)
            list_of_numbers_of_items.append(sum_items)
            try:
                list_of_norma.append(sum_items/all_time)
            except: continue

            # liczba kolumn
            tv = ttk.Treeview(root, columns=(1, 2, 3, 4, 5, 6), show='headings', height=20)
            # szerokosci kolumn
            tv.column("1", width=30, anchor="c")
            tv.column("2", width=200, anchor="c")
            tv.column("3", width=60, anchor="c")
            tv.column("4", width=70, anchor="c")
            tv.column("5", width=70, anchor="c")
            tv.column("6", width=70, anchor="c")

            tv.place(x=10, y=200)

            # nagłówki kolumn
            tv.heading(1, text="Lp")
            tv.heading(2, text="Name")
            tv.heading(3, text="UserCode")
            tv.heading(4, text="Full time")
            tv.heading(5, text="loc qty")
            tv.heading(6, text="UPH")


            # konfigurowanie scrollbar
            sb = Scrollbar(root, orient=VERTICAL)
            sb.place(x=1400, y=50, width=20, height=500)

            tv.config(yscrollcommand=sb.set)
            sb.config(command=tv.yview)

            style = ttk.Style()
            style.theme_use("alt")
            style.configure("Treeview", rowheight=20)
            style.map("Treeview")

            count = 0
            no = 1
            for i in range(len(list_of_users)):
                if count % 2 == 0:
                    tv.insert(parent='', index=i, iid=i,
                              values=(no, list_of_names[i], list_of_users[i],
                                        list_of_fulltime[i],
                                      list_of_numbers_of_items[i], list_of_norma[i]),
                              tags=('evenrow',))
                else:
                    tv.insert(parent='', index=i, iid=i,
                              values=(no, list_of_names[i], list_of_users[i],
                                       list_of_fulltime[i],
                                      list_of_numbers_of_items[i], list_of_norma[i]),
                              tags=('oddrow',))
                count += 1
                no += 1

            data = {'Name': name, 'User': user, 'Full Time': all_time, 'Loc qty': sum_items, 'UPH': sum_items/all_time}
            cc_report = cc_report.append(data, ignore_index=True)


        except IndexError:
            continue
    try:
        # ----------------------- count avr ------------------------------
        avr = round(avr/len(cc_report),0)
        stop_time_serie = cc_report['Stop']
        stop_time_serie = list(stop_time_serie.sort_values(ascending=True))
        print(stop_time_serie)
        long = len(stop_time_serie)


        name_f1 = tk.StringVar()
        name_f1.set(f"last activiy: {stop_time_serie[long-1]}")
        name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
        name_f1_print.place(x=800, y=70, height=25, width=200)

        label_avr('AVR UPH', 100, 70, 25, 80)
        label_avr(avr, 100, 100, 25, 80)
        label_avr('suma of loc', 200, 70, 25, 100)
        label_avr(loc, 200, 100, 25, 80)

    except ZeroDivisionError:
        pass

    def excell_reports():
        try:
            output = 'cc_report.xlsx'


            with pd.ExcelWriter(output) as writer:
                cc_report.to_excel(writer, sheet_name='cc_report', engine='openpyxl')

                wynik_label = tk.StringVar()
                wynik_label.set(f"zapisano")
                wynik_label_print = tk.Label(root, textvariable=wynik_label, font=("Times New Roman", 15))
                wynik_label_print.place(x=1000, y=10, height=25, width=120)
                os.startfile('cc_report.xlsx')
        except:
            wynik_label = tk.StringVar()
            wynik_label.set(f"błąd zapisu")
            wynik_label_print = tk.Label(root, textvariable=wynik_label, font=("Times New Roman", 15))
            wynik_label_print.place(x=1000, y=10, height=25, width=120)



    btt6 = tk.Button(root, text="excell", command=excell_reports, bg="red", fg="black", bd=5)
    btt6.place(x=850, y=10, height=30, width=90)

def label_avr(value, x,y,height, width):

    name_f1 = tk.StringVar()
    name_f1.set("")
    name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
    name_f1_print.place(x=x, y=y, height=height, width=width)

    name_f1 = tk.StringVar()
    name_f1.set(value)
    name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
    name_f1_print.place(x=x, y=y, height=height, width=width)
# ------------------------ under construction --------------------------
def report_all_options():
    global back
    clear_up_header_without_back_btt()
    back = 'main'

    powrot_btt_a = tk.Button(root, text=("back"), command=back_command, bg="red", fg="black")
    powrot_btt_a.place(x=10, y=10)
def loc_checks():
    clear()
    df_cc = pd.read_excel('liczenie.xlsx',
                          usecols=['Date planned', 'Kind of count', 'Location from', 'Start time 1st count',
                                   'End time 1st count', 'User code 1st count', 'Sub batch status nr'])

    loc_check = df_cc
    filtr = (loc_check['Sub batch status nr'] != 10)
    loc_check = loc_check[filtr]

    seria_loc = loc_check['Location from']
    seria_loc = list(seria_loc)
    axis=100


    if len(seria_loc) != 0:
        name_f1 = tk.StringVar()
        name_f1.set('niezamknięte lokacje')
        name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
        name_f1_print.place(x=10, y=70, height=25, width=300)

    else:
        name_f1 = tk.StringVar()
        name_f1.set('wszystkie lokacje zostały zamknięte')
        name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
        name_f1_print.place(x=10, y=70, height=25, width=300)



    for i in range(len(loc_check)):

        name_f1 = tk.StringVar()
        name_f1.set(seria_loc[i])
        name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
        name_f1_print.place(x=10, y=axis, height=25, width=200)

        axis += 30
def report_CC():
    def excell():

        output = 'cc_report.xlsx'


        with pd.ExcelWriter(output) as writer:
            cc_report.to_excel(writer, sheet_name='cc_report', engine='openpyxl')

            wynik_label = tk.StringVar()
            wynik_label.set(f"File: ' {output} '")
            wynik_label_print = tk.Label(root, textvariable=wynik_label, font=("Times New Roman", 15))
            wynik_label_print.place(x=10, y=320, height=25, width=600)


        # wynik_label = tk.StringVar()
        # wynik_label.set("THE FILE HAS NOT BEEN SAVED   ")
        # wynik_label_print = tk.Label(root, textvariable=wynik_label, font=("Times New Roman", 15))
        # wynik_label_print.place(x=10, y=320, height=25, width=600)
# ----------------------- programmed functions and working -------------------
def back_command():
    clear()
    clear_up_header_without_back_btt()
    clear_right_excell_side()
    if back == 'report':
        main()
    else:
        pass
def report():
    clear_right_excell_side()

    def all():
        clear_right_excell_side()
        label_reports_output(0,1,23,59,checkvar.get())
    def six_fourteen():
        clear_right_excell_side()
        label_reports_output(6,0,14,0,checkvar.get())
    def sixteen_twentyfour():
        clear_right_excell_side()
        label_reports_output(16,0,23,59,checkvar.get())
    def zero_six():
        clear_right_excell_side()
        label_reports_output(0,0,6,0, checkvar.get())
    def zero_eight():
        clear_right_excell_side()
        label_reports_output(0,0,8,0, checkvar.get())
    def six_eight():

        clear_right_excell_side()
        label_reports_output(6,0,8,0,checkvar.get())

    def set():
        clear()
        clear_right_excell_side()
        wynik_label_print = tk.Label(root, text="", font=("Times New Roman", 10))
        wynik_label_print.place(x=50, y=70, height=30, width=50)
        # ------------------ header ------------------------------------
        header = tk.Label(root, text="Hour  Min            Hour  Min  ", font=("Times New Roman", 15))
        header.place(x=50, y=50, height=30, width=260)

        # ----------------- hours and minute from ---------------------
        hour_min = tk.IntVar
        entry_hour_min = tk.Entry(root, textvariable=hour_min, bg="orange", bd=5, font=50)
        entry_hour_min.place(x=50, y=80, height=50, width=50)
        min_min = tk.IntVar
        entry_min_min = tk.Entry(root, textvariable=min_min, bg="orange", bd=5, font=50)
        entry_min_min.place(x=100, y=80, height=50, width=50)
        # -------------- hours and minutes to
        hour_max = tk.IntVar
        entry_hour_max = tk.Entry(root, textvariable=hour_max, bg="orange", bd=5, font=50)
        entry_hour_max.place(x=200, y=80, height=50, width=50)
        min_max = tk.IntVar
        entry_min_max= tk.Entry(root, textvariable=min_max, bg="orange", bd=5, font=50)
        entry_min_max.place(x=250, y=80, height=50, width=50)

        def set_push():
            label_reports_output(int(entry_hour_min.get()), int(entry_min_min.get()), int(entry_hour_max.get()), int(entry_min_max.get()), checkvar.get())


        conf = tk.Button(root, text="confirm", command=set_push, bg="blue", fg="white", bd=5)
        conf.place(x=350, y=80, height=30, width=90)




    global back
    back = 'report'
    clear_up_header_without_back_btt()

    checkvar = IntVar()
    break_check = Checkbutton(root, text="Break", variable=checkvar,
                     onvalue=1, offvalue=0, height=20,
                     width=30)

    break_check.place(x=70, y=10, height=30, width=70)



    btt1 = tk.Button(root, text="All", command=all, bg="green", fg="white", bd=5)
    btt1.place(x=150, y=10, height=30, width=90)

    btt2 = tk.Button(root, text="6-14", command=six_fourteen, bg="green", fg="white", bd=5)
    btt2.place(x=250, y=10, height=30, width=90)

    btt3 = tk.Button(root, text="16-24", command=sixteen_twentyfour, bg="green", fg="white", bd=5)
    btt3.place(x=350, y=10, height=30, width=90)

    btt4 = tk.Button(root, text="00-06", command=zero_six, bg="green", fg="white", bd=5)
    btt4.place(x=450, y=10, height=30, width=90)

    btt5 = tk.Button(root, text="00-08", command=zero_eight, bg="green", fg="white", bd=5)
    btt5.place(x=550, y=10, height=30, width=90)
    btt5 = tk.Button(root, text="06-08", command=six_eight, bg="green", fg="white", bd=5)
    btt5.place(x=650, y=10, height=30, width=90)
    btt6 = tk.Button(root, text="set", command=set, bg="black", fg="white", bd=5)
    btt6.place(x=750, y=10, height=30, width=90)

    powrot_btt_a = tk.Button(root, text=("back"), command=back_command, bg="red", fg="black")
    powrot_btt_a.place(x=10, y=10)
def main():

    name_f1 = tk.StringVar()
    name_f1.set(" ")
    name_f1_print = tk.Label(root, textvariable=name_f1, font=("Times New Roman", 15))
    name_f1_print.place(x=1, y=1,  height=700, width=1500)

    btt11 = tk.Button(root, text="report", command=report, bg="green", fg="white", bd=5)
    btt11.place(x=250, y=10, height=30, width=90)

    btt12 = tk.Button(root, text="check loc", command=loc_checks, bg="green", fg="white", bd=5)
    btt12.place(x=350, y=10, height=30, width=90)

root = tk.Tk()
root.geometry("1000x950")
root.title("Cycle Counting Productivity - TEST")
root.resizable(True, True)

main()
root.mainloop()

# filtr_user = (df_cc['User code 1st count'] == user)
# df_cc_user = df_cc[filtr_user]
# df_cc_user = df_cc_user.drop_duplicates(subset='Location from')
#
# filtr_czas_min = (df_cc_user["Start time 1st count"] >= datetime.time(h_min, m_min, 0)).dropna()
# filtr_czas_max = (df_cc_user["End time 1st count"] < datetime.time(h_max , m_max, 0)).dropna()
# df_cc_user = df_cc_user[filtr_czas_min & filtr_czas_max]
#
# start_time = df_cc_user["Start time 1st count"].iloc[0]
# stop_time = df_cc_user["End time 1st count"].iloc[-1]
#
# date = datetime.date(1, 1, 1)
#
#
# datetime1 = datetime.datetime.combine(date, start_time)
# datetime2 = datetime.datetime.combine(date, stop_time)
# full_time = datetime2 - datetime1
# full_time = str(full_time)
# (h, m, s) = full_time.split(':')
# full_time = int(h) * 3600 + int(m) * 60 + int(s)
#
# if break_time == 0:
#     full_time = (round(full_time / 3600, 2))
# else:
#
#     full_time = (round(full_time / 3600, 2))
#     full_time = (round(full_time - 0.5,2))
#
# print(full_time)
# if full_time < 0.2:
#     continue
# else:
#     df_cc_user_pivot = df_cc_user.pivot_table(index=['User code 1st count'],
#                                               values='Location from',
#                                               aggfunc='count')