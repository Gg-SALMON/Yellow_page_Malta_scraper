import tkinter.messagebox

import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import re
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import customtkinter
import os
from datetime import datetime
import openpyxl

#
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

path = os.getcwd()


def decode_email(e):
    de = ""
    k = int(e[:2], 16)
    for i in range(2, len(e)-1, 2):
        de += chr(int(e[i:i+2], 16)^k)
    return de


def format_address(address):
    return "-".join(address)


def list_0(list_):
    return list_[0]


def list_1(list_):
    return list_[-1]


def convert_to_list_1(string):
    return list_1(string.split(','))


def convert_to_list_0(string):
    return list_0(string.split(','))


def convert_address(x):
    return ", ".join(x)


def get_number_of_result(url, nb=60):
    response = requests.get(url, verify=False)
    soup = BeautifulSoup(response.content, "html.parser")
    number_result = soup.find('h1', attrs={'class': 'h6'}).text
    number_result = int(re.findall(r'\d+', number_result)[0])
    return number_result


def get_number_of_page(url, nb=60):
    n = 1
    response = requests.get(url, verify=False)
    soup = BeautifulSoup(response.content, "html.parser")
    try:
        number_result = soup.find('h1', attrs={'class': 'h6'}).text
        number_result = int(re.findall(r'\d+', number_result)[0])
    except:
        try:
            number_result = soup.find('p', attrs={'class': 'strong small lighter'}).text
            number_result = int(re.findall(r'\d+', number_result)[-1])
        except:
            print('Unable to find number of page')
            return 1

    if number_result % nb == 0:
        n = 0
    number_of_page = number_result // 60 + n
    print(number_of_page)
    return number_of_page


def companies_url(url):
    response = requests.get(url, verify=False)
    soup = BeautifulSoup(response.content, "html.parser")

    list_url = []
    for i in soup.find_all('h2', attrs={'class': 'h4'}):
        list_url.append("https://www.yellow.com.mt/" + i.a['href'])

    return list_url


def get_companies_url_all_pages(kw):
    list_url = []
    url_kw = kw.lower().replace(" ", "-")
    url = "https://www.yellow.com.mt/?search=" + url_kw
    n = get_number_of_page(url, nb=60)
    page = 1
    for i in range(1, n + 1):
        print(f"page {page} / {n}")
        list_url += (companies_url(url + "&pageno=" + str(i)))
        page += 1
    return list_url


def get_companies_url_all_pages_categories(kw):
    list_url = []
    url_kw = kw.lower().replace(" ", "-")
    url = f"https://www.yellow.com.mt/{url_kw}/"
    n = get_number_of_page(url, nb=60)
    page = 1
    for i in range(1, n + 1):
        print(f"page {page}")
        list_url += (companies_url(url + "&pageno=" + str(i)))
        page += 1
    return list_url


def get_info_from_website(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    company = soup.find('div', attrs={'class': 'cover-content hidden-mobile'})
    company_address = soup.find('div', attrs={'class': 'profile-contact'})
    phone_number = "Not found"
    if not company:
        return "not found", "not found", "not found", "not found", "not found"
    try:
        ca = company_address.address.find_all('span')
        address = [n.text for n in ca[0:2]]
        for n in [n.text for n in ca]:
            if n.replace(" ", "").isdigit():
                phone_number = n

    except:
        address = "Not found"
        phone_number = "Not found"

    name = company.h1.text
    title = company.address.text.strip()

    try:
        email = decode_email(company_address.address.find_all('span')[-1]['data-cfemail'])
    except:
        email = "Not found"

    return name, title, address, phone_number, email


def get_all_category(list_kw, kwd="yellow_page"):
    data = []
    m = 1
    for kw in list_kw:

        url_list = get_companies_url_all_pages(kw)
        n = 0

        for url in url_list:
            n += 1
            print(f"{n} / {len(url_list)} - {kw}")
            print(url)
            status.configure(text=f"{m}) {kw.title()} - {n} / {len(url_list)}")
            status2.configure(text=url)
            window.update()

            data_dict = dict([])
            name, title, address, phone_number, email = get_info_from_website(url)

            data_dict['name'] = name
            data_dict['title'] = title
            data_dict['key_word'] = kw
            data_dict['address'] = address
            data_dict['phone_number'] = phone_number
            data_dict['email'] = email
            data_dict['website'] = url
            data.append(data_dict)
        m += 1
    df = pd.DataFrame(data)
    df['title'] = (df.title.apply(convert_to_list_0))
    df['address'] = df.address.apply(convert_address)
    df.rename(str.capitalize, axis='columns', inplace=True)
    df.Phone_number = df.Phone_number.str.replace(" ", "")
    #df.drop(['Class'], axis=1, inplace=True)
    df = df[~(df.Name == "not found")]
    df.sort_values(by=(["Title", "Name"]))
    df.drop_duplicates(["Title", "Name", "Email", "Phone_number"], inplace=True)
    if file_combo.get()=="csv file":
        df.to_csv(f"{path}/{kwd}.csv", encoding='utf-8', index=False)
    else:
        df.to_excel(f"{path}/{kwd}.xlsx", index=False)

    messagebox.showinfo("Success", "Scrapping completed")
    status.configure(text="")
    status2.configure(text="")
    window.update()


def clear_default(event):
    event.widget.delete(0, 'end')
    event.widget.unbind('<FocusIn>')


def select_directory():
    global path
    folder_selected = filedialog.askdirectory()
    path = folder_selected
    print(path)


def quit_window():
    window.destroy()


def scrap():
    kw = input_kw.get()
    file_name = input_file.get()
    time = "".join([x for x in str(datetime.now()) if x.isdigit()][:14])
    if file_name == "":
        file_name = "Yellow_page"+time
    list_kw = kw.split(",")
    get_all_category(list_kw, file_name)

def create_dataframe(file,path):
    if file.endswith('.csv'):
        return pd.read_csv(path+"/"+file)
    elif file.endswith('.xlsx'):
        return pd.read_excel(path+"/"+file)


def merge_files():
    excel_file = filedialog.asksaveasfilename(initialdir=(path),title="Select file",
                                              filetypes=[('XLS files', '*.xlsx')], defaultextension=".xlsx")

    filenames = [file for file in os.listdir(path) if (file.endswith('.csv') or file.endswith('.xlsx'))]
    df = [create_dataframe(file, path) for file in filenames]
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        for i in range(len(filenames)):
            df[i].to_excel(writer, sheet_name=f"{filenames[i].split('.')[0]}", index=False)
    print(excel_file)
    messagebox.showinfo(title="New Excel file created!", message=f"All excel and csv filr from {path} have been merged into {excel_file}")


window = customtkinter.CTk()

window.title("SCRAP YELLOW PAGE")

window.geometry('700x200')
window.resizable(width=False, height=False)
window.configure(padx=1, pady=1)


frame1 = customtkinter.CTkFrame(window, width=550, height=175, corner_radius=1)
frame1.grid_propagate(False)
frame1.grid(row=1, column=0)


frame2 = customtkinter.CTkFrame(window, width=150, height=175, corner_radius=1)
frame2.pack_propagate(False)
frame2.grid(row=1, column=1)

frame3 = customtkinter.CTkFrame(window, width=700, height=25, corner_radius=1)
frame3.grid_propagate(False)
frame3.grid(row=2, column=0,columnspan=2)

status = customtkinter.CTkLabel(frame3, text="", anchor=W, font=("helvetica", 10), height=8)
status.grid(row=0, column=0, sticky=EW)

status2 = customtkinter.CTkLabel(frame3, text="", justify=LEFT, anchor=W,  font=("helvetica", 10), height=8)
status2.grid(row=1, column=0, sticky=EW)


label1 = customtkinter.CTkLabel(
    frame1, text="Insert keywords to research separated by coma",
    font=("helvetica", 14, "bold"), justify=LEFT, width=548, anchor=W )
label1.grid(row=0, column=0, sticky=EW, columnspan=2, pady=(0,0))
#label1.pack(padx=1, pady=0, anchor=W)

input_kw = customtkinter.CTkEntry(frame1, width=540, corner_radius=50)
input_kw.grid(row=1, column=0, sticky=EW, columnspan=2, pady=(0,0))

label2 = customtkinter.CTkLabel(
    frame1,text='For example, "Building Contractors,Concrete Building Blocks & Roofing Panels"  ',
    font=("helvetica", 11, 'italic'), text_color="SkyBlue4",justify=LEFT, anchor=W)

label2.grid(row=2, column=0, sticky=EW, columnspan=2)

label3 = customtkinter.CTkLabel(frame1, text="Insert file name", font=("helvetica", 14, "bold"), justify=LEFT, anchor=W)
label3.grid(row=3, column=0, sticky=EW, columnspan=2)

input_file = customtkinter.CTkEntry(frame1, width=280, corner_radius=50)
input_file.grid(row=4, column=0, sticky=EW)

file_type = ["csv file", "Excel file"]
file_combo = customtkinter.CTkComboBox(frame1, values=file_type)
file_combo.grid(row=4, column=1, sticky=EW, padx=(5,0))

label4 = customtkinter.CTkLabel(
    frame1, text='For example, "Building_Contractors"  ',
    font=("helvetica", 11, 'italic'), text_color="SkyBlue4", anchor=W)
label4.grid(row=5, column=0, sticky=EW, columnspan=2)


button_pady = 1

button_path = customtkinter.CTkButton(
    frame2, text="Select directory",
    command=select_directory, width=140, height=40, corner_radius=50)
button_path.pack(padx=1, pady=(10, button_pady))

button_download = customtkinter.CTkButton(
    frame2, text="Download",command=scrap, width=140, height=40, corner_radius=50)
button_download.pack(padx=1, pady=button_pady)

button_merge = customtkinter.CTkButton(
    frame2, text="Merge files",command=merge_files, width=140, height=40, corner_radius=50)
button_merge.pack(padx=1, pady=button_pady)


button_quit = customtkinter.CTkButton(
    frame2, text="Quit", command=quit_window,  width=140, height=40, corner_radius=50)
button_quit.pack(padx=1, pady=button_pady)

window.mainloop()


