import tkinter as tk
from tkinter import filedialog
from tkinter import ttk, messagebox, scrolledtext
import threading
import os
import time
import readOF
import config_for_interface
import database
import file_transfer


def get_folder_click():
    """Функция описывает логику работу кнопки выбора директории для выгрузки ОФ"""
    config_for_interface.path_to_folder = filedialog.askdirectory()
    print("Selected folder:", config_for_interface.path_to_folder)
    if not file_transfer.check_folder_writable(config_for_interface.path_to_folder):
        messagebox.showerror("Ошибка", "Выбранная папка недоступна для записи! Выберите другую папку")
        return
    if config_for_interface.path_to_folder:
        button2.configure(state="normal")
        button2.configure(bg="#1166EE")
        button1.configure(bg="#118844")

def get_excle_folder_click():
    """Получает обменные формы для внесения факта в project"""
    directory = filedialog.askdirectory()
    file_paths = []
    for root, directories, files in os.walk(directory):
        for file in files:
            file_path = os.path.abspath(os.path.join(root, file))
            file_paths.append(file_path)
    config_for_interface.path_to_of = file_paths
    print("Selected file:", config_for_interface.path_to_of)
    if not file_transfer.check_folder_writable(config_for_interface.path_to_reserv_folder):
        messagebox.showerror("Ошибка", "Выбранная временная папка недоступна для записи! Выберите другую папку")
        return
    if config_for_interface.path_to_of:
        res = file_transfer.transfer_files(config_for_interface.path_to_of,
                                           config_for_interface.path_to_reserv_folder)
        if res is not False:
            messagebox.showerror("Ошибка", f"Файл project {res} недоступен для чтения! Выберите другие файлы")
            return
        button6.configure(state="normal")
        button6.configure(bg="#1166EE")
        button5.configure(bg="#118844")


def get_project_click():
    """Функция описывает логику работу кнопки выбора файлов projects для выгрузки в ОФ"""
    directory = filedialog.askdirectory()
    file_paths = []
    for root, directories, files in os.walk(directory):
        for file in files:
            file_path = os.path.abspath(os.path.join(root, file))
            file_paths.append(file_path)
    config_for_interface.path_to_project = file_paths
    print("Selected file:", config_for_interface.path_to_project)
    if not file_transfer.check_folder_writable(config_for_interface.path_to_reserv_folder):
        messagebox.showerror("Ошибка", "Выбранная временная папка недоступна для записи! Выберите другую папку")
        return
    if config_for_interface.path_to_project:
        res = file_transfer.transfer_files(config_for_interface.path_to_project,
                                           config_for_interface.path_to_reserv_folder)
        if res is not False:
            messagebox.showerror("Ошибка", f"Файл project {res} недоступен для чтения! Выберите другие файлы")
            return
        start_button.configure(state="normal")
        start_button.configure(bg="#1166EE")
        button2.configure(bg="#118844")
        button6.configure(bg="#118844")
        start_button_for_fact.configure(state="normal")
        start_button_for_fact.configure(bg="#1166EE")

def open_folder_with_OF():
    if config_for_interface.path_to_folder:
        os.startfile(config_for_interface.path_to_folder)

def open_reserv_folder():
    if config_for_interface.path_to_reserv_folder:
        os.startfile(config_for_interface.path_to_reserv_folder)
        
def update_progress(value):
    """Функция обновляет значение количества загруженных файлов"""
    percent_label.configure(text=f"Загружено: {value} файлов из {len(config_for_interface.path_to_project)}")


def switch_info_labels(value):
    """Функция выводит текстовую информацию о результате загрузки"""
    succes = sum(1 for item in config_for_interface.path_to_excel if item is not None)
    if value == 0:
        info_label.configure(text="Пожалуйста, ожидайте, выгрузка ОФ может занимать длительное время")
    else:
        info_label.configure(
            text=f"Загрузка завершена. Загружено: {len(config_for_interface.path_to_excel)} файлов.\n Успешно: {succes}")


def change_after_work(value):
    """Функция задает изменения в стилях после окончания выгрузки обменных форм"""
    percent_label.place_forget()
    switch_info_labels(value)
    info_label.place_configure(relx=0.025, rely=0.5)
    start_button.configure(bg="#118844")
    button3.configure(state="normal", bg="#1166EE")
    button4.configure(state="normal", bg="#1166EE")

def start_progress():
    """Функция выполняет выгрузку выбранных project в ОФ"""
    percent_label.place(relx=0.025, rely=0.5)
    info_label.place(relx=0.025, rely=0.55)
    value = 0
    switch_info_labels(value)
    update_progress(value)
    window.update()
    path_to_reserv_folder_for_of = config_for_interface.path_to_reserv_folder + '\\' + "OF"
    paths_to_bad_projects = list()
    path_to_folder_bad_project = config_for_interface.path_to_reserv_folder + '\\' + "bad"
    if not os.path.exists(path_to_reserv_folder_for_of):
        os.mkdir(path_to_reserv_folder_for_of)

    for path in config_for_interface.path_to_project:
        file_name = os.path.basename(path)
        path = os.path.join(config_for_interface.path_to_reserv_folder, file_name)
        config_for_interface.path_to_excel.append(readOF.main(path, path_to_reserv_folder_for_of))
        if config_for_interface.path_to_excel[-1] is not None:
            text_area.insert(tk.INSERT,
                             f"{os.path.basename(config_for_interface.path_to_project[value])}    -    Успешно\n")
        else:
            if not os.path.exists(path_to_folder_bad_project):
                os.mkdir(path_to_folder_bad_project)
            paths_to_bad_projects.append(path)
            text_area.insert(tk.INSERT,
                             f"{os.path.basename(config_for_interface.path_to_project[value])}    -    Не успешно\n")
        value = value + 1
        update_progress(value)
        window.update()
    if paths_to_bad_projects:
        file_transfer.transfer_files(paths_to_bad_projects, path_to_folder_bad_project)
    file_transfer.transfer_files(config_for_interface.path_to_excel, config_for_interface.path_to_folder)
    change_after_work(value)
    database.fill_data(config_for_interface.path_to_excel, config_for_interface.path_to_project,
                       config_for_interface.path_to_folder)


def on_window_resize(event):
    """Обработчик события изменения размеров окна"""
    new_width = window.winfo_width()
    new_height = window.winfo_height()

    button_width = int(new_width / 7)
    button_height = int(new_height / 15)
    label_width = int(new_width / 1.2)
    label_height = int(new_height / 15)
    text_area_height = int(new_height / 3)
    # Обновляем ширину кнопок
    button1.place(width=button_width, height=button_height)
    button2.place(width=button_width, height=button_height)
    start_button.place(width=button_width, height=button_height)
    button3.place(width=button_width*1.17, height=button_height)
    button4.place(width=button_width*1.4, height=button_height)
    label1.place(width=label_width, height=label_height)
    label2.place(width=label_width, height=label_height)
    label3.place(width=label_width, height=label_height)
    text_area.place(width=label_width / 1.3, height=text_area_height)

def on_tab_selected(event):
    """Обработчик переключения вкладок приложения"""
    selected_tab = notebook.index(notebook.select())
    if selected_tab == 0:
        button1.config(state="normal")
        button2.config(state="disabled")
    elif selected_tab == 1:
        button1.config(state="disabled")
        button2.config(state="normal")

if __name__ == '__main__':
    window = tk.Tk()
    database.create_database()
    window.title("Выгрузка ОФ")
    window.geometry("1000x600")
    window.wm_minsize(1000, 600)
    window.configure(background="light blue")
    notebook = ttk.Notebook(window)
    notebook.bind("<<NotebookTabChanged>>", on_tab_selected)
    style = ttk.Style()
    style.configure("TNotebook", background="blue")
    style.configure("TFrame", background="light blue")
    button_style_active = {'background': '#1166EE', 'foreground': 'white', 'font': ('Arial', 12)}
    button_style_done = {'background': '#118844', 'foreground': 'white', 'font': ('Arial', 12)}
    button_style_block = {'background': '#969699', 'foreground': 'white', 'font': ('Arial', 12)}
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Вкладка 1")
    label0 = tk.Label(tab1, text="Эта программа предназначена для выгрузки обменных форм из файлов project в папку",
                      font=('Arial', 14), background="light blue")
    label0.pack()
    button1 = tk.Button(tab1, text="Выбрать папку", command=get_folder_click, **button_style_active, width=15)
    button1.place(relx=0.025, rely=0.17)
    label1 = tk.Label(tab1, text="Эта кнопка позволяет выбрать папку, в которую нужно выгрузить обменные формы",
                      font=('Arial', 14), anchor='w', background="light blue")
    label1.place(relx=0.2, rely=0.17)
    button2 = tk.Button(tab1, text="Выбрать папку", command=get_project_click, **button_style_block, state="disabled",
                        width=15)
    button2.place(relx=0.025, rely=0.27)
    label2 = tk.Label(tab1, text="Эта кнопка позволяет выбрать папку с файлами project для выгрузки обменной формы",
                      font=('Arial', 14), anchor='w', background="light blue")
    label2.place(relx=0.2, rely=0.27)
    button3 = tk.Button(tab1, text="Открыть папку с ОФ", command=open_folder_with_OF, **button_style_block, state="disabled",
                        width=17)
    button3.place(relx=0.6, rely=0.5)
    button4 = tk.Button(tab1, text="Открыть резервную папку", command=open_reserv_folder, **button_style_block, state="disabled",
                        width=21)
    button4.place(relx=0.78, rely=0.5)
    percent_label = tk.Label(tab1, text="Выгружено: 0 файлов", font=('Arial', 12), background="light blue")
    label3 = tk.Label(tab1, text="Эта кнопка позволяет начать выгрузку обменных форм", font=('Arial', 14), anchor='w',
                      background="light blue")
    label3.place(relx=0.2, rely=0.37)
    info_label = tk.Label(tab1, text="Пожалуйста, ожидайте, выгрузка ОФ может занимать длительное время",
                          font=('Arial', 12), background="light blue")
    start_button = tk.Button(tab1, text="Начать выгрузку", command=start_progress, **button_style_block,
                             state="disabled", width=15)
    start_button.place(relx=0.025, rely=0.37)
    window.bind("<Configure>", on_window_resize)
    text_area = scrolledtext.ScrolledText(tab1, width=80, height=10)
    text_area.place(relx=0.17, rely=0.6)




    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Вкладка 2")

    button5 = tk.Button(tab2, text="Выбрать папку с ОФ", command=get_excle_folder_click, **button_style_active, width=20)
    button5.place(relx=0.025, rely=0.17)
    button6 = tk.Button(tab2, text="Выбрать папку с project", command=get_project_click, **button_style_block, width=20)
    button6.place(relx=0.025, rely=0.27)
    start_button_for_fact = tk.Button(tab2, text="Начать внесение", command=get_project_click, **button_style_block, width=20)
    start_button_for_fact.place(relx=0.025, rely=0.37)








    notebook.pack(fill=tk.BOTH, expand=True)

    messagebox.showwarning("Предупреждение",
                           "Пожалуйста, закройте открытые файлы project для корректной работы программы")
    window.mainloop()
