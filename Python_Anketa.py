###########################################IMPORT_MODULES###########################################################
import tkinter as tk  # импортирование модуля и назначение ему сокращённого имени для дальнешего краткого обращения
import os  # подключение модуля команд операционной системы

import openpyxl  # подключение модуля создания и редактирования табличных документов


##################################################CONSTANT_VARIABLES###################################################
# список значений записываемых в колонку А
txt2 = ["ФИО:",
        "Дата рождения:",
        "Род занятий:",
        "Как Вы попали в IT индустрию?:",
        "Почему решили заняться именно тем чем занимаетесь сейчас?:",
        "Какие Ваши увлечения?:",
        '''Какую музыку предпочитаете слушать?(жанры\группы\композиции):''',
        '''Какие фильмы предпочитаете смотреть?(жанры\конкретные наименования):''',
        '''Любите ли Вы играть в видеоигры?Если нет то почему?:''',
        '''Какие Ваши любимые видеоигры?:''',
        "Смотрите ли Вы аниме? если нет то почему?:",
        '''Какое аниме предпочитаете смотреть?:''',
        "Критика\Пожелания\Напутствие автору данной анкеты:"]




#####################################################FUNCTIONS#########################################################
# функция отслеживания ввода сиволов в поля для ввода и регулировки состояния виджета кнопки
def on_key(event):

    if len(name_input.get()) and len(date_born_input.get()) and len(work_input.get()) and len(how_find_this_work_input.get()) and len(why_is_this_job_input.get()) and len(hobby_input.get()) and len(
            favorite_music_input.get()) and len(favorite_movies_input.get()) and len(do_you_like_computer_games_input.get()) and len(favorite_computer_games_input.get()) and len(do_you_like_anime_input.get()) and len(
            favorite_anime_input.get()) and len(criticism_input.get()) > 0:
        action_button['state'] = 'normal'
    else:
        action_button['state'] = 'disabled'






# функция записи результатов опроса анкетирования в файл xlsx
def write_xlsx(txt):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet["A1"] = "Вопросы:"  # запись в ячейку А1
    sheet["B1"] = "Ответы:"  # запись в ячейку B1
    row = 2  # начать запись со строки 2(для цикла записи в колонку A)
    crow = 2  # начать запись со строки 2(для цикла записи в колонку B)
    for i in txt2:
        sheet[row][0].value = i  # запись в столбец А
        row += 1  # переход на следующую строку
    for j in txt:
        sheet[crow][1].value = j  # запись в стобец B
        crow += 1  # переход на следующую строку
    book.save(str(name_input.get())+" result_anketa.xlsx")  # сохранение всех ранее записанных изменений в файл
    book.close()  # закрытие файла






# функция удаления ранее созданных файлов результата опроса при запуске программы
def del_files():
    origfolder = os.getcwd()  # путь расположения файлов относительно расположения скрипта
    test = os.listdir(origfolder)  # "прослушивание" каталога со скриптом

    # цикл удаления файлов
    for item in test:
        if item.endswith(".xlsx"):  # если раширение файла .csv
            os.remove(os.path.join(origfolder,
                                   item))  # то по ранее указанному пути в  переменной origfolder будет удаляться файл с расширением указанным строчкой выше


# функция получения информации из виджетов entry(поле для ввода данных)
def get_entry():
    # переменная содержащая в себе получение данных из всех полей сразу
    v = name_input.get(), date_born_input.get(), work_input.get(), how_find_this_work_input.get(), why_is_this_job_input.get(), hobby_input.get(), favorite_music_input.get(), favorite_movies_input.get(), do_you_like_computer_games_input.get(), favorite_computer_games_input.get(), do_you_like_anime_input.get(), favorite_anime_input.get(), criticism_input.get()

    if v:
        # список данных для записи в ячейку B
        txt = [name_input.get(), date_born_input.get(), work_input.get(), how_find_this_work_input.get(), why_is_this_job_input.get(), hobby_input.get(), favorite_music_input.get(), favorite_movies_input.get(), do_you_like_computer_games_input.get(), favorite_computer_games_input.get(),
               do_you_like_anime_input.get(), favorite_anime_input.get(), criticism_input.get()]


        write_xlsx(txt)  # вызов функции записи введённой информации в xlsx

       


################################################MAIN_WINDOW############################################################
win = tk.Tk()  # создание переменной с вызовом  модуля окна
# icon_win = tk.PhotoImage(file= "workplace-09-512.png" )  # присваивание переменной импортированного внешнего файла фото
# win.iconphoto(False,icon_win)       # вызов изменённого лого окна
win.title("Анкета by popww")  # указание  названия окна
#win.geometry("645x360+500+200")  # указание размера окна и расположение его относительно левого верхнего угла
# win.minsize(320, 240)               # минимальный размер окна
# win.maxsize(1280,720)               # максимальный размер окна
win.resizable(False, False)  # Возможность изменять размер вручную

##############################################LABELES#################################################################
name_label = tk.Label(win, text="ФИО",anchor="center",font=("Arial", 10, "bold")).grid(row=0, column=0)

date_born_label = tk.Label(win, text="Дата рождения", font=("Arial", 10, "bold"), anchor="center").grid(row=1, column=0)

work_label = tk.Label(win, text="Род занятий", font=("Arial", 10, "bold")).grid(row=2, column=0)

how_find_this_work_label = tk.Label(win, text="Как Вы попали в IT индустрию?", font=("Arial", 10, "bold")).grid(row=3, column=0)

why_is_this_job_label = tk.Label(win, text="Почему решили заняться именно тем чем занимаетесь сейчас?", font=("Arial", 10, "bold")).grid(row=4, column=0)

hobby_label = tk.Label(win, text="Какие Ваши увлечения?", font=("Arial", 10, "bold")).grid(row=5, column=0)

favorite_music_label = tk.Label(win, text='''Какую музыку предпочитаете слушать?
(жанры\группы\композиции)''', font=("Arial", 10, "bold"), anchor="center").grid(row=6, column=0)


favorite_movies_label = tk.Label(win, text='''Какие фильмы предпочитаете смотреть?
(жанры\конкретные наименования)''', font=("Arial", 10, "bold")).grid(row=7, column=0)


do_you_like_computer_games_label = tk.Label(win, text='''Любите ли Вы играть в видеоигры?
Если нет то почему?''', font=("Arial", 10, "bold")).grid(row=8, column=0)

favorite_computer_games_label = tk.Label(win, text='''Какие Ваши любимые видеоигры?(жанры\конкретные примеры)
Если предыдущий ответ отрицательный то данный пункт можете пропустить''', font=("Arial", 8, "bold")).grid(row=9, column=0)


do_you_like_anime_label = tk.Label(win, text="Смотрите ли Вы аниме? если нет то почему?", font=("Arial", 10, "bold")).grid(row=10, column=0)


favorite_anime_label = tk.Label(win, text='''Какое аниме предпочитаете смотреть? (жанры\конкретные тайтлы)
Если предыдущий ответ отрицательный то данный пункт можете пропустить''', font=("Arial", 8, "bold")).grid(row=11, column=0)


criticism_label = tk.Label(win, text="Критика\Пожелания\Напутствие автору данной анкеты", font=("Arial", 10, "bold")).grid(row=12, column=0)


empty_label = tk.Label(text = "").grid(row = 13, column = 0)


###############################################BUTTONS#################################################################
action_button = tk.Button(win, text='''нажать после заполнения всех полей''',
                 anchor="center",
                 width=70,  # ширина виджета
                 height=2,  # высота виджета
                 font=("Arial", 14,),
                 relief=tk.RAISED,  # параметр выделение по контуру виджета
                 command=get_entry,
                 state="disabled")
action_button.grid(row=14, column=0,
                rowspan=13,

                columnspan=4)


#############################################INPUT_FIELDS###############################################################
name_input = tk.Entry(win, width = 50)  # создание виджета ввода текста с клавиатуры
name_input.bind('<KeyRelease>', on_key)
name_input.grid(row=0, column=1,pady = 10)

date_born_input = tk.Entry(win,width = 50 )
date_born_input.bind('<KeyRelease>', on_key)
date_born_input.grid(row=1, column=1, pady = 10)


work_input = tk.Entry(win,width = 50 )
work_input.bind('<KeyRelease>', on_key)
work_input.grid(row=2, column=1,pady = 10)

how_find_this_work_input = tk.Entry(win,width = 50 )
how_find_this_work_input.bind('<KeyRelease>', on_key)
how_find_this_work_input.grid(row=3, column=1,pady = 10)

why_is_this_job_input = tk.Entry(win,width = 50 )
why_is_this_job_input.bind('<KeyRelease>', on_key)
why_is_this_job_input.grid(row=4, column=1,pady = 10)

hobby_input = tk.Entry(win, width = 50)
hobby_input.bind('<KeyRelease>', on_key)
hobby_input.grid(row=5, column=1,pady = 5)

favorite_music_input = tk.Entry(win,width = 50 )
favorite_music_input.bind('<KeyRelease>', on_key)
favorite_music_input.grid(row=6, column=1,pady = 5)

favorite_movies_input = tk.Entry(win,width = 50 )
favorite_movies_input.bind('<KeyRelease>', on_key)
favorite_movies_input.grid(row=7, column=1)

do_you_like_computer_games_input = tk.Entry(win,width = 50 )
do_you_like_computer_games_input.bind('<KeyRelease>', on_key)
do_you_like_computer_games_input.grid(row=8, column=1)

favorite_computer_games_input = tk.Entry(win,width = 50 )
favorite_computer_games_input.bind('<KeyRelease>', on_key)
favorite_computer_games_input.grid(row=9, column=1)

do_you_like_anime_input = tk.Entry(win, width = 50)
do_you_like_anime_input.bind('<KeyRelease>', on_key)
do_you_like_anime_input.grid(row=10, column=1,pady = 10)

favorite_anime_input = tk.Entry(win,width = 50 )
favorite_anime_input.bind('<KeyRelease>', on_key)
favorite_anime_input.grid(row=11, column=1,pady = 10)


criticism_input = tk.Entry(win,width = 50 )
criticism_input.bind('<KeyRelease>', on_key)
criticism_input.grid(row=12, column=1,pady = 10)


####################################################ENTRY_POINT########################################################
if __name__ == '__main__' :
    del_files()
    win.mainloop()  # запуск окна в режиме ожидания команд от пользователя
