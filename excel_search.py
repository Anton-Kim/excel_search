import re
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import xlwings as xw
from idlelib.tooltip import Hovertip

cyrillic_str_lower = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя'


def choose_file():
    global filepath
    filepath = filedialog.askopenfilename(
        filetypes=(('Книга Excel', '*.xls;*.xlsx'), ('Все файлы', '*.*')))
    file_name = filepath.split('/')[-1]
    file_name = file_name if len(file_name) < 26 else file_name[:23] + '...'
    lbl_choosen_file.configure(text=file_name)
    lbl_choosen_file.configure(foreground='black')


def color_on():
    lbl_choose_color['foreground'] = 'black'
    rad_color_green['state'] = 'normal'
    rad_color_yellow['state'] = 'normal'
    rad_color_red['state'] = 'normal'
    messagebox.showwarning(title='Предупреждение',
                           message='После запуска программы в файле будут '
                                   'автоматически окрашены найденные ячейки '
                                   'без возможности отменить изменения. '
                                   'Создайте копию файла при необходимости.')


def color_off():
    lbl_choose_color['foreground'] = 'gray'
    rad_color_green['state'] = 'disabled'
    rad_color_yellow['state'] = 'disabled'
    rad_color_red['state'] = 'disabled'


def expression_on():
    lbl_delimiter['foreground'] = 'black'
    ent_delimiter['state'] = 'normal'
    lbl_start_exp['foreground'] = 'black'
    ent_start_exp['state'] = 'normal'
    lbl_finish_exp['foreground'] = 'black'
    ent_finish_exp['state'] = 'normal'


def expression_off():
    lbl_delimiter['foreground'] = 'gray'
    ent_delimiter['state'] = 'disabled'
    lbl_start_exp['foreground'] = 'gray'
    ent_start_exp['state'] = 'disabled'
    lbl_finish_exp['foreground'] = 'gray'
    ent_finish_exp['state'] = 'disabled'


def check_fields(path, list_num, sch_txt, rng, vals):
    if not all((path, list_num, sch_txt, rng, vals)):
        messagebox.showwarning(title='Предупреждение',
                               message='Не все обязательные поля заполнены.')
    elif any(map(lambda x: x.lower() in cyrillic_str_lower, rng)) or any(
            map(lambda x: x.lower() in cyrillic_str_lower, vals)):
        messagebox.showwarning(title='Предупреждение',
                               message='В полях с указанием колонок '
                                       'присутствуют кириллические символы.')
    else:
        return True


def search():
    if check_fields(filepath, ent_list_num.get(), ent_search.get(),
                    ent_range.get(), ent_values.get()):
        res, reg = [], r'([A-Z]+)(\d+):\1(\d+)'
        re.match(r'\b[A-Z]+\b', ent_range.get())
        letter_range = re.match(r'\b[A-Z]+\b', ent_range.get())
        complex_range = re.match(reg, ent_range.get())
        values_col = re.match(r'\b[A-Z]+\b', ent_values.get())
        if not (letter_range or complex_range):
            messagebox.showwarning(title='Предупреждение',
                                   message='Некорректно задана колонка поиска.')
        elif complex_range:
            start_n = int(re.match(reg, ent_range.get()).group(2))
            finish_n = int(re.match(reg, ent_range.get()).group(3))
            if start_n <= finish_n:
                messagebox.showwarning(title='Предупреждение',
                                       message='Некорректно задан интервал '
                                               'поиска.')
        elif not values_col:
            messagebox.showwarning(title='Предупреждение',
                                   message='Некорректно задана колонка со '
                                           'значениями.')
        else:
            wb = xw.Book(filepath)
            sht = wb.sheets[int(ent_list_num.get()) - 1]
            if letter_range:
                counter, inf, file_length = 0, int(float('inf')), 1
                for r in sht.range(f'{letter_range}1:{letter_range}{inf}'):
                    if counter > 100:
                        break
                    else:
                        if not r.value:
                            file_length += 1
                            counter += 1
                        else:
                            file_length += 1
                            counter = 0
            span = complex_range if complex_range else f'{ent_range}1:{ent_range}{file_length}'
            for r in sht.range(span):
                if (s := r.value) and search == s:
                    if is_color:
                        r.is_color = tuple(int(i) for i in color.get().split(','))
                    res.append(f'{values_col}{r.row}')
            if color:
                wb.save()
            wb.app.quit()
        if is_expression:
            return ent_start_exp.get() + ent_delimiter.get().join(res) + ent_finish_exp.get()
        else:
            return ', '.join(res)


window = Tk()
window.title('Поиск ячеек по Excel файлу')
win_width = 455
win_height = 335
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x_win_coord = (screen_width // 2 - win_width // 2) - 10
y_win_coord = (screen_height // 2 - win_height // 2) - 10
window.geometry(f'{win_width}x{win_height}+{x_win_coord}+{y_win_coord}')
window.resizable(False, False)

frame = Frame(window, padx=10, pady=10)
frame.pack(expand=True)

filepath = ''
search_type = StringVar(value='in')
is_color = IntVar(value=0)
color = StringVar(value='146,208,80')
is_expression = IntVar(value=0)

img_color_green = PhotoImage(file='green.png')
img_color_yellow = PhotoImage(file='yellow.png')
img_color_red = PhotoImage(file='red.png')
tip = PhotoImage(file='tip.png')
cvs_search = Canvas(frame, height=17, width=17)
tip_search = cvs_search.create_image(2, 2, anchor=NW, image=tip)
cvs_range = Canvas(frame, height=17, width=17)
tip_range = cvs_range.create_image(2, 2, anchor=NW, image=tip)
cvs_values = Canvas(frame, height=17, width=17)
tip_values = cvs_values.create_image(2, 2, anchor=NW, image=tip)
cvs_expression = Canvas(frame, height=17, width=17)
tip_expression = cvs_expression.create_image(2, 2, anchor=NW, image=tip)

lbl_file = Label(frame, text='Выберите файл Excel  ')
btn_file = Button(frame, text='Выбрать', command=choose_file, width=10)
lbl_choosen_file = Label(frame, text='файл не выбран', foreground='gray')
lbl_list_num = Label(frame, text='Номер листа в книге  ',)
ent_list_num = Entry(frame, width=15)
lbl_list_num_tip = Label(frame, text='1 - это Лист 1 или название листа', foreground='gray')
ent_list_num.insert(END, '1')
lbl_search = Label(frame, text='Искомый текст  ')
ent_search = Entry(frame, width=47)
Hovertip(cvs_search, 'В тексте: в любом месте текста ячейки\n'
                     'Точное совпадение: совпадает с текстом в ячейке\n'
                     'Начинается: с искомого начинается текст в ячейке\n'
                     'Заканчивается: искомым заканчивается текст в ячейке')
rad_in = Radiobutton(frame, text='В тексте', value='in', variable=search_type)
rad_exact = Radiobutton(frame, text='Точное совпадение', value='exact', variable=search_type)
rad_startswith = Radiobutton(frame, text='Начинается', value='startswith', variable=search_type)
rad_endswith = Radiobutton(frame, text='Заканчивается', value='endswith', variable=search_type)
lbl_range = Label(frame, text='Колонка поиска  ')
ent_range = Entry(frame, width=15)
Hovertip(cvs_range, 'Буква колонки или диапазон ячеек в колонке латинскими\n'
                    'буквами, например B1:B176')
lbl_values = Label(frame, text='Колонка со значениями  ')
ent_values = Entry(frame, width=15)
Hovertip(cvs_values, 'Буква колонки со значениями, ячейка которой при\n'
                     'нахождении искомогого текста в колонке поиска, будет\n'
                     'добавлена к результату')
lbl_color = Label(frame, text='Окрасить ячейки с найденными совпаденями?  ')
rad_no_color = Radiobutton(frame, text='Нет', value=0, variable=is_color, command=color_off)
rad_color = Radiobutton(frame, text='Да', value=1, variable=is_color, command=color_on)
lbl_choose_color = Label(frame, text='Цвет:  ', foreground='gray')
rad_color_green = Radiobutton(frame, image=img_color_green, value='146,208,80', variable=color, state='disabled')
rad_color_yellow = Radiobutton(frame, image=img_color_yellow, value='255,255,0', variable=color, state='disabled')
rad_color_red = Radiobutton(frame, image=img_color_red, value='255,0,0', variable=color, state='disabled')
lbl_expression = Label(frame, text='Составить выражение из результата?  ')
rad_no_expression = Radiobutton(frame, text='Нет', value=0, variable=is_expression, command=expression_off)
rad_expression = Radiobutton(frame, text='Да', value=1, variable=is_expression, command=expression_on)
Hovertip(cvs_expression, 'По умолчанию найденные ячейки будут выведены через\n'
                         'запятые. С выражением можно изменить вывод.\n'
                         'Разделитель: символ будет вставлен между ячейками\n'
                         'В начало: будет "прилеплено" в начало\n'
                         'В конец: будет "прилеплено" в конец')
lbl_delimiter = Label(frame, text='Разделитель  ', foreground='gray')
ent_delimiter = Entry(frame, width=5, state='disabled')
lbl_start_exp = Label(frame, text='В начало  ', foreground='gray')
ent_start_exp = Entry(frame, width=12, state='disabled')
lbl_finish_exp = Label(frame, text='В конец  ', foreground='gray')
ent_finish_exp = Entry(frame, width=12, state='disabled')
lbl_warning = Label(frame, text='Все открытые окна Excel будут закрыты!', foreground='red3')
btn_search = Button(frame, text='Начать', command=search)

lbl_file.grid(row=1, column=1, sticky=W)
btn_file.grid(row=1, column=2, sticky=W)
lbl_choosen_file.grid(row=1, column=2, sticky=W, padx=(87, 0))
lbl_list_num.grid(row=2, column=1, sticky=W)
ent_list_num.grid(row=2, column=2, pady=5, sticky=W)
lbl_list_num_tip.grid(row=2, column=2, sticky=W, padx=(97, 0))
lbl_search.grid(row=3, column=1, sticky=W)
ent_search.grid(row=3, column=2, sticky=W)
cvs_search.grid(row=4, column=2, sticky=E)
rad_in.grid(row=4, column=1, columnspan=2, sticky=W)
rad_exact.grid(row=4, column=1, columnspan=2, sticky=W, padx=(72, 0))
rad_startswith.grid(row=4, column=1, columnspan=2, sticky=W, padx=(207, 0))
rad_endswith.grid(row=4, column=1, columnspan=2, sticky=W, padx=(299, 0))
lbl_range.grid(row=5, column=1, pady=5, sticky=W)
ent_range.grid(row=5, column=2, sticky=W)
cvs_range.grid(row=5, column=2, sticky=W, padx=(105, 0))
lbl_values.grid(row=6, column=1, sticky=W)
ent_values.grid(row=6, column=2, sticky=W)
cvs_values.grid(row=6, column=2, sticky=W, padx=(105, 0))
lbl_color.grid(row=7, column=1, columnspan=2, sticky=W, pady=5)
rad_no_color.grid(row=7, column=1, columnspan=2, sticky=W, padx=(270, 0))
rad_color.grid(row=7, column=1, columnspan=2, sticky=W, padx=(322, 0))
lbl_choose_color.grid(row=8, column=1, columnspan=2, sticky=W)
rad_color_green.grid(row=8, column=1, columnspan=2, sticky=W, padx=(45, 0))
rad_color_yellow.grid(row=8, column=1, columnspan=2, sticky=W, padx=(150, 0))
rad_color_red.grid(row=8, column=1, columnspan=2, sticky=W, padx=(255, 0))
lbl_expression.grid(row=9, column=1, columnspan=2, sticky=W, pady=5)
rad_no_expression.grid(row=9, column=1, columnspan=2, sticky=W, padx=(215, 0))
rad_expression.grid(row=9, column=1, columnspan=2, sticky=W, padx=(265, 0))
cvs_expression.grid(row=9, column=1, columnspan=2, sticky=W, padx=(310, 0))
lbl_delimiter.grid(row=10, column=1, columnspan=2, sticky=W)
ent_delimiter.grid(row=10, column=1, columnspan=2, sticky=W, padx=(80, 0))
lbl_start_exp.grid(row=10, column=1, columnspan=2, sticky=W, padx=(140, 0))
ent_start_exp.grid(row=10, column=1, columnspan=2, sticky=W, padx=(200, 0))
lbl_finish_exp.grid(row=10, column=1, columnspan=2, sticky=W, padx=(300, 0))
ent_finish_exp.grid(row=10, column=1, columnspan=2, sticky=E)
btn_search.grid(row=11, column=1, columnspan=2, sticky=W, padx=(5, 0), pady=(20, 5))
lbl_warning.grid(row=11, column=1, columnspan=2, sticky=W, padx=(65, 0), pady=(20, 5))

window.mainloop()
