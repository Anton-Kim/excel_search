from tkinter import *
from tkinter import filedialog, messagebox, ttk
from pandas import read_excel
from xlwings import Book
from re import match
from idlelib.tooltip import Hovertip

cyrillic_str_lower = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя'

"""Вкладка поиска ячеек"""


def choose_file():
    global filepath
    filepath = filedialog.askopenfilename(
        filetypes=(('Книга Excel', '*.xls;*.xlsx;*.xlsm'), ('Все файлы', '*.*')))
    file_name = filepath.split('/')[-1]
    file_name = file_name if len(file_name) < 26 else file_name[:23] + '...'
    if file_name:
        lbl_choosen_file['text'] = file_name
        lbl_choosen_file['foreground'] = 'black'
    elif not file_name and lbl_choosen_file['foreground'] != 'gray':
        lbl_choosen_file['text'] = 'файл не выбран'
        lbl_choosen_file['foreground'] = 'gray'


def color_on():
    lbl_choose_color['foreground'] = 'black'
    rad_color_green['state'] = 'normal'
    rad_color_yellow['state'] = 'normal'
    rad_color_red['state'] = 'normal'
    rad_color_off['state'] = 'normal'
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
    rad_color_off['state'] = 'disabled'


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
        letter_range = match(r'\b[A-Z]+\b', ent_range.get())
        complex_range = match(reg, ent_range.get())
        values_col = match(r'\b[A-Z]+\b', ent_values.get())
        if not (letter_range or complex_range):
            messagebox.showwarning(title='Предупреждение',
                                   message='Некорректно задана колонка поиска.')
        elif complex_range and int(
                match(reg, ent_range.get()).group(2)) > int(
                match(reg, ent_range.get()).group(3)):
            messagebox.showwarning(title='Предупреждение',
                                   message='Некорректно задан интервал поиска.')
        elif not values_col:
            messagebox.showwarning(title='Предупреждение',
                                   message='Некорректно задана колонка со '
                                           'значениями.')
        else:
            final_range = letter_range.group() if letter_range else complex_range.group()
            values_col = values_col.group()
            wb = Book(filepath)
            try:
                if ent_list_num.get().isdigit():
                    sht = wb.sheets[int(ent_list_num.get()) - 1]
                else:
                    sht = wb.sheets[ent_list_num.get()]
                if letter_range:
                    file_length = len(read_excel(filepath)) + 3
                span = final_range if ':' in final_range else f'{letter_range.group()}1:{letter_range.group()}{file_length}'
                for r in sht.range(span):
                    if (s := r.value) and search_type_controller(search_type.get(), ent_search.get(), s):
                        if is_colorize.get():
                            if color.get() == 'None':
                                r.color = None
                            else:
                                r.color = tuple(int(i) for i in color.get().split(','))
                        res.append(f'{values_col}{r.row}')
                if is_colorize.get():
                    wb.save()
                wb.app.quit()
                if is_expression.get() and res:
                    final_res = ent_start_exp.get() + ent_delimiter.get().join(res) + ent_finish_exp.get()
                else:
                    final_res = ', '.join(res)
                if final_res:
                    txt_result.delete('1.0', END)
                    txt_result.insert(INSERT, chars=final_res)
                    btn_copy['text'] = 'скопировать в буфер'
                    btn_copy['state'] = 'normal'
                else:
                    txt_result.delete('1.0', END)
                    txt_result.insert(INSERT, chars='Совпадения не найдены.')
                    btn_copy['text'] = ''
                    btn_copy['state'] = 'disabled'
            except Exception as err:
                print(err)
                wb.app.quit()
                messagebox.showwarning(title='Ошибка',
                                       message='Что-то пошло не так...')


def search_type_controller(tp, txt_1, txt_2):
    str_txt_2 = str(txt_2)
    if tp == 'in' and txt_1 in str_txt_2:
        return True
    elif tp == 'exact' and txt_1 == str_txt_2:
        return True
    elif tp == 'startswith' and txt_1.startswith(str_txt_2):
        return True
    elif tp == 'endswith' and txt_1.endswith(str_txt_2):
        return True
    else:
        return False


def copy_to_clipboard():
    window.clipboard_clear()
    a = txt_result.get('1.0', END)[:-1]
    window.clipboard_append(a)


def fix_keyboard_shortcuts(event):
    ctrl = (event.state & 0x4) != 0
    if event.keycode == 88 and ctrl and event.keysym.lower() != 'x':
        event.widget.event_generate('<<Cut>>')
    if event.keycode == 86 and ctrl and event.keysym.lower() != 'v':
        event.widget.event_generate('<<Paste>>')
    if event.keycode == 67 and ctrl and event.keysym.lower() != 'c':
        event.widget.event_generate('<<Copy>>')
    if event.keycode == 65 and ctrl and event.keysym.lower() != 'a':
        event.widget.event_generate("<<SelectAll>>")


window = Tk()
window.title('Поиск по Excel файлу')
window.iconphoto(False, PhotoImage(file='win_icon.png'))
win_width = 455
win_height = 475
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x_win_coord = (screen_width // 2 - win_width // 2) - 10
y_win_coord = (screen_height // 2 - win_height // 2) - 10
window.geometry(f'{win_width}x{win_height}+{x_win_coord}+{y_win_coord}')
window.resizable(False, False)
window.bind_all('<Key>', fix_keyboard_shortcuts, '+')

notebook = ttk.Notebook()
notebook.pack(expand=True, fill=BOTH)

frame = Frame(notebook, padx=10, pady=10)
frame_uq = Frame(notebook, padx=10, pady=10)

notebook.add(frame, text='Поиск ячеек')
notebook.add(frame_uq, text='Поиск уникальных значений колонки')

filepath = ''
search_type = StringVar(value='exact')
is_colorize = IntVar(value=0)
color = StringVar(value='146,208,80')
is_expression = IntVar(value=1)

img_color_green = PhotoImage(file='green.png')
img_color_yellow = PhotoImage(file='yellow.png')
img_color_red = PhotoImage(file='red.png')
img_color_off = PhotoImage(file='no_color.png')
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
Hovertip(cvs_search, 'Точное совпадение: совпадает с текстом в ячейке\n'
                     'В тексте: в любом месте текста ячейки\n'
                     'Начинается: с искомого начинается текст в ячейке\n'
                     'Заканчивается: искомым заканчивается текст в ячейке')
rad_exact = Radiobutton(frame, text='Точное совпадение', value='exact', variable=search_type)
rad_in = Radiobutton(frame, text='В тексте', value='in', variable=search_type)
rad_startswith = Radiobutton(frame, text='Начинается', value='startswith', variable=search_type)
rad_endswith = Radiobutton(frame, text='Заканчивается', value='endswith', variable=search_type)
lbl_range = Label(frame, text='Колонка поиска  ')
ent_range = Entry(frame, width=15)
Hovertip(cvs_range, 'Буква колонки или диапазон ячеек в колонке латинскими\n'
                    'заглавными буквами, например B1:B176')
lbl_values = Label(frame, text='Колонка со значениями  ')
ent_values = Entry(frame, width=15)
Hovertip(cvs_values, 'Буква колонки со значениями, ячейка которой при\n'
                     'нахождении искомогого текста в колонке поиска, будет\n'
                     'добавлена к результату')
lbl_color = Label(frame, text='Окрасить ячейки с найденными совпаденями?  ')
rad_no_color = Radiobutton(frame, text='Нет', value=0, variable=is_colorize, command=color_off)
rad_color = Radiobutton(frame, text='Да', value=1, variable=is_colorize, command=color_on)
lbl_choose_color = Label(frame, text='Цвет:  ', foreground='gray')
rad_color_green = Radiobutton(frame, image=img_color_green, value='146,208,80', variable=color, state='disabled')
rad_color_yellow = Radiobutton(frame, image=img_color_yellow, value='255,255,0', variable=color, state='disabled')
rad_color_red = Radiobutton(frame, image=img_color_red, value='255,0,0', variable=color, state='disabled')
rad_color_off = Radiobutton(frame, image=img_color_off, value='None', variable=color, state='disabled')
lbl_expression = Label(frame, text='Составить выражение из результата?  ')
rad_no_expression = Radiobutton(frame, text='Нет', value=0, variable=is_expression, command=expression_off)
rad_expression = Radiobutton(frame, text='Да', value=1, variable=is_expression, command=expression_on)
Hovertip(cvs_expression, 'По умолчанию найденные ячейки будут выведены через\n'
                         'запятые. С выражением можно изменить вывод.\n'
                         'Разделитель: символ будет вставлен между ячейками\n'
                         'В начало: будет "прилеплено" в начало\n'
                         'В конец: будет "прилеплено" в конец')
lbl_delimiter = Label(frame, text='Разделитель  ')
ent_delimiter = Entry(frame, width=5)
ent_delimiter.insert(END, '+')
lbl_start_exp = Label(frame, text='В начало  ')
ent_start_exp = Entry(frame, width=12)
ent_start_exp.insert(END, '=')
lbl_finish_exp = Label(frame, text='В конец  ')
ent_finish_exp = Entry(frame, width=12)
btn_search = Button(frame, text='Начать', command=search)
lbl_warning = Label(frame, text='Все открытые окна Excel будут закрыты!', foreground='red3')
btn_copy = Button(frame, text='', relief=FLAT, command=copy_to_clipboard)
btn_copy['state'] = 'disabled'
txt_result = Text(frame, width=51, height=7, foreground='gray')
txt_result.insert(INSERT, chars='Здесь будет результат поиска')

scrollbar = Scrollbar(frame, orient=VERTICAL, command=txt_result.yview)
txt_result['yscrollcommand'] = scrollbar.set

lbl_file.grid(row=1, column=1, sticky=W)
btn_file.grid(row=1, column=2, sticky=W)
lbl_choosen_file.grid(row=1, column=2, sticky=W, padx=(87, 0))
lbl_list_num.grid(row=2, column=1, sticky=W)
ent_list_num.grid(row=2, column=2, pady=5, sticky=W)
lbl_list_num_tip.grid(row=2, column=2, sticky=W, padx=(97, 0))
lbl_search.grid(row=3, column=1, sticky=W)
ent_search.grid(row=3, column=2, sticky=W)
cvs_search.grid(row=4, column=2, sticky=E)
rad_exact.grid(row=4, column=1, columnspan=2, sticky=W)
rad_in.grid(row=4, column=1, columnspan=2, sticky=W, padx=(135, 0))
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
rad_color_off.grid(row=8, column=1, columnspan=2, sticky=W, padx=(360, 0))
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
btn_search.grid(row=11, column=1, columnspan=2, sticky=W, padx=(5, 0), pady=(15, 10))
lbl_warning.grid(row=11, column=1, columnspan=2, sticky=W, padx=(65, 0), pady=(15, 10))
btn_copy.grid(row=11, column=1, columnspan=2, sticky=SE, pady=(15, 0))
txt_result.grid(row=12, column=1, columnspan=2, sticky=W)
scrollbar.grid(row=12, column=1, columnspan=2, sticky=N+S, padx=(410, 0))

"""Вкладка поиска уникальных значений колонки"""


def choose_file_uq():
    global filepath_uq
    filepath_uq = filedialog.askopenfilename(
        filetypes=(('Книга Excel', '*.xls;*.xlsx;*.xlsm'), ('Все файлы', '*.*')))
    file_name = filepath_uq.split('/')[-1]
    file_name = file_name if len(file_name) < 26 else file_name[:23] + '...'
    if file_name:
        lbl_choosen_file_uq['text'] = file_name
        lbl_choosen_file_uq['foreground'] = 'black'
    elif not file_name and lbl_choosen_file_uq['foreground'] != 'gray':
        lbl_choosen_file_uq['text'] = 'файл не выбран'
        lbl_choosen_file_uq['foreground'] = 'gray'


def check_fields_uq(path, list_num, rng):
    if not all((path, list_num, rng)):
        messagebox.showwarning(title='Предупреждение',
                               message='Не все обязательные поля заполнены.')
    elif any(map(lambda x: x.lower() in cyrillic_str_lower, rng)):
        messagebox.showwarning(title='Предупреждение',
                               message='В поле с указанием колонки '
                                       'присутствуют кириллические символы.')
    else:
        return True


def search_uq():
    if check_fields_uq(filepath_uq, ent_list_num_uq.get(), ent_range_uq.get()):
        res, reg = [], r'([A-Z]+)(\d+):\1(\d+)'
        letter_range = match(r'\b[A-Z]+\b', ent_range_uq.get())
        complex_range = match(reg, ent_range_uq.get())
        if not (letter_range or complex_range):
            messagebox.showwarning(title='Предупреждение',
                                   message='Некорректно задана колонка поиска.')
        elif complex_range and int(
                match(reg, ent_range_uq.get()).group(2)) > int(
                match(reg, ent_range_uq.get()).group(3)):
            messagebox.showwarning(title='Предупреждение',
                                   message='Некорректно задан интервал поиска.')
        else:
            final_range = letter_range.group() if letter_range else complex_range.group()
            wb = Book(filepath_uq)
            try:
                if ent_list_num_uq.get().isdigit():
                    sht = wb.sheets[int(ent_list_num_uq.get()) - 1]
                else:
                    sht = wb.sheets[ent_list_num_uq.get()]
                if letter_range:
                    file_length = len(read_excel(filepath_uq)) + 3
                span = final_range if ':' in final_range else f'{letter_range.group()}1:{letter_range.group()}{file_length}'
                for r in sht.range(span):
                    if (s := r.value) and s not in res:
                        res.append(s)
                wb.app.quit()
                res = sorted(res) if is_sorted.get() else res
                final_res = '\n'.join(res)
                if final_res:
                    txt_result_uq.delete('1.0', END)
                    txt_result_uq.insert(INSERT, chars=final_res)
                    btn_copy_uq['text'] = 'скопировать в буфер'
                    btn_copy_uq['state'] = 'normal'
                else:
                    txt_result_uq.delete('1.0', END)
                    txt_result_uq.insert(INSERT, chars='Совпадения не найдены.')
                    btn_copy_uq['text'] = ''
                    btn_copy_uq['state'] = 'disabled'
            except Exception as err:
                print(err)
                wb.app.quit()
                messagebox.showwarning(title='Ошибка',
                                       message='Что-то пошло не так...')


def copy_to_clipboard_uq():
    window.clipboard_clear()
    a_uq = txt_result_uq.get('1.0', END)[:-1]
    window.clipboard_append(a_uq)


filepath_uq = ''
is_sorted = IntVar(value=1)

cvs_range_uq = Canvas(frame_uq, height=17, width=17)
tip_range_uq = cvs_range_uq.create_image(2, 2, anchor=NW, image=tip)

lbl_file_uq = Label(frame_uq, text='Выберите файл Excel       ')
btn_file_uq = Button(frame_uq, text='Выбрать', command=choose_file_uq, width=10)
lbl_choosen_file_uq = Label(frame_uq, text='файл не выбран', foreground='gray')
lbl_list_num_uq = Label(frame_uq, text='Номер листа в книге  ',)
ent_list_num_uq = Entry(frame_uq, width=15)
lbl_list_num_tip_uq = Label(frame_uq, text='1 - это Лист 1 или название листа', foreground='gray')
ent_list_num_uq.insert(END, '1')
lbl_range_uq = Label(frame_uq, text='Колонка поиска  ')
ent_range_uq = Entry(frame_uq, width=15)
Hovertip(cvs_range_uq, 'Буква колонки или диапазон ячеек в колонке латинскими\n'
                    'заглавными буквами, например B1:B176')
lbl_sorted_uq = Label(frame_uq, text='Отсортировать результат по алфавиту?  ')
rad_sorted_uq = Radiobutton(frame_uq, text='Да', value=1, variable=is_sorted)
rad_no_sorted_uq = Radiobutton(frame_uq, text='Нет', value=0, variable=is_sorted)
btn_search_uq = Button(frame_uq, text='Начать', command=search_uq)
lbl_warning_uq = Label(frame_uq, text='Все открытые окна Excel будут закрыты!', foreground='red3')
btn_copy_uq = Button(frame_uq, text='', relief=FLAT, command=copy_to_clipboard_uq)
btn_copy_uq['state'] = 'disabled'
txt_result_uq = Text(frame_uq, width=51, height=17, foreground='gray')
txt_result_uq.insert(INSERT, chars='Здесь будет результат поиска')

scrollbar_uq = Scrollbar(frame_uq, orient=VERTICAL, command=txt_result_uq.yview)
txt_result_uq['yscrollcommand'] = scrollbar_uq.set

lbl_file_uq.grid(row=1, column=1, sticky=W)
btn_file_uq.grid(row=1, column=2, sticky=W)
lbl_choosen_file_uq.grid(row=1, column=2, sticky=W, padx=(87, 0))
lbl_list_num_uq.grid(row=2, column=1, sticky=W)
ent_list_num_uq.grid(row=2, column=2, pady=5, sticky=W)
lbl_list_num_tip_uq.grid(row=2, column=2, sticky=W, padx=(97, 0))
lbl_range_uq.grid(row=5, column=1, sticky=W)
ent_range_uq.grid(row=5, column=2, sticky=W)
cvs_range_uq.grid(row=5, column=2, sticky=W, padx=(105, 0))
lbl_sorted_uq.grid(row=9, column=1, columnspan=2, sticky=W, pady=5)
rad_sorted_uq.grid(row=9, column=1, columnspan=2, sticky=W, padx=(225, 0))
rad_no_sorted_uq.grid(row=9, column=1, columnspan=2, sticky=W, padx=(275, 0))
btn_search_uq.grid(row=11, column=1, columnspan=2, sticky=W, padx=(5, 0), pady=(10, 10))
lbl_warning_uq.grid(row=11, column=1, columnspan=2, sticky=W, padx=(65, 0), pady=(15, 10))
btn_copy_uq.grid(row=11, column=1, columnspan=2, sticky=SE, pady=(15, 0))
txt_result_uq.grid(row=12, column=1, columnspan=2, sticky=W)
scrollbar_uq.grid(row=12, column=1, columnspan=2, sticky=N+S, padx=(410, 0))

window.mainloop()
