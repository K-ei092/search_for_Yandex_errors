from tkinter import *
from tkinter import messagebox, ttk

import result_in_xlsx
from parser_mod import Parser
from powermanagement import long_running

import logging, os, time

# Инициализируем логгер
logger = logging.getLogger(__name__)


# декоратор для отключения спящего режима пока работает программа (для Windows 10, 11)
@long_running
def main():

    # Конфигурируем логирование
    logging.basicConfig(
        level=logging.WARNING,
        filename="logs.log",
        filemode='a',
        format='%(filename)s:%(lineno)d #%(levelname)-8s '
               '[%(asctime)s] - %(name)s - %(message)s')

    # Выводим в консоль информацию о начале запуска бота
    logger.info('Starting bot')

    list_words = txt.get().split(',')

    try:

        pars_ya = Parser()
        pars_ya.go_to_yandex()
        for i in range(len(list_words)):
            pars_ya.search_on_yandex(search_phrase=list_words[i])
        pars_ya.close_driver()

    except Exception:
        logger.exception()

    if result_in_xlsx.result_number == 1:
        os.remove(result_in_xlsx.name_file)
        messagebox.showinfo(message='ошибок не найдено')

    else:
        messagebox.showinfo(message=f'Ваш результат в файле {result_in_xlsx.name_file}\nНажмите "ОК" для открытия')
        time.sleep(2)
        os.system(f"start excel {result_in_xlsx.name_file}")


# Инициализируем Tkinter и его настройки
root = Tk()
root.title("Поиск ошибок Яндекса")
root.geometry('650x300')
lbl_1 = Label(
    root,
    text='Задайте поиск, например: авто в Москве, работа на Колыме (до 15 групп, обязательно разделенных запятой):'
)
lbl_1.grid(column=0, row=1)
txt = Entry(width=90)
txt.grid(column=0, row=4)
txt.focus()

lbl_2 = Label(
    root,
    text='После перехода на Яндекс поиск дождитесь начала скролинга либо решите капчу вручную'
)
lbl_2.grid(column=0, row=5)

lbl_3 = Label(
    root,
    text='далее можете свернуть окно Chrome, действия выполняются автоматически'
)
lbl_3.grid(column=0, row=6)

btn = ttk.Button(root, text='Начать', command=main)
btn.grid(column=0, row=7)


if __name__ == '__main__':
    root.mainloop()
