import pandas as pd
import time
import tkinter as tk
import tkinter.filedialog as fd
from tqdm import tqdm

object = fd.askopenfilename()
row = int(input("Введите порядковый номер строки: ")) - 1            # Выбор строки в таблице
column = [int(input("Введите порядковый номер столбца: ")) - 1]      # Выбор столбца в таблице
excel_data = pd.read_excel(object, usecols=column, skiprows=row)     # Чтение определенного диапазона
excel_data = excel_data.dropna()                                     # Удаление значений NaN

with open('result.txt', 'w') as file:                                # Создает файл, преобразуя
    excel_data.to_string(file, index=False)                          # .xls в файл .txt
with open("result.txt", "r") as file_space:
    file_content: str = file_space.read()
file_content = file_content.replace(" ", "")                         # Первый символ в файле - пробел
with open("result.txt", "w") as file_space:                          # удаляем пробел
    file_space.write(file_content)

target_list = []
try:
    with open('result.txt', 'r') as source_file:
        for line in source_file.readlines():
            line = line.rstrip()
            if len(line) == 12:
                line = line[:3] + '-' + line[3:6] + '-' + line[6:] + '\n'
                target_list.append(line)
            elif len(line) < 12:
                line = '0' + line[:2] + '-' + line[2:5] + '-' + line[5:] + '\n'
                target_list.append(line)
            else:
                print('Incorrect format of the entries!')
except IOError:
    print("Can't open the source. No such file!")

with open('result.txt', 'w') as target_file:
    for line in tqdm(target_list):
        target_file.write(line)
        time.sleep(0.004)