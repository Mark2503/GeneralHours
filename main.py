import re
import os

import pandas as pd

BASE_PATH: str = os.path.dirname(os.path.abspath(__file__))

DESKTOP = os.path.join(os.environ['PUBLIC'], 'Отчет')


# Конвертация времени в часы
class ConvertTimeClock:

    def __init__(self, path_file, new_name_file):

        self.path_file = path_file
        self.new_name_file = new_name_file

    # Чтения файла excel
    def read_excel(self):

        cols = [0, 1, 2]
        return pd.read_excel(self.path_file, sheet_name='Поиск в контроле доступа', usecols=cols)

    # Запись в файл excel
    def write_xlsx(self, data, new_path_file):

        df = pd.DataFrame(data)
        df.to_excel(new_path_file)
        os.startfile(new_path_file)

    # Сортировка колонок
    def sort_colum(self, data):

        colum = list(data.keys())

        names, times, access_point, original_names = list(), list(), list(), set()

        for key, value in data.items():

            for data_value in value.values():

                if key == colum[0]:
                    names.append(data_value)
                    original_names.add(data_value)

                if key == colum[1]:
                    times.append(data_value)

                if key == colum[2]:
                    access_point.append(data_value)

        return names, times, access_point, original_names

    # Сортировка
    def sort_data(self, sc_data):

        sd = {
            original_names: [f'{sc_data[2][i].split(" ")[0]} {sc_data[1][i]}'
                             for i in range(0, len(sc_data[0])) if original_names == sc_data[0][i]] for original_names
            in sc_data[3]
        }

        return sd

    # Сортировка вход выход
    def sort_inputs_outputs(self, data_lists):

        try:

            custom_data_list = ['Выход', 'Вход']

            inputs, outputs = list(), list()

            i = 0

            while i < len(data_lists) - 1:

                d = [data_lists[i:i + 2][0].split(" ")[0], data_lists[i:i + 2][1].split(" ")[0]]

                if d != custom_data_list:
                    del data_lists[i]
                else:
                    i += 2

            for dl in data_lists:
                dl = str(dl).split(' ')
                access_point, date, times = dl[0], dl[1], dl[2]

                if access_point == custom_data_list[0]:
                    outputs.append(self.convert_time_to_minutes(times))

                elif access_point == custom_data_list[1]:
                    inputs.append(self.convert_time_to_minutes(times))

            return (sum(outputs) - sum(inputs)) // 60

        except Exception as e:
            return e

    # Конвертация время в минуты
    def convert_time_to_minutes(self, time_str):

        hours, minutes, seconds = map(int, time_str.split(':'))

        total_minutes = hours * 60 + minutes + seconds / 60

        return total_minutes

    # Запуск скрипта
    def start(self):

        names, hours = list(), list()
        data = self.read_excel().to_dict()

        sc_data = self.sort_colum(data)

        sd = self.sort_data(sc_data)

        for key, value in sd.items():
            names.append(key)
            hours.append(self.sort_inputs_outputs(data_lists=value))

        self.write_xlsx(data={'Name': names, 'Hours': hours}, new_path_file=self.new_name_file)


# Добавляет все найденные xlsx файлы в базовом пути к проекту
def add_file_xlsx():

    try:

        return [f'{root}/{file}' for root, dirs, files in os.walk(BASE_PATH) for file in files if file.endswith('xlsx')]

    except Exception as e:
        print(f'{e}\n ->')


def create_folder(path):

    if not os.path.exists(path):
        os.makedirs(path)

    return path


def main():

    pattern_number = r"[0-9]+"
    pattern_text = r"[A-Za-z0-9]+"


    try:

        afx = add_file_xlsx()

        print(f'№ - Имя файла')
        for i in range(0, len(afx)):

            print(f'{i} - {os.path.split(afx[i])[1]}')

        while True:

            choose_file = input("Введите номер файла: ")

            if re.match(pattern_number, choose_file):
                choose_file = int(choose_file)

                try:
                    if afx[choose_file]:
                        print(afx[choose_file])
                        new_file_name = input('Введите название нового файла используйте латиницу: ')
                        if re.match(pattern_text, new_file_name):
                            new_file_name = os.path.join(create_folder(DESKTOP), f'{new_file_name}.xlsx')

                            ConvertTimeClock(path_file=afx[choose_file], new_name_file=new_file_name).start()
                            break

                        else:
                            print(f'Не корректное название файла => {new_file_name}')

                except Exception as e:
                    print(f'{e}\n -> Такого номера файла нету в списке => {choose_file}')

            else:
                print(f'Вы ввели не номер файла => {choose_file}')

    except Exception as e:
        print(f'{e}\n ->')


main()