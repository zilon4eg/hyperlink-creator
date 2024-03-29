from xls_w import Excel
import os
from GUI import GUI


def eng_to_rus_letters(text):
    return text.replace('a', 'а').replace('e', 'е').replace('o', 'о')\
        .replace('p', 'р').replace('c', 'с').replace('y', 'у').replace('x', 'х')


def find_column_by_text(xxl, text, header_string_count):
    text = ' '.join(str(text).strip().lower().split()) if ' ' in text else str(text).strip().lower()

    for string in range(1, header_string_count + 1):
        for count in range(1, xxl.size_string(header_string_count) + 1):
            cell_data = str(xxl.ws[f'{Excel.number_to_letter(count)}{string}'].value).strip().lower()
            if ' ' in cell_data:
                cell_data = ' '.join(cell_data.split())
            else:
                cell_data = cell_data

            if eng_to_rus_letters(cell_data) == eng_to_rus_letters(text):
                return count
    GUI.popup_error('Столбец с заданным текстом не найден!')
    return None


def data_analysis(xxl, files_dir, settings, dir_scan):
    # ===(Находим столбец с регистрационными номерами)===
    registry_column_type = settings['file']['registry_column']['enabled']
    if registry_column_type == 'text':
        text_for_find = settings['file']['registry_column'][registry_column_type]
        header_string_count = settings['file']['header_string_count']
        registry_column = find_column_by_text(xxl, text_for_find, header_string_count)
        registry_column = Excel.number_to_letter(registry_column)
    elif registry_column_type == 'number':
        registry_column = Excel.number_to_letter(settings['file']['registry_column'][registry_column_type])
    else:
        registry_column = settings['file']['registry_column'][registry_column_type]

    # ===(Находим столбец для гиперссылок)===
    hyperlink_column_type = settings['file']['hyperlink_column']['enabled']
    if hyperlink_column_type == 'text':
        text_for_find = settings['file']['hyperlink_column'][hyperlink_column_type]
        header_string_count = settings['file']['header_string_count']
        hyperlink_column = find_column_by_text(xxl, text_for_find, header_string_count)
        hyperlink_column = Excel.number_to_letter(hyperlink_column)
    elif hyperlink_column_type == 'number':
        hyperlink_column = Excel.number_to_letter(settings['file']['hyperlink_column'][hyperlink_column_type])
    else:
        hyperlink_column = settings['file']['hyperlink_column'][hyperlink_column_type]

    # ===(Номер первой строки таблицы после шапки)===
    first_string = settings['file']['header_string_count'] + 1

    # ===(initializing PROGRESSBAR)===
    pg_size = xxl.size_column(registry_column) - 1
    pg_window = GUI.progress_bar(pg_size)
    pg_window.read(timeout=10)
    pg_window.TKroot.focus_force()
    # ===(initializing PROGRESSBAR)===

    for position, string_number in enumerate(range(first_string, xxl.size_column(registry_column) + 1)):
        cell_data = str(xxl.ws[f'{registry_column}{string_number}'].value).strip()
        cell_data = cell_data[:cell_data.rfind('.')] if '.' in cell_data else cell_data

        if '/' in cell_data:
            cell_data = cell_data.replace(r'/', r'-')
            xxl.ws[f'{registry_column}{string_number}'].value = cell_data
        elif '_' in cell_data:
            cell_data = cell_data.replace(r'_', r'-')
            xxl.ws[f'{registry_column}{string_number}'].value = cell_data

        miss_file_count = 0
        for file in files_dir:
            file_type = file[:file.find('.') + 1] if file[:file.find('.') + 1].lower() in ['вх.', 'исх.'] else ''
            file_hl_name = file.split()[0] if ' ' in file else file[:file.rfind('.')]
            # file_extension = file[file.rfind('.') + 1:]
            file_name_for_find = file_hl_name[file_hl_name.rfind('.') + 1:] if '.' in file_hl_name else file_hl_name
            if '-' in file_name_for_find:
                if int(file_name_for_find[file_name_for_find.rfind('-') + 1:]) > 20:
                    file_name_for_find = file_name_for_find[:file_name_for_find.rfind('-')]

            hl_name = f'{file_type}{file_name_for_find}'
            if cell_data == file_name_for_find:
                xxl.create_hyperlinks(hl_name, file, f'{hyperlink_column}{string_number}')
            else:
                miss_file_count += 1
        print(f'Не найдено сопоставление регистрационному номеру {cell_data} среди файлов.') if miss_file_count == len(files_dir) else None
        # ===(increment PROGRESSBAR)===
        pg_window['PROGRESSBAR'].update_bar(position + 1)
    # ===(close PROGRESSBAR)===
    pg_window.close()


def body(file_path, dir_scan, ws_name, settings):
    xxl = Excel(file_path, dir_scan, ws_name, settings)
    files_dir = os.listdir(path=dir_scan)

    for file in files_dir:
        if ',' in file:
            old_name = f'{dir_scan}\\{file}'
            new_name = old_name.replace(',', '')
            os.rename(old_name, new_name)

    files_dir = os.listdir(path=dir_scan)

    print(f'Получен список файлов в папке {dir_scan}.')
    print('Анализ данных и формирование гиперссылок...')
    data_analysis(xxl, files_dir, settings, dir_scan)
    print('Гиперссылки сформированы.')
    print('Complete...' + '\n' * 1)


if __name__ == '__main__':
    pass
