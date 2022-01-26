from xls_w import Excel
import os
from GUI import GUI
from natsort import natsorted, ns


def light_files_in_dir(list_files):
    list_files = list(str(file) for file in list_files if not str(file)[:str(file).rfind('.')].isdigit())
    list_files = list(map(lambda x: x[x.rfind('№') + 1:x.rfind('.')].lower(), list_files))
    return list_files


def miss_files(list1, list2):
    miss_list = list(set(list1) - set(list2))
    return natsorted(miss_list, alg=ns.IGNORECASE)  # or alg=ns.IC


def body(registry_path, dir_scan, ws_name, settings):
    xxl = Excel(registry_path, dir_scan, ws_name, settings)
    registry_path = xxl.get_path_active_book() if registry_path in '' else registry_path

    if registry_path[registry_path.rfind('\\') + 1:registry_path.rfind('.')].lower().count('исходящ') > 0:
        file_pref = 'исх.№'
        print('Загруженный документ идентифицирован как реестр Исходящих.')
    elif registry_path[registry_path.rfind('\\') + 1:registry_path.rfind('.')].lower().count('входящ') > 0:
        file_pref = 'вход.№'
        print('Загруженный документ идентифицирован как реестр Входящих.')
    else:
        file_pref = '№'
        print('Реестр не идентифицирован.')

    files_a = xxl.get_column()
    files_a_sort = list(map(lambda x: str(x).replace(r'/', r'-').strip().split()[0], files_a))
    files_dir = os.listdir(path=dir_scan)
    print(f'Получен список файлов в папке {dir_scan}.')

    files_dir_clear = light_files_in_dir(files_dir)
    miss_list = miss_files(files_a_sort, files_dir_clear)
    for miss in miss_list:
        print(f'Не найдено сопоставление регистрационному номеру {miss} среди файлов.')

    print('Формирование гиперссылок...')

    pg_size = len(files_a)
    pg_window = GUI.progress_bar(pg_size)
    pg_window.read(timeout=10)
    pg_window.TKroot.focus_force()

    for position, file_a in enumerate(files_a, 3):
        file_a_clear = file_a.replace(r'/', r'-').strip().split()[0]

        pg_window['PROGRESSBAR'].update_bar(position + 1)

        for file_dir in files_dir:
            if not file_dir.isdigit():
                file_type = file_dir[file_dir.rfind('.'):].lower()
                file_dir_clear = file_dir[file_dir.rfind('№') + 1:file_dir.rfind('.')].lower()

                if file_dir_clear == file_a_clear:
                    name = f'{file_pref.capitalize()}{file_a_clear}'
                    link_name = f'{file_pref.upper()}{file_a_clear}{file_type}'
                    if not xxl.check_hyperlink(name, link_name, position):
                        xxl.create_hyperlinks(name, link_name, position)

    pg_window.close()
    print('Гиперссылки сформированы.')
    print('Complete...' + '\n' * 1)


if __name__ == '__main__':
    pass
