import os
import PySimpleGUI as sg

import config
import registry
import xls_w
from config import Config


class GUI:
    def __init__(self):
        self.version = 'v2.6.2'
        self.config = Config()

    def main_menu(self):
        sg.theme('LightBrown13')
        settings = self.config.load()
        file_path = settings['path']['file']
        dir_path = settings['path']['directory']

        # ------ Menu Definition ------ #
        menu_def = [
            ['Меню', ['Настройки', 'Выход']],
            ['Помощь', ['О приложении']],
        ]
        # ----------------------------- #
        autoselection_is_true = settings['path']['autoselection']
        section_is_visible = False if settings['path']['autoselection'] else True
        section_is_disable = settings['path']['autoselection']
        registry_section_size = (638, 56) if settings['path']['autoselection'] else (638, 116)

        layout = [
            [
                sg.Menu(menu_def, tearoff=False)
            ],
            [
                [
                    sg.Frame(layout=[
                        [
                            sg.Checkbox(text='Выбрать автоматически активный лист открытого файла Excel',
                                        default=autoselection_is_true, key='AUTOSELECTION', enable_events=True)
                        ],
                        [

                        ],
                        [
                            sg.pin(sg.Text(text='Путь к файлу реестра:', key='FILE_TEXT', visible=section_is_visible, size=19)),
                            sg.pin(sg.InputText(key='FILE', readonly=True, visible=section_is_visible, disabled=section_is_disable, enable_events=True, size=56)),
                            sg.pin(sg.FileBrowse(button_text='Обзор', target='FILE', initial_folder=file_path, key='FILE_BROWSE', visible=section_is_visible, disabled=section_is_disable)),
                        ],
                        [
                            sg.pin(sg.Text('Название листа в файле:', key='SHEET_TEXT', visible=section_is_visible, size=19)),
                            # sg.pin(sg.InputText(key='SHEET', visible=section_is_visible, disabled=section_is_disable, size=56)),
                            sg.pin(sg.DropDown(values=[], key='SHEETS', readonly=True, visible=section_is_visible, disabled=section_is_disable, size=56)),
                        ],
                    ], title='Файл реестра', relief=sg.RELIEF_SUNKEN, size=registry_section_size, key='REGISTRY_SECTION'),
                ],
            ],
            [
                [
                    sg.Frame(layout=[
                        [
                            sg.Text('Путь к папке со сканами:', visible=True, size=19),
                            sg.InputText(key='FOLDER', size=56),
                            sg.FolderBrowse(button_text='Обзор', target='FOLDER', initial_folder=dir_path)
                        ],
                    ], title='Папка со сканами', relief=sg.RELIEF_SUNKEN, size=registry_section_size),
                ],
            ],
            [
                sg.Output(size=(88, 20))
            ],
            [
                sg.Submit(button_text='Запуск', key='START'),
                sg.Submit(button_text='Настройки', key='SETTINGS'),
                sg.Cancel(button_text='Выход', key='EXIT'),
            ]
        ]

        window_main = sg.Window(f'Hyperlinks creator {self.version}', layout)

        while True:  # The Event Loop
            event, values = window_main.read()
            print('event', event)
            print('values', values)

            if event in ['EXIT', 'Выход', sg.WINDOW_CLOSED]:
                window_main.close()
                break

            if values['AUTOSELECTION'] is False and str(values['FILE'])[values['FILE'].rfind('.') + 1:] == 'xlsx':
                ws_list = xls_w.get_all_ws(values['FILE'])
                window_main['SHEETS'].update(values=ws_list)
                pass


            if values['AUTOSELECTION'] is True:
                window_main['FILE_TEXT'].update(visible=False)
                window_main['FILE'].update(visible=False)
                window_main['FILE_BROWSE'].update(disabled=True, visible=False)
                window_main['SHEET_TEXT'].update(visible=False)
                window_main['SHEETS'].update(value='', disabled=True, visible=False)
                registry_section_size = (638, 56)
                window_main['REGISTRY_SECTION'].set_size(registry_section_size)
            else:
                registry_section_size = (638, 116)
                window_main['REGISTRY_SECTION'].set_size(registry_section_size)
                window_main['FILE_TEXT'].update(visible=True)
                window_main['FILE'].update(visible=True)
                window_main['FILE_BROWSE'].update(disabled=False, visible=True)
                window_main['SHEET_TEXT'].update(visible=True)
                window_main['SHEETS'].update(disabled=False, visible=True)

            if event in ['SETTINGS', 'Настройки']:
                self.settings_menu()

            elif event in 'Start':
                file_path = values['file']
                if file_path:
                    file_path = file_path[:file_path.rfind('/')]
                    self.config.save({'file': {'path': file_path}})

                dir_path = values['folder']
                if dir_path:
                    self.config.save({'directory': {'scan': {'path': dir_path}}})

                registry_path = values['file']

                if not values['AUTOSELECTION']:
                    ws_name = values['SHEET']
                else:
                    ws_name = True

                dir_scan = values['folder']
                if os.path.exists(dir_scan):
                    print(f'Доступность каталога сканов проверена.')
                else:
                    print(f'Каталог сканов недоступен.')
                    break

                registry.body(registry_path, dir_scan, ws_name, self.config.load())

    @staticmethod
    def font_style_text(settings):
        settings = settings['font']['style']
        if settings['bold'] and settings['italic']:
            return 'Полужирный Курсив'
        elif not settings['bold'] and settings['italic']:
            return 'Курсив'
        elif settings['bold'] and not settings['italic']:
            return 'Полужирный'
        else:
            return 'Обычный'

    @staticmethod
    def underline_style_text(settings):
        settings = settings['font']['style']
        if settings['underline'] == 'single':
            return 'Одинарное'
        elif settings['underline'] == 'double':
            return 'Двойное'
        else:
            return '(Нет)'

    def settings_menu(self):
        settings = self.config.load()
        font_name = settings['font']['name']
        font_style = GUI.font_style_text(settings)
        font_size = settings['font']['size']
        color = settings['font']['color']
        underline_style = GUI.underline_style_text(settings)
        header_string_count = settings['file']['header_string_count']
        font_name_list = config.lists['font_name_list']
        font_style_list = config.lists['font_style_list']
        font_size_list = config.lists['font_size_list']
        underline_style_list = config.lists['underline_style_list']
        header_string_list = config.lists['header_string_list']

        column_registry_number = settings['file']['registry_column']['number']
        column_registry_number_is_disabled = False if settings['file']['registry_column']['enabled'] == 'number' else True
        column_registry_number_is_default = False if column_registry_number_is_disabled else True
        column_registry_letter = settings['file']['registry_column']['letter']
        column_registry_letter_is_disabled = False if settings['file']['registry_column']['enabled'] == 'letter' else True
        column_registry_letter_is_default = False if column_registry_letter_is_disabled else True
        column_registry_text = settings['file']['registry_column']['text']
        column_registry_text_is_disabled = False if settings['file']['registry_column']['enabled'] == 'text' else True
        column_registry_text_is_default = False if column_registry_text_is_disabled else True
        column_registry_number_list = config.lists['column_number_list']
        column_registry_letter_list = config.lists['column_letter_list']

        column_hyperlink_number = settings['file']['hyperlink_column']['number']
        column_hyperlink_number_is_disabled = False if settings['file']['hyperlink_column']['enabled'] == 'number' else True
        column_hyperlink_number_is_default = False if column_hyperlink_number_is_disabled else True
        column_hyperlink_letter = settings['file']['hyperlink_column']['letter']
        column_hyperlink_letter_is_disabled = False if settings['file']['hyperlink_column']['enabled'] == 'letter' else True
        column_hyperlink_letter_is_default = False if column_hyperlink_letter_is_disabled else True
        column_hyperlink_text = settings['file']['hyperlink_column']['text']
        column_hyperlink_text_is_disabled = False if settings['file']['hyperlink_column']['enabled'] == 'text' else True
        column_hyperlink_text_is_default = False if column_hyperlink_text_is_disabled else True
        column_hyperlink_number_list = config.lists['column_number_list']
        column_hyperlink_letter_list = config.lists['column_letter_list']

        left_col = sg.Column([
            [sg.Text(text='Шрифт', auto_size_text=True)],
            [sg.InputText(default_text=font_name, key='FONT_NAME', readonly=True, size=22)],
            [sg.Listbox(values=font_name_list, default_values=[font_name], key='FONT_NAME_LIST', enable_events=True, size=(20, 8))],
        ], size=(180, 220))

        mid_col = sg.Column([
            [sg.Text(text='Начертание', auto_size_text=True)],
            [sg.InputText(default_text=font_style, key='FONT_STYLE', readonly=True, size=22)],
            [sg.Listbox(values=font_style_list, default_values=[font_style], key='FONT_STYLE_LIST', enable_events=True, size=(20, 8))],
        ], size=(180, 220))

        right_col = sg.Column([
            [sg.Text(text='Размер', auto_size_text=True)],
            [sg.InputText(default_text=font_size, key='FONT_SIZE', readonly=True, size=8)],
            [sg.Listbox(values=font_size_list, default_values=[font_size], key='FONT_SIZE_LIST', enable_events=True, size=(6, 8))],
        ], size=(70, 220))

        tab_font_layout = [
            [left_col, mid_col, right_col],
            [
                sg.Text(text='  Подчеркивание:', auto_size_text=True),
                sg.DropDown(values=underline_style_list, default_value=underline_style, readonly=True, key='UNDERLINE_STYLE_LIST'),
                sg.Text(text='       ', auto_size_text=True),
                sg.Text(text='Цвет текста:', auto_size_text=True),
                sg.Button(button_text='', button_color=color, size=(2, 1), disabled=True, key='IMG_COLOR'),
                sg.Input(key='COLOR', readonly=True, size=(7, 1), enable_events=True, visible=False),
                sg.ColorChooserButton(button_text='Изменить', key='KEY_COLOR'),
            ],
        ]

        tab_file_layout = [
            [
                sg.Frame(layout=[
                    [
                        sg.Radio(text='Количество строк:',
                                 group_id='HEADER_STRING_COUNT_RADIO_GROUP',
                                 enable_events=False,
                                 key='HEADER_STRING_COUNT_RADIO',
                                 default=True,
                                 size=14),
                        sg.DropDown(values=header_string_list,
                                    default_value=header_string_count,
                                    readonly=True,
                                    key='HEADER_STRING_COUNT',
                                    size=3),
                    ]
                ], title='Количество строк в шапке таблицы', relief=sg.RELIEF_SUNKEN, size=(460, 50)),
            ],
            [
                sg.Frame(layout=[
                    [
                        sg.Radio(text='По номеру:',
                                 group_id='COLUMN_REGISTRY_NUMBER',
                                 enable_events=True,
                                 key='NUMBER_COLUMN_REGISTRY_NUMBER_RADIO',
                                 default=column_registry_number_is_default,
                                 size=14),
                        sg.DropDown(values=column_registry_number_list,
                                    default_value=column_registry_number,
                                    readonly=True,
                                    key='NUMBER_COLUMN_REGISTRY_NUMBER',
                                    disabled=column_registry_number_is_disabled,
                                    size=3),
                    ],
                    [
                        sg.Radio(text='По букве:',
                                 group_id='COLUMN_REGISTRY_NUMBER',
                                 enable_events=True, key='LETTER_COLUMN_REGISTRY_NUMBER_RADIO',
                                 default=column_registry_letter_is_default,
                                 size=14),
                        sg.DropDown(values=column_registry_letter_list,
                                    default_value=column_registry_letter,
                                    readonly=True,
                                    key='LETTER_COLUMN_REGISTRY_NUMBER',
                                    disabled=column_registry_letter_is_disabled,
                                    size=3),
                    ],
                    [
                        sg.Radio(text='По тексту в шапке:',
                                 group_id='COLUMN_REGISTRY_NUMBER',
                                 key='TEXT_COLUMN_REGISTRY_NUMBER_RADIO',
                                 enable_events=True,
                                 default=column_registry_text_is_default,
                                 size=14),
                        sg.InputText(default_text=column_registry_text,
                                     key='TEXT_COLUMN_REGISTRY_NUMBER',
                                     disabled=column_hyperlink_text_is_disabled,
                                     size=42),
                    ],
                ], title='Выбор столбца с регистрационными номерами', relief=sg.RELIEF_SUNKEN),
            ],
            [
                sg.Frame(layout=[
                    [
                        sg.Radio(text='По номеру:',
                                 group_id='COLUMN_HYPERLINK_NUMBER',
                                 enable_events=True,
                                 key='NUMBER_COLUMN_HYPERLINK_NUMBER_RADIO',
                                 default=column_hyperlink_number_is_default,
                                 size=14),
                        sg.DropDown(values=column_hyperlink_number_list,
                                    default_value=column_hyperlink_number,
                                    readonly=True,
                                    key='NUMBER_COLUMN_HYPERLINK_NUMBER',
                                    disabled=column_hyperlink_number_is_disabled,
                                    size=3),
                    ],
                    [
                        sg.Radio(text='По букве:',
                                 group_id='COLUMN_HYPERLINK_NUMBER',
                                 enable_events=True,
                                 key='LETTER_COLUMN_HYPERLINK_NUMBER_RADIO',
                                 default=column_hyperlink_letter_is_default,
                                 size=14),
                        sg.DropDown(values=column_hyperlink_letter_list,
                                    default_value=column_hyperlink_letter,
                                    readonly=True,
                                    key='LETTER_COLUMN_HYPERLINK_NUMBER',
                                    disabled=column_hyperlink_letter_is_disabled,
                                    size=3),
                    ],
                    [
                        sg.Radio(text='По тексту в шапке:',
                                 group_id='COLUMN_HYPERLINK_NUMBER',
                                 key='TEXT_COLUMN_HYPERLINK_NUMBER_RADIO',
                                 enable_events=True,
                                 default=column_hyperlink_text_is_default,
                                 size=14),
                        sg.InputText(default_text=column_hyperlink_text,
                                     key='TEXT_COLUMN_HYPERLINK_NUMBER',
                                     disabled=column_hyperlink_text_is_disabled,
                                     size=42),
                    ],
                ], title='Выбор столбца для гиперссылок', relief=sg.RELIEF_SUNKEN),
            ],
        ]

        layout = [
            [
                sg.TabGroup([
                    [
                        sg.Tab('Шрифт', tab_font_layout, key='TAB_FONT'),
                        sg.Tab('Файл', tab_file_layout, key='TAB_FILE'),
                    ]
                ], size=(472, 300), key='SETTINGS_TABS')
            ],
            [
                sg.Button(button_text='По умолчанию', key='SET_DEFAULT_SETTINGS', size=15),
                sg.Text(text='', size=27),
                sg.Button(button_text='Ок', key='SAVE_SETTINGS', size=4),
                sg.Button(button_text='Отмена', key='CANCEL_SETTINGS', size=6),
            ]
        ]

        window = sg.Window('My window with tabs', layout, default_element_size=(12, 1))

        while True:
            event, values = window.read()
            # print('event =', event)
            # print('values =', values)
            # ===(exit if)===
            if event in [sg.WIN_CLOSED, 'CANCEL_SETTINGS']:  # always,  always give a way out!
                window.close()
                break
            # ===(update color)===
            if values['COLOR'] in [None, 'None', '']:
                window['IMG_COLOR'].update(button_color='#0563c1')
            else:
                window['IMG_COLOR'].update(button_color=values['COLOR'])
            # ===(update font name)===
            if event in 'FONT_NAME_LIST':
                font_name = values['FONT_NAME_LIST']
                window['FONT_NAME'].update(value=font_name[0])
            # ===(update font size)===
            if event in 'FONT_SIZE_LIST':
                font_name = values['FONT_SIZE_LIST']
                window['FONT_SIZE'].update(value=font_name[0])
            # ===(update font style)===
            if event in 'FONT_STYLE_LIST':
                font_name = values['FONT_STYLE_LIST']
                window['FONT_STYLE'].update(value=font_name[0])
            # ===(set default settings)===
            if event in 'SET_DEFAULT_SETTINGS':
                # ===(set default variables)===
                def_settings = config.default_config
                # ===(tab font variables)===
                def_img_color = def_settings['font']['color']
                def_underline_style = GUI.underline_style_text(def_settings)
                def_font_name = def_settings['font']['name']
                def_index_font_name = config.lists['font_name_list'].index(def_font_name)
                def_font_size = def_settings['font']['size']
                def_index_font_size = config.lists['font_size_list'].index(def_font_size)
                def_font_style = GUI.font_style_text(def_settings)
                def_index_font_style = config.lists['font_style_list'].index(def_font_style)
                # ===(tab file variables)===
                def_header_string_count = def_settings['file']['header_string_count']

                def_column_registry_number_is_default = True if def_settings['file']['registry_column']['enabled'] == 'number' else False
                def_column_registry_number = def_settings['file']['registry_column']['number']
                def_column_registry_number_is_disabled = False if def_settings['file']['registry_column']['enabled'] == 'number' else True
                def_column_registry_letter_is_default = True if def_settings['file']['registry_column']['enabled'] == 'letter' else False
                def_column_registry_letter = def_settings['file']['registry_column']['letter']
                def_column_registry_letter_is_disabled = False if def_settings['file']['registry_column']['enabled'] == 'letter' else True
                def_column_registry_text_is_default = True if def_settings['file']['registry_column']['enabled'] == 'text' else False
                def_column_registry_text = def_settings['file']['registry_column']['text']
                def_column_registry_text_is_disabled = False if def_settings['file']['registry_column']['enabled'] == 'text' else True

                def_column_hyperlink_number_is_default = True if def_settings['file']['hyperlink_column']['enabled'] == 'number' else False
                def_column_hyperlink_number = def_settings['file']['hyperlink_column']['number']
                def_column_hyperlink_number_is_disabled = False if def_settings['file']['hyperlink_column']['enabled'] == 'number' else True
                def_column_hyperlink_letter_is_default = True if def_settings['file']['hyperlink_column']['enabled'] == 'letter' else False
                def_column_hyperlink_letter = def_settings['file']['hyperlink_column']['letter']
                def_column_hyperlink_letter_is_disabled = False if def_settings['file']['hyperlink_column']['enabled'] == 'letter' else True
                def_column_hyperlink_text_is_default = True if def_settings['file']['hyperlink_column']['enabled'] == 'text' else False
                def_column_hyperlink_text = def_settings['file']['hyperlink_column']['text']
                def_column_hyperlink_text_is_disabled = False if def_settings['file']['hyperlink_column']['enabled'] == 'text' else True
                # ===(set fields with default variables)===
                # ===(tab font)===
                if values['SETTINGS_TABS'] == 'TAB_FONT':
                    window['IMG_COLOR'].update(button_color=def_img_color)
                    window['COLOR'].update(value=def_img_color)
                    window['UNDERLINE_STYLE_LIST'].update(value=def_underline_style)
                    window['FONT_NAME'].update(value=def_font_name)
                    window['FONT_NAME_LIST'].update(set_to_index=def_index_font_name)
                    window['FONT_SIZE'].update(value=def_font_size)
                    window['FONT_SIZE_LIST'].update(set_to_index=def_index_font_size)
                    window['FONT_STYLE'].update(value=def_font_style)
                    window['FONT_STYLE_LIST'].update(set_to_index=def_index_font_style)
                # ===(tab file)===
                elif values['SETTINGS_TABS'] == 'TAB_FILE':
                    window['HEADER_STRING_COUNT'].update(value=def_header_string_count)
                    window['NUMBER_COLUMN_REGISTRY_NUMBER_RADIO'].update(value=def_column_registry_number_is_default)
                    window['NUMBER_COLUMN_REGISTRY_NUMBER'].update(value=def_column_registry_number)
                    window['NUMBER_COLUMN_REGISTRY_NUMBER'].update(disabled=def_column_registry_number_is_disabled)
                    window['LETTER_COLUMN_REGISTRY_NUMBER_RADIO'].update(value=def_column_registry_letter_is_default)
                    window['LETTER_COLUMN_REGISTRY_NUMBER'].update(value=def_column_registry_letter)
                    window['LETTER_COLUMN_REGISTRY_NUMBER'].update(disabled=def_column_registry_letter_is_disabled)
                    window['TEXT_COLUMN_REGISTRY_NUMBER_RADIO'].update(value=def_column_registry_text_is_default)
                    window['TEXT_COLUMN_REGISTRY_NUMBER'].update(value=def_column_registry_text)
                    window['TEXT_COLUMN_REGISTRY_NUMBER'].update(disabled=def_column_registry_text_is_disabled)
                    window['NUMBER_COLUMN_HYPERLINK_NUMBER_RADIO'].update(value=def_column_hyperlink_number_is_default)
                    window['NUMBER_COLUMN_HYPERLINK_NUMBER'].update(value=def_column_hyperlink_number)
                    window['NUMBER_COLUMN_HYPERLINK_NUMBER'].update(disabled=def_column_hyperlink_number_is_disabled)
                    window['LETTER_COLUMN_HYPERLINK_NUMBER_RADIO'].update(value=def_column_hyperlink_letter_is_default)
                    window['LETTER_COLUMN_HYPERLINK_NUMBER'].update(value=def_column_hyperlink_letter)
                    window['LETTER_COLUMN_HYPERLINK_NUMBER'].update(disabled=def_column_hyperlink_letter_is_disabled)
                    window['TEXT_COLUMN_HYPERLINK_NUMBER_RADIO'].update(value=def_column_hyperlink_text_is_default)
                    window['TEXT_COLUMN_HYPERLINK_NUMBER'].update(value=def_column_hyperlink_text)
                    window['TEXT_COLUMN_HYPERLINK_NUMBER'].update(disabled=def_column_hyperlink_text_is_disabled)
            # ===(set registry column)===
            if event in 'NUMBER_COLUMN_REGISTRY_NUMBER_RADIO':
                window['NUMBER_COLUMN_REGISTRY_NUMBER'].update(disabled=False)
                window['LETTER_COLUMN_REGISTRY_NUMBER'].update(disabled=True)
                window['TEXT_COLUMN_REGISTRY_NUMBER'].update(disabled=True)
            elif event in 'LETTER_COLUMN_REGISTRY_NUMBER_RADIO':
                window['NUMBER_COLUMN_REGISTRY_NUMBER'].update(disabled=True)
                window['LETTER_COLUMN_REGISTRY_NUMBER'].update(disabled=False)
                window['TEXT_COLUMN_REGISTRY_NUMBER'].update(disabled=True)
            elif event in 'TEXT_COLUMN_REGISTRY_NUMBER_RADIO':
                window['NUMBER_COLUMN_REGISTRY_NUMBER'].update(disabled=True)
                window['LETTER_COLUMN_REGISTRY_NUMBER'].update(disabled=True)
                window['TEXT_COLUMN_REGISTRY_NUMBER'].update(disabled=False)
            # ===(set hyperlink column)===
            if event in 'NUMBER_COLUMN_HYPERLINK_NUMBER_RADIO':
                window['NUMBER_COLUMN_HYPERLINK_NUMBER'].update(disabled=False)
                window['LETTER_COLUMN_HYPERLINK_NUMBER'].update(disabled=True)
                window['TEXT_COLUMN_HYPERLINK_NUMBER'].update(disabled=True)
            elif event in 'LETTER_COLUMN_HYPERLINK_NUMBER_RADIO':
                window['NUMBER_COLUMN_HYPERLINK_NUMBER'].update(disabled=True)
                window['LETTER_COLUMN_HYPERLINK_NUMBER'].update(disabled=False)
                window['TEXT_COLUMN_HYPERLINK_NUMBER'].update(disabled=True)
            elif event in 'TEXT_COLUMN_HYPERLINK_NUMBER_RADIO':
                window['NUMBER_COLUMN_HYPERLINK_NUMBER'].update(disabled=True)
                window['LETTER_COLUMN_HYPERLINK_NUMBER'].update(disabled=True)
                window['TEXT_COLUMN_HYPERLINK_NUMBER'].update(disabled=False)
            # ===(save settings)===
            if event in 'SAVE_SETTINGS':
                if values['LETTER_COLUMN_REGISTRY_NUMBER_RADIO']:
                    registry_column_enabled = 'letter'
                elif values['NUMBER_COLUMN_REGISTRY_NUMBER_RADIO']:
                    registry_column_enabled = 'number'
                elif values['NUMBER_COLUMN_REGISTRY_NUMBER_RADIO']:
                    registry_column_enabled = 'text'
                else:
                    registry_column_enabled = config.default_config['file']['registry_column']['enabled']

                if values['LETTER_COLUMN_HYPERLINK_NUMBER_RADIO']:
                    hyperlink_column_enabled = 'letter'
                elif values['NUMBER_COLUMN_HYPERLINK_NUMBER_RADIO']:
                    hyperlink_column_enabled = 'number'
                elif values['NUMBER_COLUMN_HYPERLINK_NUMBER_RADIO']:
                    hyperlink_column_enabled = 'text'
                else:
                    hyperlink_column_enabled = config.default_config['file']['hyperlink_column']['enabled']

                settings = {
                    'file': {
                        'header_string_count': values['HEADER_STRING_COUNT'],
                        'registry_column': {
                            'letter': values['LETTER_COLUMN_REGISTRY_NUMBER'],
                            'number': values['NUMBER_COLUMN_REGISTRY_NUMBER'],
                            'text': values['TEXT_COLUMN_REGISTRY_NUMBER'],
                            'enabled': registry_column_enabled
                        },
                        'hyperlink_column': {
                            'letter': values['LETTER_COLUMN_HYPERLINK_NUMBER'],
                            'number': values['NUMBER_COLUMN_HYPERLINK_NUMBER'],
                            'text': values['TEXT_COLUMN_HYPERLINK_NUMBER'],
                            'enabled': hyperlink_column_enabled
                        },
                    }
                }

                self.config.save(settings)
                window.close()
                break
        pass

    @staticmethod
    def progress_bar(size):
        # layout the window
        layout = [
            [sg.Text('Working...')],
            [sg.ProgressBar(size, orientation='h', size=(28, 20), key='PROGRESSBAR')]
        ]
        return sg.Window('Create hyperlinks', layout, disable_minimize=True, keep_on_top=True)


if __name__ == '__main__':
    gg = GUI()
    gg.main_menu()
    pass
