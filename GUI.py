import os
import PySimpleGUI as sg

import config
import registry
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
            ['File', ['Exit']],
            ['Settings', ['Font', 'Hyperlink', 'Test']],
            ['Help'],
        ]
        # ----------------------------- #
        layout = [
            [
                sg.Menu(menu_def, tearoff=False)
            ],
            [
                sg.Text('Путь к файлу реестра: ', size=(17, 1)),
                sg.InputText(key='file', size=(58, 1)),
                sg.FileBrowse(target='file', initial_folder=file_path, size=(7, 1))
            ],
            [
                sg.Text('Название листа в книге Excel: ', size=(23, 1)),
                sg.InputText(key='SHEET', size=(32, 1), disabled=True),
                sg.Checkbox('Использовать активный лист', default=True, key='WS_CHECKBOX', enable_events=True)
            ],
            [
                sg.Text('Путь к папке со сканами: ', size=(19, 1)),
                sg.InputText(key='folder', size=(56, 1)),
                sg.FolderBrowse(target='folder', initial_folder=dir_path, size=(7, 1))
            ],
            [
                sg.Output(size=(88, 20))
            ],
            [
                sg.Submit(button_text='Start'),
                sg.Cancel(button_text='Exit'),
            ]
        ]

        window_main = sg.Window(f'Hyperlinks creator {self.version}', layout)

        while True:  # The Event Loop
            event, values = window_main.read()
            # print(event, values) #debug

            if event in (None, 'Exit', 'Cancel', sg.WINDOW_CLOSED):
                break

            if values['WS_CHECKBOX'] is True:
                window_main['SHEET'].update('', disabled=True)
            else:
                window_main['SHEET'].update(disabled=False)

            if event in 'Font':
                window_main.hide()
                self.font_menu()
                window_main.UnHide()

            elif event in 'Hyperlink':
                window_main.hide()
                self.color_chooser_menu()
                window_main.UnHide()

            elif event in 'Test':
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

                if not values['WS_CHECKBOX']:
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
        underline_style = GUI.underline_style_text(settings)
        font_name_list = config.lists['font_name_list']
        font_style_list = config.lists['font_style_list']
        font_size_list = config.lists['font_size_list']
        color = settings['font']['color']
        underline_style_list = config.lists['underline_style_list']

        left_col = sg.Column([
            [sg.Text(text='Шрифт', auto_size_text=True)],
            [sg.In(default_text=font_name, key='FONT_NAME_LIST', enable_events=True, readonly=True, size=22)],
            [sg.Listbox(values=font_name_list, default_values=[font_name], key='FONT_NAME_LIST', enable_events=True, size=(20, 8))],
        ], size=(180, 220))

        mid_col = sg.Column([
            [sg.Text(text='Начертание', auto_size_text=True)],
            [sg.In(default_text=font_style, key='FONT_STYLE', readonly=True, size=22)],
            [sg.Listbox(values=font_style_list, default_values=[font_style], key='FONT_STYLE_LIST', enable_events=True, size=(20, 8))],
        ], size=(180, 220))

        right_col = sg.Column([
            [sg.Text(text='Размер', auto_size_text=True)],
            [sg.In(default_text=font_size, key='FONT_SIZE', readonly=True, size=8)],
            [sg.Listbox(values=font_size_list, default_values=[font_size], key='FONT_SIZE_LIST', enable_events=True, size=(6, 8))],
        ], size=(70, 220))

        tab1_layout = [
            [left_col, mid_col, right_col],
            [
                sg.Text(text='  Подчеркивание:', auto_size_text=True),
                sg.DropDown(values=underline_style_list, default_value=underline_style, key='UNDERLINE_STYLE_LIST'),
                sg.Text(text='       ', auto_size_text=True),
                sg.Text(text='Цвет текста:', auto_size_text=True),
                sg.Button(button_text='', button_color=color, size=(2, 1), disabled=True, key='IMG_COLOR'),
                sg.Input(key='COLOR', readonly=True, size=(7, 1), enable_events=True, visible=True),
                sg.ColorChooserButton(button_text='Изменить', key='KEY_COLOR')
            ],
        ]

        tab2_layout = [
            [sg.T('This is inside tab 2')],
            [sg.In(key='in')]
        ]

        layout = [
            [
                sg.TabGroup(
                    [[
                        sg.Tab('Шрифт', tab1_layout),
                        sg.Tab('Гиперссылки', tab2_layout),
                    ]], size=(470, 300))
            ],
            [
                sg.Button(button_text='По умолчанию', key='SET_DEFAULT_SETTINGS'),
                sg.Button(button_text='Ок', key='SAVE_SETTINGS'),
                sg.Button(button_text='Отмена', key='CANCEL_SETTINGS'),
            ]
        ]

        window = sg.Window('My window with tabs', layout, default_element_size=(12, 1))

        while True:
            event, values = window.read()
            print(event)
            print(values)
            # ===(exit)===
            if event is sg.WIN_CLOSED:  # always,  always give a way out!
                break
            # ===(update color)===
            if values['COLOR'] in [None, 'None', '']:
                img_color = '#0563c1'
            else:
                img_color = values['COLOR']
            window['IMG_COLOR'].update(button_color=img_color)
            # ===(update font name)===
            if event in 'FONT_NAME_LIST':
                font_name = values['FONT_NAME']
                window['FONT_NAME'].update()

        pass

    def font_menu(self):
        settings = self.config.load()
        font_list = settings['font_list']
        font_name = settings['hyperlink']['font']['name']
        font_size = settings['hyperlink']['font']['size']

        layout = [
            [
                sg.Combo(font_list, default_value=font_name, key='drop-down', enable_events=True)
            ],
            [
                sg.Spin([sz for sz in range(10, 21)], font='Arial 20', initial_value=font_size,
                                 change_submits=True, key='spin'),
                sg.Slider(range=(10, 20), orientation='h', size=(10, 25), change_submits=True,
                                   key='slider', font=f'{font_name.replace(" ", "")} 20', default_value=font_size),
                sg.Text("Ab", size=(2, 1), font=f'{font_name.replace(" ", "")} {str(font_size)}', key='text')
            ],
            [
                sg.Submit(button_text='Ok'),
                sg.Cancel(button_text='Cancel')
            ]
        ]

        window = sg.Window("Font size selector", layout, grab_anywhere=False)
        # Event Loop

        while True:
            event, values = window.read()

            window['text'].update(font=f'{values["drop-down"].replace(" ", "")} {str(font_size)}')

            if event in (sg.WIN_CLOSED, 'Cancel'):
                window.close()
                break
            sz_spin = int(values['spin'])
            sz_slider = int(values['slider'])
            sz = sz_spin if sz_spin != font_size else sz_slider
            if sz != font_size:
                font_size = sz
                font = f'{values["drop-down"].replace(" ", "")} {str(font_size)}'
                window['text'].update(font=font)
                window['slider'].update(sz)
                window['spin'].update(sz)

            if event in 'Ok':
                self.config.save(
                    {
                        'hyperlink': {
                            'font': {
                                'size': int(font_size),
                                'name': values['drop-down']
                            }
                        }
                    }
                )
                window.close()
                break

    def color_chooser_menu(self):
        settings = self.config.load()
        img_color = settings['hyperlink']['font']['color']

        layout = [
            [
                sg.Text('Код цвета:'),
                sg.Input(key='COLOR', readonly=True, size=(7, 1), enable_events=True),
                sg.ColorChooserButton(button_text='Choose color', key='COLOR')
            ],
            [
                sg.Submit(button_text='Ok'),
                sg.Cancel(button_text='Cancel'),
                sg.Text(' Пример цвета:'),
                sg.Button(button_text='', button_color=img_color, size=(2, 1), disabled=True, key='IMG_COLOR'),
            ],
        ]

        window = sg.Window("Font size selector", layout, grab_anywhere=False)
        # Event Loop

        while True:
            event, values = window.read()

            if values['COLOR'] in [None, 'None', '']:
                img_color = '#0563c1'
            else:
                img_color = values['COLOR']
            window['IMG_COLOR'].update(button_color=img_color)

            if event in (sg.WIN_CLOSED, 'Cancel'):
                window.close()
                break

            if event in 'Ok':
                hyperlink_color = values['COLOR']
                if hyperlink_color in [None, 'None', '']:
                    hyperlink_color = '#0563c1'
                self.config.save({'hyperlink': {'font': {'color': hyperlink_color}}})
                window.close()
                break

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
