import json
import os
from pathlib import Path


class Config:
    def __init__(self):
        self.local_dir_config_path = f'{str(Path.home())}\\HyperlinkCreator'
        self.local_file_config_path = f'{str(Path.home())}\\HyperlinkCreator\\config.json'
        self.default_config = default_config
        self.create_local_config()

    def create_local_config(self):
        is_dir_exist = os.path.exists(self.local_dir_config_path)
        if not is_dir_exist:
            os.mkdir(self.local_dir_config_path)
        is_dir_exist = os.path.exists(self.local_dir_config_path)
        is_file_exist = os.path.exists(self.local_file_config_path)
        if is_dir_exist and not is_file_exist:
            with open(self.local_file_config_path, 'w', encoding='cp1251') as file:
                json.dump(self.default_config, file, sort_keys=True, indent=2)

    def save(self, settings):
        is_file_exist = os.path.exists(self.local_file_config_path)
        config = self.load()
        if is_file_exist:
            with open(self.local_file_config_path, 'w', encoding='cp1251') as file:
                for key in list(settings):
                    config[key].update(settings[key])
                json.dump(config, file, sort_keys=True, indent=2)
        else:
            self.default_config = config

    def load(self):
        is_file_exist = os.path.exists(self.local_file_config_path)
        if is_file_exist:
            with open(self.local_file_config_path, 'r', encoding='cp1251') as file:
                local_config = json.load(file)
            return local_config
        else:
            return self.default_config

    def reset_config(self):
        is_file_exist = os.path.exists(self.local_file_config_path)
        if is_file_exist:
            os.remove(self.local_file_config_path)
        self.create_local_config()


lists = {
    'font_name_list': [
        'Arial',
        'Arial Black',
        'Comic Sans MS',
        'Courier New',
        'Georgia',
        'Impact',
        'Times New Roman',
        'Trebuchet MS',
        'Verdana'
    ],
    'column_number_list': list(number for number in range(1, 22)),
    'column_letter_list': list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'),
    'font_size_list': list(number for number in range(10, 21)),
    'header_string_list': list(number for number in range(1, 11)),
    'font_style_list': ['Обычный', 'Курсив', 'Полужирный', 'Полужирный Курсив'],
    'underline_style_list': ['(нет)', 'Одинарное', 'Двойное']
}


default_config = {

    'font': {
        'color': '#0563c1',
        'name': 'Times New Roman',
        'size': 12,
        'style': {
            'bold': False,
            'italic': False,
            'underline': 'single'
        },

    },
    'path': {
        'file': '',
        'directory': ''
    },
    'file': {
        'header_string_count': 1,
        'registry_column': {
            'letter': 'A',
            'number': 1,
            'text': 'Вх. №',
            'enabled': 'number'
        },
        'hyperlink_column': {
            'letter': 'M',
            'number': 13,
            'text': 'Скан образ документа',
            'enabled': 'text'
        },
    }
}


if __name__ == '__main__':
    pass
    # local_dir_config_path = f'{str(Path.home())}\\HyperlinkCreator'
    # local_file_config_path = f'{str(Path.home())}\\HyperlinkCreator\\config.json'
    # is_dir_exist = os.path.exists(local_dir_config_path)
    # is_file_exist = os.path.exists(local_file_config_path)
    # print(is_dir_exist and not is_file_exist)